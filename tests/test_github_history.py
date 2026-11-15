import base64
import json
import unittest

import requests

from github_history import DEFAULT_MAX_JSON_BYTES, GitHubHistoryError, GitHubHistoryStore, HistoryConflictError


REPOSITORY = "ExampleOwner/example-repository"
FILENAME = "novembre_2026-11-01_2026-11-14_snapshot-123.json"
DIRECTORY = "validation-history"
BRANCH = "validation-history"
SHA = "a" * 40


def document(**changes):
    data = {"schemaVersion": 1, "snapshotID": "snapshot-123", "rows": [{"value": 0}],
            "label": "Règles de validation", "complete": True, "missing": None}
    data.update(changes)
    return data


def reference(branch=BRANCH):
    return {"ref": "refs/heads/" + branch, "object": {"type": "commit", "sha": SHA}}


def file_content(payload=None, raw=None, filename=FILENAME, **changes):
    if raw is None:
        raw = json.dumps(payload if payload is not None else document(), ensure_ascii=False).encode("utf-8")
    data = {"type": "file", "name": filename, "path": DIRECTORY + "/" + filename,
            "encoding": "base64", "content": base64.b64encode(raw).decode("ascii"),
            "size": len(raw), "sha": SHA}
    data.update(changes)
    return data


def entry(filename=FILENAME, **changes):
    data = {"type": "file", "name": filename, "path": DIRECTORY + "/" + filename,
            "size": 250, "sha": SHA}
    data.update(changes)
    return data


class FakeResponse:
    def __init__(self, status=200, payload=None, headers=None, raw=None):
        self.status_code = status
        self.payload = payload
        self.headers = headers or {}
        self.content = raw

    def json(self):
        if isinstance(self.payload, Exception):
            raise self.payload
        return self.payload


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if not self.responses:
            raise AssertionError("Unexpected request")
        response = self.responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response


class GitHubHistoryTests(unittest.TestCase):
    def store(self, responses=(), **kwargs):
        session = FakeSession(responses)
        store = GitHubHistoryStore(REPOSITORY, "dummy-token-for-tests", session=session, **kwargs)
        return store, session

    def assert_error_code(self, code, action):
        with self.assertRaises(GitHubHistoryError) as caught:
            action()
        self.assertEqual(caught.exception.code, code)
        return caught.exception

    def test_list_downloads_only_metadata_and_uses_dedicated_branch(self):
        earlier = "a_2026-10-01_2026-10-31_old.json"
        store, session = self.store([
            FakeResponse(payload=reference()),
            FakeResponse(payload=[entry(earlier), entry(), {"type": "dir", "name": "other"},
                                  {"type": "file", "name": "README.md"}]),
        ])
        result = store.list_snapshots()
        self.assertEqual([item.name for item in result], [FILENAME, earlier])
        self.assertEqual(result[0].path, DIRECTORY + "/" + FILENAME)
        self.assertEqual(result[0].sha, SHA)
        self.assertTrue(result[0].html_url.startswith("https://github.com/" + REPOSITORY + "/blob/" + BRANCH))
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET"])
        self.assertEqual(session.calls[1][2]["params"], {"ref": BRANCH})
        self.assertNotIn("content", session.calls[1][2])

    def test_missing_branch_is_empty_after_repository_access_is_confirmed_without_mutation(self):
        store, session = self.store([FakeResponse(404), FakeResponse(payload={"default_branch": "main"})])
        self.assertEqual(store.list_snapshots(), [])
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET"])

    def test_inaccessible_repository_is_not_reported_as_empty(self):
        store, session = self.store([FakeResponse(404), FakeResponse(404)])
        self.assert_error_code("not_found", store.list_snapshots)
        self.assertEqual(len(session.calls), 2)

    def test_missing_directory_is_empty_without_mutation(self):
        store, session = self.store([FakeResponse(payload=reference()), FakeResponse(404)])
        self.assertEqual(store.list_snapshots(), [])
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET"])

    def test_directory_limit_is_explicit_even_when_only_some_entries_are_json(self):
        store, _ = self.store([FakeResponse(payload=reference()),
                               FakeResponse(payload=[{"name": "README.md", "type": "file"}] * 1000)])
        self.assert_error_code("directory_limit", store.list_snapshots)

    def test_pagination_link_cannot_silently_truncate_history(self):
        store, _ = self.store([FakeResponse(payload=reference()),
                               FakeResponse(payload=[], headers={"Link": '<next>; rel="next"'})])
        self.assert_error_code("directory_limit", store.list_snapshots)

    def test_list_rejects_untrusted_paths_and_unexpected_response_shapes(self):
        for payload in [{"message": "sensitive"}, [entry(path="other/" + FILENAME)],
                        [entry(sha="not-a-sha")], [entry(size=True)],
                        [entry(submodule_git_url="https://example.invalid")], [None]]:
            with self.subTest(payload=payload):
                store, _ = self.store([FakeResponse(payload=reference()), FakeResponse(payload=payload)])
                self.assert_error_code("response", store.list_snapshots)

    def test_read_validates_utf8_base64_json_on_selected_branch(self):
        content = file_content()
        content["content"] = content["content"][:20] + "\n" + content["content"][20:] + "\n"
        store, session = self.store([FakeResponse(payload=content)])
        self.assertEqual(store.read_snapshot(FILENAME), document())
        call = session.calls[0]
        self.assertEqual(call[0], "GET")
        self.assertEqual(call[2]["params"], {"ref": BRANCH})
        self.assertEqual(call[2]["headers"]["Authorization"], "Bearer dummy-token-for-tests")
        self.assertEqual(call[2]["headers"]["Accept"], "application/vnd.github.raw+json")
        self.assertEqual(call[2]["timeout"], (5, 30))
        self.assertFalse(call[2]["allow_redirects"])

    def test_read_rejects_invalid_content_without_exposing_response(self):
        for data in [file_content(content="sensitive-secret-not-base64"),
                     file_content(encoding="none"), file_content(type="symlink"),
                     file_content(size=-1), file_content(size=1),
                     file_content(path="different/" + FILENAME), file_content(raw=b"\xff"),
                     file_content(raw=b'{"schemaVersion":1,"snapshotID":"sensitive-secret", "x": NaN}'),
                     file_content(raw=b'{"schemaVersion":1,"snapshotID":"x","snapshotID":"y"}'),
                     file_content(payload={"sensitive-secret": True})]:
            with self.subTest(data=data):
                store, _ = self.store([FakeResponse(payload=data)])
                error = self.assert_error_code("response", lambda: store.read_snapshot(FILENAME))
                self.assertNotIn("sensitive-secret", str(error))

    def test_read_limits_declared_and_encoded_size_before_decoding(self):
        for data in [file_content(size=901_000), file_content(content="A" * 1_200_004)]:
            store, _ = self.store([FakeResponse(payload=data)], max_json_bytes=900_000)
            self.assert_error_code("too_large", lambda: store.read_snapshot(FILENAME))

    def test_raw_content_above_one_megabyte_is_supported_without_download_url(self):
        payload = document(label="x" * 1_100_000)
        raw = json.dumps(payload).encode("utf-8")
        store, session = self.store([FakeResponse(raw=raw)])
        self.assertEqual(store.read_snapshot(FILENAME), payload)
        self.assertEqual(len(session.calls), 1)
        self.assertTrue(session.calls[0][1].startswith("https://api.github.com/repos/"))
        self.assertEqual(store.max_json_bytes, 20 * 1024 * 1024)

    def test_raw_size_is_limited_before_json_parsing(self):
        store, _ = self.store([FakeResponse(raw=b"not json" * 20)], max_json_bytes=100)
        self.assert_error_code("too_large", lambda: store.read_snapshot(FILENAME))

    def test_raw_json_rejects_invalid_utf8_constants_and_duplicate_keys(self):
        for raw in [b"\xff", b'{"schemaVersion":1,"snapshotID":"x","rows":[NaN]}',
                    b'{"schemaVersion":1,"snapshotID":"x","snapshotID":"y"}']:
            store, _ = self.store([FakeResponse(raw=raw)])
            self.assert_error_code("response", lambda: store.read_snapshot(FILENAME))

    def test_read_public_snapshots_without_authorization_header(self):
        session = FakeSession([FakeResponse(raw=json.dumps(document()).encode("utf-8"))])
        store = GitHubHistoryStore(REPOSITORY, session=session)
        self.assertEqual(store.read_snapshot(FILENAME), document())
        self.assertNotIn("Authorization", session.calls[0][2]["headers"])

    def test_public_listing_is_read_only_without_token(self):
        session = FakeSession([FakeResponse(payload=reference()), FakeResponse(payload=[entry()])])
        store = GitHubHistoryStore(REPOSITORY, token=None, session=session)
        self.assertEqual(len(store.list_snapshots()), 1)
        self.assertNotIn("Authorization", session.calls[0][2]["headers"])

    def test_save_without_token_is_rejected_before_network(self):
        for token in (None, "", " "):
            session = FakeSession([])
            store = GitHubHistoryStore(REPOSITORY, token=token, session=session)
            self.assert_error_code("authentication", lambda: store.save_snapshot(FILENAME, document(), "snapshot-123"))
            self.assertEqual(session.calls, [])

    def test_save_creates_without_sha_and_without_committing_on_main(self):
        store, session = self.store([FakeResponse(payload=reference()), FakeResponse(404), FakeResponse(201)])
        result = store.save_snapshot(FILENAME, document(), "snapshot-123")
        self.assertTrue(result.created)
        self.assertEqual(result.name, FILENAME)
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET", "PUT"])
        sent = session.calls[-1][2]["json"]
        self.assertEqual(sent["branch"], BRANCH)
        self.assertNotIn("sha", sent)
        self.assertEqual(json.loads(base64.b64decode(sent["content"])), document())
        self.assertNotIn("Règles", sent["message"])

    def test_save_initializes_branch_from_repository_default_only_when_requested(self):
        store, session = self.store([
            FakeResponse(404), FakeResponse(payload={"default_branch": "main"}),
            FakeResponse(payload=reference("main")), FakeResponse(201),
            FakeResponse(404), FakeResponse(201),
        ])
        result = store.save_snapshot(FILENAME, document(), "snapshot-123")
        self.assertTrue(result.created)
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET", "GET", "POST", "GET", "PUT"])
        self.assertEqual(session.calls[3][2]["json"], {"ref": "refs/heads/" + BRANCH, "sha": SHA})
        self.assertTrue(session.calls[2][1].endswith("/git/ref/heads/main"))

    def test_parallel_branch_initialization_is_rechecked(self):
        store, session = self.store([
            FakeResponse(404), FakeResponse(payload={"default_branch": "main"}),
            FakeResponse(payload=reference("main")), FakeResponse(422),
            FakeResponse(payload=reference()), FakeResponse(payload=file_content()),
        ])
        self.assertFalse(store.save_snapshot(FILENAME, document(), "snapshot-123").created)
        self.assertNotIn("PUT", [call[0] for call in session.calls])

    def test_failed_branch_initialization_cannot_fall_through_to_a_file_write(self):
        store, session = self.store([
            FakeResponse(404), FakeResponse(payload={"default_branch": "main"}),
            FakeResponse(payload=reference("main")), FakeResponse(422), FakeResponse(404),
        ])
        with self.assertRaises(HistoryConflictError):
            store.save_snapshot(FILENAME, document(), "snapshot-123")
        self.assertNotIn("PUT", [call[0] for call in session.calls])

    def test_identical_existing_snapshot_is_idempotent_without_put(self):
        # A JSON file with different whitespace/key order remains the same snapshot.
        store, session = self.store([FakeResponse(payload=reference()), FakeResponse(payload=file_content())])
        self.assertFalse(store.save_snapshot(FILENAME, document(), "snapshot-123").created)
        self.assertEqual([call[0] for call in session.calls], ["GET", "GET"])

    def test_existing_snapshot_is_never_overwritten_even_if_id_matches(self):
        store, session = self.store([FakeResponse(payload=reference()),
                                    FakeResponse(payload=file_content(payload=document(rows=[])))])
        with self.assertRaises(HistoryConflictError):
            store.save_snapshot(FILENAME, document(), "snapshot-123")
        self.assertNotIn("PUT", [call[0] for call in session.calls])

    def test_parallel_file_creation_is_idempotent_for_identical_content(self):
        for status in (409, 422):
            with self.subTest(status=status):
                store, session = self.store([FakeResponse(payload=reference()), FakeResponse(404),
                                            FakeResponse(status), FakeResponse(payload=file_content())])
                self.assertFalse(store.save_snapshot(FILENAME, document(), "snapshot-123").created)
                self.assertEqual([call[0] for call in session.calls], ["GET", "GET", "PUT", "GET"])
                self.assertNotIn("sha", session.calls[2][2]["json"])

    def test_parallel_file_creation_preserves_other_content(self):
        store, session = self.store([FakeResponse(payload=reference()), FakeResponse(404),
                                    FakeResponse(409), FakeResponse(payload=file_content(payload=document(rows=[])))])
        with self.assertRaises(HistoryConflictError):
            store.save_snapshot(FILENAME, document(), "snapshot-123")
        self.assertEqual(sum(call[0] == "PUT" for call in session.calls), 1)

    def test_invalid_file_names_are_rejected_before_network(self):
        for filename in ["../file.json", "folder/file.json", "file\\name.json", "file.json?x=1",
                         "/file.json", "..json", "file..json", "file.txt", "", "é.json", None]:
            with self.subTest(filename=filename):
                store, session = self.store()
                self.assert_error_code("configuration", lambda: store.read_snapshot(filename))
                self.assert_error_code("configuration", lambda: store.save_snapshot(filename, document(), "snapshot-123"))
                self.assertEqual(session.calls, [])

    def test_invalid_payloads_are_rejected_before_network_or_branch_creation(self):
        for payload, identifier in [([], "snapshot-123"), ({}, "snapshot-123"),
                                    (document(schemaVersion=True), "snapshot-123"),
                                    (document(snapshotID="different"), "snapshot-123"),
                                    (document(), "invalid/id"), (document(rows=[float("nan")]), "snapshot-123"),
                                    (document(rows=[float("inf")]), "snapshot-123"),
                                    (document(rows=[b"secret"]), "snapshot-123"),
                                    (document(rows={1: "non-string key"}), "snapshot-123")]:
            with self.subTest(payload=payload, identifier=identifier):
                store, session = self.store()
                with self.assertRaises(GitHubHistoryError):
                    store.save_snapshot(FILENAME, payload, identifier)
                self.assertEqual(session.calls, [])

    def test_size_is_utf8_bytes_and_checked_before_network(self):
        store, session = self.store(max_json_bytes=250)
        self.assert_error_code("too_large", lambda: store.save_snapshot(
            FILENAME, document(label="é" * 100), "snapshot-123"))
        self.assertEqual(session.calls, [])

    def test_cycles_and_excessive_nesting_are_safe_errors(self):
        cyclic = document()
        cyclic["cycle"] = cyclic
        nested = document()
        child = nested
        for _ in range(60):
            child["nested"] = {}
            child = child["nested"]
        for payload in (cyclic, nested):
            store, session = self.store()
            self.assert_error_code("response", lambda: store.save_snapshot(FILENAME, payload, "snapshot-123"))
            self.assertEqual(session.calls, [])

    def test_safe_configuration_accepts_nested_branches_and_directories(self):
        store, session = self.store([FakeResponse(404)], branch="archive/validation-2026",
                                    directory="archives/validation/")
        self.assertEqual(store.directory, "archives/validation")
        self.assert_error_code("not_found", lambda: store.read_snapshot(FILENAME))
        self.assertIn("/contents/archives/validation/", session.calls[0][1])
        self.assertEqual(session.calls[0][2]["params"], {"ref": "archive/validation-2026"})

    def test_invalid_configuration_is_rejected(self):
        for changes in [{"repository": "https://github.com/owner/repo"}, {"repository": "owner/repo/extra"},
                        {"repository": "../repo"}, {"repository": "owner/.."},
                        {"branch": "../main"}, {"branch": "refs//heads/main"}, {"branch": "main.lock"},
                        {"branch": "main?token=secret"}, {"branch": "main/.hidden"},
                        {"directory": "../history"}, {"directory": "/history"}, {"directory": "a\\b"},
                        {"directory": "a//b"}, {"directory": ""}, {"token": 123},
                        {"token": "dummy\nheader"}, {"max_json_bytes": 100_000_001},
                        {"max_json_bytes": True}, {"timeout": (5, 0)}, {"timeout": (5, float("inf"))}]:
            with self.subTest(changes=changes):
                config = {"repository": REPOSITORY, "token": "dummy-token-for-tests", "session": FakeSession([])}
                config.update(changes)
                self.assert_error_code("configuration", lambda: GitHubHistoryStore(**config))

    def test_network_and_http_errors_are_sanitized(self):
        cases = [(requests.Timeout("sensitive-token request-body"), "unavailable"),
                 (FakeResponse(401, {"message": "sensitive-token"}), "authentication"),
                 (FakeResponse(403, {"message": "sensitive-token"}), "permission"),
                 (FakeResponse(404, {"message": "sensitive-token"}), "not_found"),
                 (FakeResponse(429, {"message": "sensitive-token"}), "unavailable"),
                 (FakeResponse(503, {"message": "sensitive-token"}), "unavailable"),
                 (FakeResponse(302), "unavailable"),
                 (FakeResponse(payload=ValueError("sensitive-token")), "response")]
        for response, code in cases:
            with self.subTest(code=code):
                store, _ = self.store([response])
                error = self.assert_error_code(code, lambda: store.read_snapshot(FILENAME))
                self.assertNotIn("sensitive-token", str(error))
                self.assertIsNone(error.__cause__)

    def test_wrong_ref_response_cannot_be_used_as_history_branch(self):
        for payload in [reference("main"), [reference()], {"ref": "refs/heads/" + BRANCH,
                                                         "object": {"type": "tag", "sha": SHA}}]:
            store, session = self.store([FakeResponse(payload=payload)])
            self.assert_error_code("response", store.list_snapshots)
            self.assertEqual(len(session.calls), 1)


if __name__ == "__main__":
    unittest.main()
