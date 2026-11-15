"""Archive immutable validation snapshots on a dedicated GitHub branch.

Credentials are supplied by the caller. This module never reads configuration
files or environment variables and never includes API response bodies in errors.
"""

import base64
import binascii
from dataclasses import dataclass
import json
import math
import re
from urllib.parse import quote

import requests


DEFAULT_MAX_JSON_BYTES = 20 * 1024 * 1024
DIRECTORY_LIMIT = 1000
_API_ROOT = "https://api.github.com"
_API_VERSION = "2022-11-28"
_SHA = re.compile(r"[0-9a-fA-F]{40,64}\Z")
_SNAPSHOT_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,127}\Z")
_FILENAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,230}\.json\Z")
_PATH_PART = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,99}\Z")


class GitHubHistoryError(Exception):
    """A safe, user-displayable error without credentials or snapshot data."""

    def __init__(self, message, code="response", status=None):
        super().__init__(message)
        self.code = code
        self.status = status


class HistoryConflictError(GitHubHistoryError):
    def __init__(self):
        super().__init__(
            "Ce nom existe déjà avec un contenu différent. Aucun fichier n'a été remplacé.",
            "conflict", 409,
        )


@dataclass(frozen=True)
class HistoryFile:
    name: str
    path: str
    sha: str
    size: int
    html_url: str


@dataclass(frozen=True)
class HistorySaveResult:
    name: str
    path: str
    created: bool
    html_url: str


def _configuration_error():
    return GitHubHistoryError("La configuration de l'historique GitHub est invalide.", "configuration")


def _response_error():
    return GitHubHistoryError("La réponse GitHub ne permet pas de lire cet historique.", "response")


def _validate_branch(branch):
    if not isinstance(branch, str) or not 1 <= len(branch) <= 240:
        raise _configuration_error()
    if not re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_./-]*", branch):
        raise _configuration_error()
    if ".." in branch or "//" in branch or branch.endswith(("/", ".")):
        raise _configuration_error()
    if any(part.startswith(".") or part.endswith(".lock") for part in branch.split("/")):
        raise _configuration_error()
    return branch


def _validate_filename(filename):
    if not isinstance(filename, str) or not _FILENAME.fullmatch(filename) or ".." in filename:
        raise GitHubHistoryError("Le nom du fichier d'historique est invalide.", "configuration")
    return filename


def _validate_json_value(value, depth=0):
    if depth > 50:
        raise ValueError("depth")
    if value is None or type(value) in (str, bool, int):
        return
    if type(value) is float and math.isfinite(value):
        return
    if type(value) is list:
        for item in value:
            _validate_json_value(item, depth + 1)
        return
    if type(value) is dict and all(type(key) is str for key in value):
        for item in value.values():
            _validate_json_value(item, depth + 1)
        return
    raise ValueError("json")


def _validate_document(payload, snapshot_id=None):
    try:
        _validate_json_value(payload)
        if type(payload) is not dict:
            raise ValueError("document")
        if type(payload.get("schemaVersion")) is not int or payload["schemaVersion"] < 1:
            raise ValueError("version")
        identifier = payload.get("snapshotID")
        if not isinstance(identifier, str) or not _SNAPSHOT_ID.fullmatch(identifier):
            raise ValueError("identifier")
        if snapshot_id is not None and identifier != snapshot_id:
            raise ValueError("identity")
    except (ValueError, RecursionError):
        raise GitHubHistoryError("Le document d'historique JSON est invalide.", "response") from None


def _canonical_bytes(payload):
    try:
        return (json.dumps(payload, ensure_ascii=False, sort_keys=True,
                           separators=(",", ":"), allow_nan=False) + "\n").encode("utf-8")
    except (TypeError, ValueError, UnicodeError, RecursionError):
        raise GitHubHistoryError("Le document d'historique JSON est invalide.", "response") from None


def _json_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate key")
        result[key] = value
    return result


def _invalid_json_constant(value):
    raise ValueError("constant")


class GitHubHistoryStore:
    """GitHub Contents API client; only ``save_snapshot`` performs writes.

    Existing paths are immutable: no request ever sends a file SHA for update.
    ``session`` is injectable so tests do not contact GitHub.
    """

    def __init__(self, repository, token=None, branch="validation-history",
                 directory="validation-history", session=None, timeout=(5, 30),
                 max_json_bytes=DEFAULT_MAX_JSON_BYTES):
        if not isinstance(repository, str) or repository.count("/") != 1:
            raise _configuration_error()
        owner, repo = repository.split("/")
        if not re.fullmatch(r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?", owner):
            raise _configuration_error()
        if not re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", repo) or repo in {".", ".."}:
            raise _configuration_error()
        if token is not None and (not isinstance(token, str) or "\n" in token or "\r" in token):
            raise _configuration_error()
        if not isinstance(directory, str) or not 1 <= len(directory) <= 240:
            raise _configuration_error()
        directory = directory.rstrip("/")
        if not directory or any(not _PATH_PART.fullmatch(part) or ".." in part
                                for part in directory.split("/")):
            raise _configuration_error()
        if type(max_json_bytes) is not int or not 1 <= max_json_bytes <= 100_000_000:
            raise _configuration_error()
        if (not isinstance(timeout, tuple) or len(timeout) != 2 or
                any(type(value) not in (int, float) or not math.isfinite(value) or value <= 0
                    for value in timeout)):
            raise _configuration_error()
        self.repository = repository
        self.branch = _validate_branch(branch)
        self.directory = directory
        self.max_json_bytes = max_json_bytes
        self.timeout = timeout
        self._session = session if session is not None else requests.Session()
        self._base = f"{_API_ROOT}/repos/{repository}"
        self._has_token = bool(token and token.strip())
        self._headers = {
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": _API_VERSION,
        }
        if self._has_token:
            self._headers["Authorization"] = "Bearer " + token.strip()

    def _request(self, method, endpoint="", expected=(200,), accept=None, **kwargs):
        headers = self._headers.copy()
        if accept:
            headers["Accept"] = accept
        try:
            response = self._session.request(
                method, self._base + endpoint, headers=headers,
                timeout=self.timeout, allow_redirects=False, **kwargs,
            )
        except requests.RequestException:
            raise GitHubHistoryError(
                "GitHub est momentanément inaccessible. Réessayez la sauvegarde ou la lecture.",
                "unavailable",
            ) from None
        status = response.status_code
        if status in expected:
            return response
        if status == 401:
            raise GitHubHistoryError("Le jeton GitHub n'est pas accepté.", "authentication", status)
        if status == 403:
            raise GitHubHistoryError(
                "GitHub refuse l'accès : vérifiez les droits du jeton ou la limite de requêtes.",
                "permission", status,
            )
        if status == 404:
            raise GitHubHistoryError("Le dépôt ou l'historique GitHub est introuvable.", "not_found", status)
        if status in (409, 422):
            raise HistoryConflictError()
        raise GitHubHistoryError("L'opération GitHub a échoué. Réessayez plus tard.", "unavailable", status)

    @staticmethod
    def _json(response):
        try:
            return response.json()
        except (TypeError, ValueError):
            raise _response_error() from None

    def _branch_ref(self, branch, missing_ok=False):
        response = self._request("GET", "/git/ref/heads/" + quote(branch, safe="/"),
                                 expected=(200, 404) if missing_ok else (200,))
        if response.status_code == 404:
            return None
        data = self._json(response)
        if (not isinstance(data, dict) or data.get("ref") != "refs/heads/" + branch or
                not isinstance(data.get("object"), dict) or
                data["object"].get("type") != "commit" or
                not isinstance(data["object"].get("sha"), str) or
                not _SHA.fullmatch(data["object"]["sha"])):
            raise _response_error()
        return data["object"]["sha"]

    def _repository_metadata(self):
        metadata = self._json(self._request("GET"))
        if not isinstance(metadata, dict) or not isinstance(metadata.get("default_branch"), str):
            raise _response_error()
        try:
            _validate_branch(metadata["default_branch"])
        except GitHubHistoryError:
            raise _response_error() from None
        return metadata

    def _branch_exists(self):
        if self._branch_ref(self.branch, missing_ok=True) is not None:
            return True
        # A 404 can hide an inaccessible private repository. Verify access before
        # presenting a missing history branch as an empty archive.
        self._repository_metadata()
        return False

    def _ensure_branch(self):
        if self._branch_ref(self.branch, missing_ok=True) is not None:
            return
        metadata = self._repository_metadata()
        default_branch = metadata["default_branch"]
        if default_branch == self.branch:
            raise GitHubHistoryError("Le dépôt GitHub doit contenir un premier commit.", "configuration")
        sha = self._branch_ref(default_branch)
        response = self._request("POST", "/git/refs", expected=(201, 409, 422), json={
            "ref": "refs/heads/" + self.branch, "sha": sha,
        })
        if response.status_code in (409, 422):
            # Another user may have initialized the branch concurrently.
            if self._branch_ref(self.branch, missing_ok=True) is None:
                raise HistoryConflictError()

    def _file_path(self, filename):
        return self.directory + "/" + _validate_filename(filename)

    def _file_url(self, path):
        return (f"https://github.com/{self.repository}/blob/" +
                quote(self.branch, safe="/") + "/" + quote(path, safe="/"))

    def _get_file(self, filename, missing_ok=False):
        path = self._file_path(filename)
        response = self._request("GET", "/contents/" + quote(path, safe="/"),
                                 expected=(200, 404) if missing_ok else (200,),
                                 params={"ref": self.branch}, accept="application/vnd.github.raw+json")
        if response.status_code == 404:
            return None
        # Raw media is supported above the 1 MB Contents/base64 threshold. Never
        # follow download_url: the only request target is the fixed API origin.
        response_bytes = getattr(response, "content", None)
        if isinstance(response_bytes, bytes):
            if len(response_bytes) > self.max_json_bytes:
                raise GitHubHistoryError("Ce fichier dépasse la taille maximale de l'historique.", "too_large")
            data = self._parse_document_bytes(response_bytes)
        else:
            data = self._json(response)
        if isinstance(data, dict) and "schemaVersion" in data:
            _validate_document(data)
            if len(_canonical_bytes(data)) > self.max_json_bytes:
                raise GitHubHistoryError("Ce fichier dépasse la taille maximale de l'historique.", "too_large")
            return data
        # Accept the normal metadata/base64 envelope if a compatible server
        # ignores the requested raw media type.
        if (not isinstance(data, dict) or data.get("type") != "file" or
                data.get("name") != filename or data.get("path") != path or
                data.get("submodule_git_url") or data.get("encoding") != "base64" or
                not isinstance(data.get("content"), str) or
                type(data.get("size")) is not int or data["size"] < 0):
            raise _response_error()
        content = data["content"].replace("\n", "").replace("\r", "")
        if data["size"] > self.max_json_bytes or len(content) > 4 * ((self.max_json_bytes + 2) // 3):
            raise GitHubHistoryError("Ce fichier dépasse la taille maximale de l'historique.", "too_large")
        try:
            raw = base64.b64decode(content, validate=True)
        except (ValueError, binascii.Error):
            raise _response_error() from None
        if len(raw) != data["size"]:
            raise _response_error()
        if len(raw) > self.max_json_bytes:
            raise GitHubHistoryError("Ce fichier dépasse la taille maximale de l'historique.", "too_large")
        payload = self._parse_document_bytes(raw)
        _validate_document(payload)
        return payload

    @staticmethod
    def _parse_document_bytes(raw):
        try:
            return json.loads(raw.decode("utf-8"), object_pairs_hook=_json_object,
                              parse_constant=_invalid_json_constant)
        except (ValueError, UnicodeError, RecursionError):
            raise _response_error() from None

    def list_snapshots(self):
        """List filenames; never create a branch or download snapshot contents."""
        if not self._branch_exists():
            return []
        response = self._request("GET", "/contents/" + quote(self.directory, safe="/"),
                                 expected=(200, 404), params={"ref": self.branch})
        if response.status_code == 404:
            return []
        entries = self._json(response)
        if not isinstance(entries, list):
            raise _response_error()
        if len(entries) >= DIRECTORY_LIMIT or 'rel="next"' in response.headers.get("Link", ""):
            raise GitHubHistoryError(
                "Le dossier d'historique atteint la limite de lecture GitHub. Configurez un nouveau dossier.",
                "directory_limit",
            )
        files = []
        for entry in entries:
            if not isinstance(entry, dict):
                raise _response_error()
            name = entry.get("name")
            if entry.get("type") != "file" or not isinstance(name, str) or not name.endswith(".json"):
                continue
            _validate_filename(name)
            path = self._file_path(name)
            if (entry.get("path") != path or entry.get("submodule_git_url") or
                    not isinstance(entry.get("sha"), str) or not _SHA.fullmatch(entry["sha"]) or
                    type(entry.get("size")) is not int or entry["size"] < 0):
                raise _response_error()
            files.append(HistoryFile(name, path, entry["sha"], entry["size"], self._file_url(path)))
        return sorted(files, key=lambda item: item.name, reverse=True)

    def read_snapshot(self, filename):
        """Read one strict UTF-8/base64 JSON file on the configured branch."""
        return self._get_file(_validate_filename(filename))

    def save_snapshot(self, filename, payload, snapshot_id):
        """Create one immutable file, or return an identical existing snapshot."""
        if not self._has_token:
            raise GitHubHistoryError(
                "Un jeton GitHub avec accès en écriture est nécessaire pour sauvegarder.",
                "authentication",
            )
        path = self._file_path(filename)
        if not isinstance(snapshot_id, str) or not _SNAPSHOT_ID.fullmatch(snapshot_id):
            raise GitHubHistoryError("L'identifiant de l'historique est invalide.", "configuration")
        _validate_document(payload, snapshot_id)
        raw = _canonical_bytes(payload)
        if len(raw) > self.max_json_bytes:
            raise GitHubHistoryError("Le document dépasse la taille maximale de l'historique.", "too_large")
        self._ensure_branch()
        existing = self._get_file(filename, missing_ok=True)
        if existing is not None:
            if _canonical_bytes(existing) != raw:
                raise HistoryConflictError()
            return HistorySaveResult(filename, path, False, self._file_url(path))
        # Omitting SHA is intentional: GitHub rejects an existing path instead
        # of replacing it, including a path created after our preceding GET.
        response = self._request("PUT", "/contents/" + quote(path, safe="/"),
                                 expected=(201, 409, 422), json={
                                     "message": "Archive validation snapshot " + snapshot_id,
                                     "content": base64.b64encode(raw).decode("ascii"),
                                     "branch": self.branch,
                                 })
        if response.status_code in (409, 422):
            existing = self._get_file(filename, missing_ok=True)
            if existing is None or _canonical_bytes(existing) != raw:
                raise HistoryConflictError()
            return HistorySaveResult(filename, path, False, self._file_url(path))
        return HistorySaveResult(filename, path, True, self._file_url(path))
