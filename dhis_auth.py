"""Vérification des comptes DHIS2 avec délais distincts et erreurs lisibles."""

import math

import requests


def _positive_seconds(value, default):
    try:
        number = float(value)
        return number if math.isfinite(number) and number > 0 else default
    except (TypeError, ValueError, OverflowError):
        return default


def check_dhis2_credentials(base_url, username, password, timeout_connect=10,
                           timeout_read=60, retries=1, request_get=None):
    """Teste /api/me, avec au plus une relance des seuls incidents transitoires.

    Les identifiants sont transmis à Requests sans être inclus dans les messages
    retournés. Les réponses du serveur et les textes d'exception peuvent contenir
    des informations sensibles ; ils ne sont jamais affichés à l'utilisateur.
    """
    if not isinstance(base_url, str) or not base_url.strip():
        return False, "L’adresse du serveur DHIS2 est manquante."
    if not username or not password:
        return False, "Renseignez votre nom d’utilisateur et votre mot de passe DHIS2."
    connect_seconds = _positive_seconds(timeout_connect, 10)
    read_seconds = _positive_seconds(timeout_read, 60)
    try:
        retry_count = min(1, max(0, int(retries)))
    except (TypeError, ValueError, OverflowError):
        retry_count = 1
    get = request_get or requests.get
    endpoint = f"{base_url.strip().rstrip('/')}/api/me"
    for attempt in range(retry_count + 1):
        try:
            response = get(endpoint, auth=(username, password), timeout=(connect_seconds, read_seconds))
        except requests.exceptions.SSLError:
            return False, "La connexion sécurisée à DHIS2 a échoué (certificat TLS). Contactez l’administrateur."
        except requests.exceptions.ReadTimeout:
            if attempt < retry_count:
                continue
            return False, (
                f"DHIS2 n’a pas répondu dans le délai de lecture de {read_seconds:g} secondes. "
                "Réessayez dans quelques instants."
            )
        except requests.exceptions.ConnectionError:
            if attempt < retry_count:
                continue
            return False, "Connexion au serveur DHIS2 impossible. Vérifiez l’adresse et votre connexion, puis réessayez."
        except (requests.exceptions.InvalidURL, requests.exceptions.MissingSchema, requests.exceptions.InvalidSchema):
            return False, "L’adresse du serveur DHIS2 est invalide. Contactez l’administrateur."
        except Exception:
            return False, "La connexion à DHIS2 a échoué. Veuillez réessayer."
        status = response.status_code
        if status == 200:
            return True, None
        if status == 401:
            return False, "Nom d’utilisateur ou mot de passe DHIS2 incorrect. Vérifiez vos identifiants."
        if status == 403:
            return False, "Accès au compte refusé par DHIS2 (HTTP 403). Vérifiez vos identifiants ou contactez l’administrateur."
        if status in (502, 503, 504):
            if attempt < retry_count:
                continue
            return False, f"Le serveur DHIS2 est temporairement indisponible (HTTP {status}). Réessayez dans quelques instants."
        return False, f"Connexion DHIS2 échouée (HTTP {status}). Contactez l’administrateur si le problème persiste."
