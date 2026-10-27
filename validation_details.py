"""Détails des violations DHIS2, indépendants du réseau et de Streamlit."""

import math
import re
import unicodedata

import pandas as pd


DISPLAY_COLUMNS = [
    "Zone de santé", "Unité d’organisation", "Période", "Importance",
    "Règle de validation", "Valeur gauche", "Opérateur", "Valeur droite", "Détails",
]
INTERNAL_COLUMNS = ["_zone_id", "_rule_id", "_ou_id", "_period_id", "_result_id"]
IMPORTANCE_OPTIONS = ["Toutes", "HIGH", "MEDIUM", "LOW", "Non renseignée"]
_MONTHS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
    "Septembre", "Octobre", "Novembre", "Décembre",
]
_OPERATORS = {
    "equal_to": "==", "equal": "==", "equals": "==", "==": "==", "=": "=",
    "not_equal_to": "!=", "not_equal": "!=", "notequal": "!=", "!=": "!=", "<>": "!=",
    "less_than": "<", "<": "<", "less_than_or_equal_to": "<=", "less_than_or_equal": "<=", "<=": "<=",
    "greater_than": ">", ">": ">", "greater_than_or_equal_to": ">=", "greater_than_or_equal": ">=", ">=": ">=",
    "compulsory_pair": "Paire obligatoire", "exclusive_pair": "Paire exclusive",
}


def _missing(value):
    if value is None:
        return True
    if isinstance(value, (dict, list, tuple, set)):
        return not value
    return bool(pd.isna(value))


def _first(*values):
    for value in values:
        if not _missing(value) and not (isinstance(value, str) and not value.strip()):
            return value
    return None


def _text(value):
    return "" if _missing(value) else str(value).strip()


def _id(value):
    if isinstance(value, dict):
        return _id(_first(value.get("id"), value.get("uid")))
    if _missing(value):
        return ""
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "none", "null", "<na>"} else text


def _object(value):
    return value if isinstance(value, dict) else {}


def _name(value):
    item = _object(value)
    return _text(_first(item.get("displayName"), item.get("name")))


def _records(results):
    if results is None:
        return []
    if isinstance(results, pd.DataFrame):
        return results.to_dict("records")
    return [row for row in results if isinstance(row, dict)]


def _metadata(mapping, identifier):
    return _object((mapping or {}).get(identifier, {}))


def _rule_id(row):
    return _id(_first(row.get("validationRuleId"), row.get("validationRule"), row.get("_rule_id")))


def _ou_id(row):
    return _id(_first(row.get("organisationUnitId"), row.get("organisationUnit"), row.get("_ou_id")))


def validation_detail_identifiers(results):
    """Liste les IDs à enrichir, dans l'ordre des résultats, sans lire les noms."""
    rows = _records(results)
    return {
        "rule_ids": list(dict.fromkeys(identifier for row in rows if (identifier := _rule_id(row)))),
        "org_ids": list(dict.fromkeys(identifier for row in rows if (identifier := _ou_id(row)))),
    }


def _period(row):
    source = row.get("period")
    item = _object(source)
    identifier = _id(_first(
        row.get("periodId"), item.get("iso"), item.get("id"),
        source if not isinstance(source, dict) else None,
        row.get("_requested_period"), item.get("name"), item.get("startDate"),
    ))
    explicit_name = _text(_first(row.get("periodDisplayName"), item.get("displayName"), item.get("name")))
    match = re.fullmatch(r"(\d{4})(\d{2})", identifier)
    if match and 1 <= int(match[2]) <= 12:
        return identifier, f"{_MONTHS[int(match[2]) - 1]} {match[1]}"
    if re.fullmatch(r"\d{4}", identifier):
        return identifier, identifier
    return identifier, explicit_name or identifier or "Non renseignée"


def _zone(row, ou_id, embedded_ou, metadata_ou, org_metadata):
    zone_id = _id(row.get("_zone_id"))
    zone_name = _text(row.get("_zone_name"))
    ancestors = (embedded_ou.get("ancestors") or []) + (metadata_ou.get("ancestors") or [])
    zone_ancestors = [
        ancestor for ancestor in ancestors
        if isinstance(ancestor, dict) and _text(ancestor.get("level")) == "3"
    ]
    if not zone_id:
        own_level = _first(embedded_ou.get("level"), metadata_ou.get("level"))
        if _text(own_level) == "3":
            zone_id = ou_id
        elif zone_ancestors:
            candidates = list(dict.fromkeys(_id(ancestor) for ancestor in zone_ancestors if _id(ancestor)))
            if len(candidates) == 1:
                zone_id = candidates[0]
        if not zone_id:
            path = _text(_first(row.get("organisationUnitPath"), embedded_ou.get("path"), metadata_ou.get("path")))
            if path.startswith("/"):
                path_ids = [identifier for identifier in path.split("/") if identifier]
                if len(path_ids) >= 3:
                    zone_id = path_ids[2]
    if not zone_name and zone_id:
        zone_name = _name(_metadata(org_metadata, zone_id))
        if not zone_name:
            zone_name = next((_name(item) for item in zone_ancestors if _id(item) == zone_id and _name(item)), "")
        if not zone_name and zone_id == ou_id:
            zone_name = _text(_first(row.get("organisationUnitDisplayName"), _name(embedded_ou), _name(metadata_ou)))
    return zone_id, zone_name or zone_id or "Non renseignée"


def _number(value):
    if _missing(value):
        return float("nan")
    numeric = pd.to_numeric(value, errors="coerce")
    return numeric if not pd.isna(numeric) and math.isfinite(float(numeric)) else float("nan")


def _freeze(value):
    if _missing(value):
        return None
    if isinstance(value, dict):
        return tuple(sorted((str(key), _freeze(item)) for key, item in value.items()))
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(item) for item in value)
    return str(value)


def _value_key(value):
    numeric = _number(value)
    return float(numeric) if not pd.isna(numeric) else _freeze(value)


def _combo(row, key):
    source = row.get(key)
    identifier = _id(_first(row.get(key + "Id"), source))
    name = _text(_first(row.get(key + "DisplayName"), _name(source)))
    return identifier, name or identifier


def _side(embedded, metadata, key):
    first, second = _object(embedded.get(key)), _object(metadata.get(key))
    description = _text(_first(first.get("description"), second.get("description")))
    expression = _text(_first(first.get("expression"), second.get("expression")))
    if description and expression and description != expression:
        return f"{description} ({expression})"
    return description or expression


def build_validation_details(results, rule_metadata=None, org_metadata=None):
    """Normalise les résultats plats d'analyse et les anciens résultats imbriqués.

    Les observations de périodes différentes sont conservées. Les IDs et
    combinaisons servent à dédupliquer les répétitions, jamais les libellés.
    Les expressions sont affichées comme du texte et ne sont pas exécutées.
    """
    records, seen = [], set()
    for row in _records(results):
        rule_id, ou_id = _rule_id(row), _ou_id(row)
        rule, ou = _object(row.get("validationRule")), _object(row.get("organisationUnit"))
        metadata_rule, metadata_ou = _metadata(rule_metadata, rule_id), _metadata(org_metadata, ou_id)
        period_id, period_name = _period(row)
        zone_id, zone_name = _zone(row, ou_id, ou, metadata_ou, org_metadata)
        rule_name = _text(_first(
            row.get("validationRuleDescription"), row.get("validationRuleDisplayName"),
            _name(rule), _name(metadata_rule), rule.get("description"), metadata_rule.get("description"), rule_id,
        )) or "Non renseignée"
        ou_name = _text(_first(row.get("organisationUnitDisplayName"), _name(ou), _name(metadata_ou), ou_id)) or "Non renseignée"
        raw_left = _first(row.get("leftSideValue"), row.get("leftsideValue"), row.get("leftValue"))
        raw_right = _first(row.get("rightSideValue"), row.get("rightsideValue"), row.get("rightValue"))
        attribute_id, attribute_name = _combo(row, "attributeOptionCombo")
        category_id, category_name = _combo(row, "categoryOptionCombo")
        result_id = _id(_first(row.get("id"), row.get("uid"), row.get("_result_id")))
        day = _freeze(row.get("dayInPeriod"))
        context_key = (period_id, attribute_id or attribute_name, category_id or category_name, day)
        identity = (
            ("id", result_id, *context_key)
            if result_id else
            ("observation", ou_id or ou_name, rule_id or rule_name, *context_key, _value_key(raw_left), _value_key(raw_right))
        )
        if identity in seen:
            continue
        seen.add(identity)

        importance = _text(_first(row.get("importance"), rule.get("importance"), metadata_rule.get("importance")))
        importance = importance.upper() if importance.upper() in {"HIGH", "MEDIUM", "LOW"} else importance
        operator = _text(_first(row.get("operator"), rule.get("operator"), metadata_rule.get("operator")))
        operator_label = _OPERATORS.get(operator.casefold(), operator)
        details = []
        instruction = _text(_first(row.get("instruction"), rule.get("instruction"), metadata_rule.get("instruction")))
        if instruction:
            details.append(instruction)
        for label, text in [("Gauche", _side(rule, metadata_rule, "leftSide")), ("Droite", _side(rule, metadata_rule, "rightSide"))]:
            if text:
                details.append(f"{label} : {text}")
        if attribute_name:
            details.append(f"Combinaison d’attributs : {attribute_name}")
        if category_name:
            details.append(f"Combinaison de catégories : {category_name}")
        if not _missing(row.get("dayInPeriod")):
            details.append(f"Jour : {_text(row['dayInPeriod'])}")
        records.append({
            "Zone de santé": zone_name, "Unité d’organisation": ou_name, "Période": period_name,
            "Importance": importance or "Non renseignée", "Règle de validation": rule_name,
            "Valeur gauche": _number(raw_left), "Opérateur": operator_label,
            "Valeur droite": _number(raw_right), "Détails": "\n".join(details),
            "_zone_id": zone_id, "_rule_id": rule_id, "_ou_id": ou_id,
            "_period_id": period_id, "_result_id": result_id,
        })
    return pd.DataFrame(records, columns=DISPLAY_COLUMNS + INTERNAL_COLUMNS)


def _normalize(value):
    return "".join(
        character for character in unicodedata.normalize("NFKD", " ".join(_text(value).casefold().split()))
        if not unicodedata.combining(character)
    )


def filter_validation_details(df, query="", importance="Toutes", zone_id=None, period_id=None):
    """Filtre une copie par texte littéral, importance et identifiants exacts."""
    result = df.copy()
    if result.empty:
        return result
    for column, identifier in [("_zone_id", zone_id), ("_period_id", period_id)]:
        if identifier is not None:
            if column not in result.columns:
                return result.iloc[:0].copy()
            result = result.loc[result[column].map(_id).eq(_id(identifier))].copy()
    if _normalize(importance) != "toutes":
        if "Importance" not in result.columns:
            return result.iloc[:0].copy()
        result = result.loc[result["Importance"].map(_normalize).eq(_normalize(importance))].copy()
    needle = _normalize(query)
    if needle:
        searchable = pd.Series("", index=result.index, dtype=object)
        for column in DISPLAY_COLUMNS:
            if column in result.columns:
                searchable = searchable + " " + result[column].map(_normalize)
        result = result.loc[searchable.str.contains(needle, regex=False)].copy()
    return result.copy()


def detail_export_frame(df, include_ids=False):
    """Expose les colonnes lisibles, et les IDs techniques seulement sur demande."""
    result = df.reindex(columns=DISPLAY_COLUMNS).copy()
    if include_ids:
        for source, label in [("_rule_id", "ID règle"), ("_ou_id", "ID unité")]:
            result[label] = df[source] if source in df.columns else ""
    return result


def style_importance(value):
    colors = {
        "HIGH": ("white", "#991B1B"),
        "MEDIUM": ("black", "#FBBF24"),
        "LOW": ("#1E3A8A", "#DBEAFE"),
    }
    color = colors.get(_text(value).upper())
    return f"color: {color[0]}; background-color: {color[1]}" if color else ""
