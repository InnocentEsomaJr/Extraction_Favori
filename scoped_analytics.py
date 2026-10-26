"""Requête Analytics d'un favori, limitée aux unités réellement sélectionnées.

Les périodes sont un filtre agrégé, afin que les taux ne soient jamais obtenus
par une moyenne locale des pourcentages mensuels.
"""

import re

import pandas as pd


METRIC_LABELS = {
    "REPORTING_RATE": "Reporting rate",
    "REPORTING_RATE_ON_TIME": "Reporting rate on time",
    "ACTUAL_REPORTS": "Actual reports",
    "ACTUAL_REPORTS_ON_TIME": "Actual reports on time",
    "EXPECTED_REPORTS": "Expected reports",
}
ORG_COLUMNS = ["Organisation unit", "Organisation unit ID"]


def _id(value):
    if isinstance(value, dict):
        value = value.get("id") or value.get("uid")
    return str(value).strip() if value is not None else ""


def _name(value):
    if not isinstance(value, dict):
        return ""
    return str(value.get("displayName") or value.get("name") or "").strip()


def _metric_parts(identifier):
    base, separator, metric = str(identifier).rpartition(".")
    return (base, metric) if separator and metric in METRIC_LABELS else (None, None)


def _canonical_label(identifier, label="", metadata=None):
    dataset, metric = _metric_parts(identifier)
    if metric is None:
        return label or identifier
    suffix = METRIC_LABELS[metric]
    dataset_name = _name((metadata or {}).get(dataset, {}))
    if not dataset_name and label:
        ending = re.compile(r"[\s_\-]+".join(re.escape(word) for word in suffix.split()) + r"\s*$", re.I)
        if ending.search(label):
            dataset_name = ending.sub("", label).strip(" -_")
    return f"{dataset_name or dataset} {suffix}"


def _data_item(item):
    if not isinstance(item, dict):
        identifier = _id(item)
        return identifier, _canonical_label(identifier)
    reporting = item.get("reportingRate") or item.get("dataSetReportingRate")
    if reporting is None and (item.get("metric") or item.get("dataDimensionItemType") == "REPORTING_RATE"):
        reporting = item
    if reporting is not None:
        reporting = reporting if isinstance(reporting, dict) else {"id": reporting}
        dataset = reporting.get("dataSet") or item.get("dataSet") or {}
        identifier = _id(reporting)
        metric = reporting.get("metric") or item.get("metric")
        if _id(dataset) and metric:
            identifier = f"{_id(dataset)}.{metric}"
        _, parsed_metric = _metric_parts(identifier)
        if parsed_metric is None:
            raise ValueError("Une métrique de reporting du favori est absente ou non prise en charge.")
        label = (
            f"{_name(dataset)} {METRIC_LABELS[parsed_metric]}"
            if _name(dataset)
            else _canonical_label(identifier, _name(reporting) or _name(item))
        )
        return identifier, label
    for kind in ["indicator", "dataElement", "programIndicator", "dataElementOperand"]:
        candidate = item.get(kind)
        if candidate:
            if kind == "dataElementOperand" and not _id(candidate):
                identifier = _id(candidate.get("dataElement"))
                category = _id(candidate.get("categoryOptionCombo"))
                if identifier and category:
                    identifier += "." + category
            else:
                identifier = _id(candidate)
            return identifier, _name(candidate) or identifier
    identifier = _id(item)
    return identifier, _canonical_label(identifier, _name(item))


def _unique_labels(labels):
    duplicates = {label for label in labels.values() if list(labels.values()).count(label) > 1}
    resolved = {}
    for identifier, label in labels.items():
        if label in duplicates:
            dataset, metric = _metric_parts(identifier)
            if metric:
                suffix = METRIC_LABELS[metric]
                label = f"{label[:-len(suffix)].strip()} ({dataset}) {suffix}"
            else:
                label += f" [{identifier}]"
        resolved[identifier] = label
    if len(set(resolved.values())) != len(resolved):
        raise ValueError("Les libellés des éléments du favori ne sont pas uniques.")
    return resolved


def _enrich_labels(labels, metadata):
    """Enrichit les libellés de secours sans remplacer ceux choisis dans le favori."""
    enriched = {}
    for identifier, label in labels.items():
        fallback = _canonical_label(identifier)
        if not label or label in {identifier, fallback}:
            label = _canonical_label(identifier, _name(metadata.get(identifier, {})), metadata)
        enriched[identifier] = label
    return enriched


def build_analytics_query(visualization, period, ou_dimension):
    """Retourne ``(paramètres GET répétés, libellés par identifiant dx)``.

    Les dimensions de données et les filtres supplémentaires explicites sont
    conservés. Les dimensions organisationnelles et temporelles du favori sont
    remplacées par le périmètre choisi. Une dimension supplémentaire sans
    sélection précise est refusée plutôt qu'agrégée implicitement.
    """
    if not isinstance(visualization, dict):
        raise ValueError("La définition du favori DHIS2 est invalide.")
    if not period or not ou_dimension:
        raise ValueError("La période et le périmètre géographique sont nécessaires.")
    if isinstance(period, (list, tuple)):
        period = ";".join(str(value) for value in period)
    labels = {}
    for item in visualization.get("dataDimensionItems", []) or []:
        identifier, label = _data_item(item)
        if not identifier:
            raise ValueError("Un élément de données du favori n'est pas pris en charge.")
        labels[identifier] = label

    extra = {}
    layouts = []
    for location in ["columns", "rows", "filters"]:
        for dimension in visualization.get(location, []) or []:
            if not isinstance(dimension, dict):
                raise ValueError("Une dimension du favori n'est pas prise en charge.")
            dimension_id = str(dimension.get("dimension") or dimension.get("id") or "").strip()
            if not dimension_id:
                raise ValueError("Une dimension du favori ne possède pas d'identifiant.")
            items = dimension.get("items") or []
            if dimension_id == "dx":
                for item in items:
                    identifier, label = _data_item(item)
                    if identifier:
                        labels.setdefault(identifier, label)
            elif dimension_id not in {"pe", "ou"}:
                layouts.append(dimension_id)
                if items:
                    extra[dimension_id] = [_id(value) for value in items]

    detail_dimensions = {
        "categoryDimensions": ("category", "categoryOptions"),
        "categoryOptionGroupSetDimensions": ("categoryOptionGroupSet", "categoryOptionGroups"),
        "organisationUnitGroupSetDimensions": ("organisationUnitGroupSet", "organisationUnitGroups"),
        "dataElementGroupSetDimensions": ("dataElementGroupSet", "dataElementGroups"),
    }
    for field, (owner_key, items_key) in detail_dimensions.items():
        for dimension in visualization.get(field, []) or []:
            dimension_id = _id(dimension.get(owner_key))
            items = [_id(value) for value in dimension.get(items_key, []) or []]
            if not dimension_id or not items or any(not value for value in items):
                raise ValueError(f"Une sélection explicite est nécessaire pour la dimension {field}.")
            extra[dimension_id] = items
    if any(visualization.get(field) for field in ["attributeDimensions", "dataElementDimensions", "programIndicatorDimensions"]):
        raise ValueError("Ce favori comporte un filtre Tracker non pris en charge par les analyses agrégées.")
    for dimension_id in layouts:
        if not extra.get(dimension_id):
            raise ValueError(f"La dimension {dimension_id} du favori nécessite une sélection explicite.")
    if not labels:
        raise ValueError("Le favori ne contient aucun élément de données exploitable.")
    labels = _unique_labels(labels)
    params = [
        ("dimension", "dx:" + ";".join(labels)),
        ("dimension", "ou:" + str(ou_dimension)),
        ("filter", "pe:" + str(period)),
        ("includeMetadataDetails", "true"),
    ]
    for dimension_id, items in extra.items():
        if any(not value for value in items):
            raise ValueError(f"La sélection de la dimension {dimension_id} est invalide.")
        params.append(("filter", dimension_id + ":" + ";".join(dict.fromkeys(items))))
    for option in ["aggregationType", "skipRounding", "completedOnly"]:
        value = visualization.get(option)
        if value is not None:
            params.append((option, str(value).lower() if isinstance(value, bool) else str(value)))
    return params, labels


def analytics_to_reporting_frame(response, labels=None):
    """Transforme une réponse Analytics ``ou,dx,value`` sans agréger localement."""
    if not isinstance(response, dict) or not isinstance(response.get("rows"), list):
        raise ValueError("La réponse Analytics DHIS2 est invalide.")
    requested_labels = dict(labels or {})
    metadata = (response.get("metaData") or {}).get("items") or {}
    requested_labels = _enrich_labels(requested_labels, metadata)
    if not response["rows"]:
        return pd.DataFrame(columns=ORG_COLUMNS + list(_unique_labels(requested_labels).values()))
    headers = [header.get("name", header.get("column")) for header in response.get("headers", [])]
    if set(headers) != {"ou", "dx", "value"} or len(headers) != 3:
        raise ValueError("La réponse Analytics doit contenir uniquement les dimensions ou, dx et value.")
    if any(not isinstance(row, (list, tuple)) or len(row) != len(headers) for row in response["rows"]):
        raise ValueError("Une ligne Analytics DHIS2 est invalide.")
    data = pd.DataFrame(response["rows"], columns=headers)
    if data[["ou", "dx"]].isna().any().any() or (data[["ou", "dx"]].astype(str).eq("")).any().any():
        raise ValueError("Une unité ou un élément de données Analytics n'a pas d'identifiant.")
    if data.duplicated(["ou", "dx"]).any():
        raise ValueError("Analytics retourne plusieurs valeurs pour une même unité et métrique ; aucune moyenne implicite n'est appliquée.")
    numeric = pd.to_numeric(data["value"], errors="coerce")
    if (numeric.isna() & data["value"].notna() & data["value"].astype(str).str.strip().ne("")).any():
        raise ValueError("Une valeur Analytics n'est pas numérique.")
    data["value"] = numeric
    for identifier in data["dx"].unique():
        if identifier not in requested_labels:
            requested_labels[identifier] = _canonical_label(identifier, _name(metadata.get(identifier, {})), metadata)
    requested_labels = _unique_labels(requested_labels)
    wide = data.pivot(index="ou", columns="dx", values="value").reindex(
        index=data["ou"].drop_duplicates(), columns=list(requested_labels)
    )
    wide = wide.rename(columns=requested_labels)
    wide.columns.name = None
    wide.index.name = "Organisation unit ID"
    result = wide.reset_index()
    result.insert(0, "Organisation unit", result["Organisation unit ID"].map(lambda identifier: _name(metadata.get(identifier, {})) or identifier))
    return result
