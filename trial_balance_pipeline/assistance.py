from __future__ import annotations

import copy
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .config import _coerce_sheet_name
from .models import ClientConfig, WorkbookPreview, WorkbookRule
from .normalize import normalize_match_text


GENERIC_WORKBOOK_TOKENS = {
    "balance",
    "book",
    "books",
    "client",
    "current",
    "data",
    "export",
    "file",
    "final",
    "gl",
    "import",
    "prelim",
    "qb",
    "quickbooks",
    "raw",
    "review",
    "tb",
    "trial",
    "workbook",
    "year",
}


@dataclass(frozen=True)
class PreflightItem:
    path: Path
    suggested_entity: str
    entity_source: str
    parser_name: str
    sheet_name: str | int | None
    note: str = ""
    needs_entity_confirmation: bool = False
    entity_options: tuple[str, ...] = ()


def default_user_memory() -> dict:
    return {
        "version": 1,
        "workbook_rules": {},
        "entities": [],
    }


def sanitize_user_memory(raw: object) -> dict:
    base = default_user_memory()
    if not isinstance(raw, dict):
        return base

    workbook_rules = raw.get("workbook_rules", {})
    if isinstance(workbook_rules, dict):
        cleaned_rules: dict[str, dict[str, object]] = {}
        for key, value in workbook_rules.items():
            if not str(key).strip() or not isinstance(value, dict):
                continue
            cleaned_rules[str(key).strip().lower()] = {
                "parser": str(value.get("parser", "")).strip(),
                "sheet_name": _coerce_sheet_name(value.get("sheet_name")),
                "entity": str(value.get("entity", "")).strip(),
            }
        base["workbook_rules"] = cleaned_rules

    entities = raw.get("entities", [])
    if isinstance(entities, list):
        unique_entities: list[str] = []
        seen: set[str] = set()
        for value in entities:
            text = str(value).strip()
            key = normalize_match_text(text)
            if not text or not key or key in seen:
                continue
            seen.add(key)
            unique_entities.append(text)
        base["entities"] = unique_entities

    return base


def build_memory_client_config(memory: dict | None) -> ClientConfig | None:
    cleaned = sanitize_user_memory(memory)
    raw_rules = cleaned["workbook_rules"]
    if not raw_rules:
        return None

    workbook_rules: dict[str, WorkbookRule] = {}
    entity_overrides: dict[str, str] = {}
    for key, value in raw_rules.items():
        rule = WorkbookRule(
            parser=str(value.get("parser", "")).strip(),
            sheet_name=_coerce_sheet_name(value.get("sheet_name")),
            entity=str(value.get("entity", "")).strip(),
        )
        workbook_rules[str(key)] = rule
        if rule.entity:
            entity_overrides[str(key)] = rule.entity

    return ClientConfig(
        name="Learned Client Memory",
        entity_overrides=entity_overrides,
        workbook_rules=workbook_rules,
        source_path=None,
    )


def known_entities_from_memory(memory: dict | None) -> list[str]:
    cleaned = sanitize_user_memory(memory)
    entities: list[str] = []
    seen: set[str] = set()

    for value in cleaned["entities"]:
        key = normalize_match_text(value)
        if key and key not in seen:
            seen.add(key)
            entities.append(str(value).strip())

    for value in cleaned["workbook_rules"].values():
        entity = str(value.get("entity", "")).strip()
        key = normalize_match_text(entity)
        if key and key not in seen:
            seen.add(key)
            entities.append(entity)

    return entities


def _workbook_memory_keys(path: Path) -> list[str]:
    return [path.name.strip().lower(), path.stem.strip().lower()]


def _is_distinctive_workbook_name(path: Path) -> bool:
    tokens = [token for token in normalize_match_text(path.stem).split() if token and token not in GENERIC_WORKBOOK_TOKENS]
    if not tokens:
        return False
    if any(len(token) >= 3 for token in tokens):
        return True
    return len(tokens) >= 2


def remember_successful_workbooks(memory: dict | None, current_df: pd.DataFrame) -> dict:
    cleaned = sanitize_user_memory(memory)
    updated = copy.deepcopy(cleaned)
    workbook_rules = dict(updated.get("workbook_rules", {}))
    entity_values = list(updated.get("entities", []))
    seen_entities = {normalize_match_text(value) for value in entity_values}

    if current_df.empty:
        updated["workbook_rules"] = workbook_rules
        updated["entities"] = entity_values
        return sanitize_user_memory(updated)

    grouped = current_df.groupby("source_file", dropna=False)
    for source_file, frame in grouped:
        path_text = str(source_file).strip()
        if not path_text:
            continue
        path = Path(path_text)
        if not _is_distinctive_workbook_name(path):
            continue

        entity_candidates = [str(value).strip() for value in frame["entity"].tolist() if str(value).strip()]
        parser_candidates = [str(value).strip() for value in frame["parser_name"].tolist() if str(value).strip()]
        sheet_candidates = [value for value in frame["source_sheet"].tolist() if str(value).strip()]

        entity = entity_candidates[0] if len(set(entity_candidates)) == 1 else ""
        parser_name = parser_candidates[0] if len(set(parser_candidates)) == 1 else ""
        sheet_name = sheet_candidates[0] if len({str(value).strip() for value in sheet_candidates}) == 1 else None

        payload = {
            "parser": parser_name,
            "sheet_name": sheet_name,
            "entity": entity,
        }
        for key in _workbook_memory_keys(path):
            workbook_rules[key] = payload.copy()

        entity_key = normalize_match_text(entity)
        if entity and entity_key and entity_key not in seen_entities:
            seen_entities.add(entity_key)
            entity_values.append(entity)

    updated["workbook_rules"] = workbook_rules
    updated["entities"] = entity_values
    return sanitize_user_memory(updated)


def merge_known_entities(*groups: list[str]) -> list[str]:
    merged: list[str] = []
    seen: set[str] = set()
    for group in groups:
        for value in group:
            text = str(value).strip()
            key = normalize_match_text(text)
            if not text or not key or key in seen:
                continue
            seen.add(key)
            merged.append(text)
    return merged


def build_preflight_items(
    previews: list[WorkbookPreview],
    *,
    known_entities: list[str] | None = None,
    prior_entities: list[str] | None = None,
    memory_entities: list[str] | None = None,
) -> list[PreflightItem]:
    known = merge_known_entities(known_entities or [], prior_entities or [], memory_entities or [])
    prior_entity_keys = {normalize_match_text(value) for value in (prior_entities or []) if normalize_match_text(value)}
    items: list[PreflightItem] = []

    for preview in previews:
        note = ""
        needs_confirmation = False

        if preview.entity_source == "derived from file name":
            needs_confirmation = True
            note = (
                "The program only has the file name for this workbook. Confirm the company name once and it will remember it next time."
            )
        elif prior_entity_keys and normalize_match_text(preview.entity) not in prior_entity_keys:
            note = (
                "No prior-year official TB rows were found for this company. The build can still run, "
                "but a consolidated official prior-year TB workbook may give a closer match."
            )

        options = merge_known_entities([preview.entity], known)
        items.append(
            PreflightItem(
                path=preview.path,
                suggested_entity=preview.entity,
                entity_source=preview.entity_source,
                parser_name=preview.parser_name,
                sheet_name=preview.sheet_name,
                note=note,
                needs_entity_confirmation=needs_confirmation,
                entity_options=tuple(options),
            )
        )

    return items
