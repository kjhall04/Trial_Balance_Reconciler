from __future__ import annotations

import json
from pathlib import Path

from .models import ClientConfig, WorkbookRule


def _normalize_key(value: str | Path) -> str:
    return str(value).strip().lower()


def _candidate_keys(path: Path) -> list[str]:
    return [
        _normalize_key(path),
        _normalize_key(path.name),
        _normalize_key(path.stem),
    ]


def parse_entity_overrides(value: str) -> dict[str, str]:
    overrides: dict[str, str] = {}
    for raw_line in (value or "").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, entity = line.split("=", 1)
        key = key.strip().lower()
        entity = entity.strip()
        if key and entity:
            overrides[key] = entity
    return overrides


def split_path_text(value: str) -> list[str]:
    return [piece.strip() for piece in (value or "").split("|") if piece.strip()]


def _coerce_sheet_name(value: object) -> str | int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    return text


def load_client_config(path: str | Path) -> ClientConfig:
    config_path = Path(path).expanduser()
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("Client config must be a JSON object.")

    raw_overrides = raw.get("entity_overrides", {}) or {}
    if not isinstance(raw_overrides, dict):
        raise ValueError("'entity_overrides' must be a JSON object.")
    entity_overrides = {
        _normalize_key(key): str(value).strip()
        for key, value in raw_overrides.items()
        if str(key).strip() and str(value).strip()
    }

    raw_rules = raw.get("workbooks", {}) or {}
    if not isinstance(raw_rules, dict):
        raise ValueError("'workbooks' must be a JSON object.")
    workbook_rules: dict[str, WorkbookRule] = {}
    for key, value in raw_rules.items():
        if not isinstance(value, dict):
            raise ValueError("Each workbook rule must be a JSON object.")
        workbook_rules[_normalize_key(key)] = WorkbookRule(
            parser=str(value.get("parser", "")).strip(),
            sheet_name=_coerce_sheet_name(value.get("sheet_name")),
            entity=str(value.get("entity", "")).strip(),
        )

    return ClientConfig(
        name=str(raw.get("name", "")).strip(),
        entity_overrides=entity_overrides,
        workbook_rules=workbook_rules,
        source_path=config_path,
    )


def entity_override_for_path(config: ClientConfig | None, path: Path) -> str:
    if config is None:
        return ""
    for key in _candidate_keys(path):
        if key in config.entity_overrides:
            return config.entity_overrides[key]
    return ""


def workbook_rule_for_path(config: ClientConfig | None, path: Path) -> WorkbookRule:
    if config is None:
        return WorkbookRule()
    for key in _candidate_keys(path):
        if key in config.workbook_rules:
            return config.workbook_rules[key]
    return WorkbookRule()

