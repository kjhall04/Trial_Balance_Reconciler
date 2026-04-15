from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

from .config import entity_override_for_path, workbook_rule_for_path
from .models import ClientConfig, WorkbookPreview, WorkbookSpec
from .normalize import (
    clean_account_name,
    clean_text,
    extract_leaf_account_number,
    extract_parent_account_number,
    extract_path_numbers,
    is_zero_balance,
    normalize_match_text,
    path_depth,
)


CURRENT_PARSER_PROFILES = (
    "quickbooks_debit_credit",
    "extracted_balance_list",
)

ENTITY_SUFFIX_TOKENS = {
    "inc",
    "incorporated",
    "llc",
    "l l c",
    "corp",
    "corporation",
    "co",
    "company",
    "ltd",
    "limited",
    "lp",
    "llp",
    "pllc",
}


def available_current_parser_profiles() -> list[str]:
    return list(CURRENT_PARSER_PROFILES)


def _excel_column_label(index: int) -> str:
    return get_column_letter(int(index) + 1)


def _entity_tokens(value: str) -> list[str]:
    normalized = "".join(ch.lower() if ch.isalnum() else " " for ch in value)
    return [token for token in normalized.split() if token and token not in ENTITY_SUFFIX_TOKENS]


def _infer_entity_from_candidates(path: Path, known_entities: list[str]) -> str:
    if not known_entities:
        return ""
    stem = path.stem.lower()
    stem_alnum = "".join(ch for ch in stem if ch.isalnum())
    best_entity = ""
    best_score = -1
    for entity in known_entities:
        lowered = entity.lower()
        entity_alnum = "".join(ch for ch in lowered if ch.isalnum())
        score = 0
        if lowered in stem:
            score += 4
        if entity_alnum and entity_alnum in stem_alnum:
            score += 3
        meaningful_tokens = _entity_tokens(entity)
        initials = "".join(token[0] for token in meaningful_tokens if token)
        if initials and initials in stem_alnum:
            score += 4
        first_word = meaningful_tokens[0] if meaningful_tokens else ""
        if first_word and first_word in stem:
            score += 1
        if score > best_score:
            best_score = score
            best_entity = entity
    return best_entity if best_score > 0 else ""


def _resolve_entity(
    spec: WorkbookSpec,
    path: Path,
    config: ClientConfig | None,
    entity_overrides: dict[str, str] | None,
    known_entities: list[str],
) -> tuple[str, str]:
    if spec.entity:
        return spec.entity, "provided directly"

    rule = workbook_rule_for_path(config, path)
    if rule.entity:
        return rule.entity, "setup file workbook rule"

    manual = ""
    if entity_overrides:
        for key in (str(path).lower(), path.name.lower(), path.stem.lower()):
            if key in entity_overrides:
                manual = entity_overrides[key]
                break
    if manual:
        return manual, "manual company match"

    configured = entity_override_for_path(config, path)
    if configured:
        return configured, "setup file company match"

    inferred = _infer_entity_from_candidates(path, known_entities)
    if inferred:
        return inferred, "matched from known companies"

    return path.stem, "derived from file name"


def _choose_sheet(path: Path, preferred: str | int | None = None) -> tuple[str | int, str]:
    if preferred is not None and preferred != "":
        return preferred, "setup file workbook rule"
    workbook = pd.ExcelFile(path)
    try:
        if "Sheet1" in workbook.sheet_names:
            return "Sheet1", "selected automatically"
        if workbook.sheet_names:
            return workbook.sheet_names[0], "selected automatically"
        raise ValueError(f"No worksheets were found in '{path.name}'.")
    finally:
        workbook.close()


def _detect_quickbooks_debit_credit(raw: pd.DataFrame) -> tuple[int, int, int] | None:
    for row_index, row in raw.iterrows():
        lowered = [clean_text(value).lower() for value in row.tolist()]
        if "debit" in lowered and "credit" in lowered:
            debit_index = lowered.index("debit")
            credit_index = lowered.index("credit")
            return int(row_index), debit_index, credit_index
    return None


def _detect_extracted_balance_list(raw: pd.DataFrame) -> bool:
    sample = raw.iloc[:12, :4].fillna("")
    nonblank_rows = 0
    for _, row in sample.iterrows():
        values = [clean_text(value) for value in row.tolist()[:4]]
        if len([value for value in values if value]) >= 3:
            nonblank_rows += 1
    return nonblank_rows >= 2


def _detect_parser(path: Path, sheet_name: str | int) -> tuple[str, str]:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    if _detect_quickbooks_debit_credit(raw) is not None:
        return "quickbooks_debit_credit", "detected automatically"
    if _detect_extracted_balance_list(raw):
        return "extracted_balance_list", "detected automatically"
    raise ValueError(f"Could not recognize the current-year workbook layout for '{path.name}'.")


def preview_current_workbooks(
    specs: list[WorkbookSpec],
    *,
    entity_overrides: dict[str, str] | None = None,
    known_entities: list[str] | None = None,
    client_config: ClientConfig | None = None,
) -> list[WorkbookPreview]:
    previews: list[WorkbookPreview] = []
    resolved_entities = known_entities or []
    for spec in specs:
        path = spec.path.expanduser()
        entity, entity_source = _resolve_entity(spec, path, client_config, entity_overrides, resolved_entities)
        rule = workbook_rule_for_path(client_config, path)
        sheet_name, sheet_source = _choose_sheet(path, rule.sheet_name)
        if rule.parser:
            parser_name = rule.parser
            parser_source = "setup file workbook rule"
        else:
            parser_name, parser_source = _detect_parser(path, sheet_name)
        previews.append(
            WorkbookPreview(
                path=path,
                entity=entity,
                entity_source=entity_source,
                parser_name=parser_name,
                parser_source=parser_source,
                sheet_name=sheet_name,
                sheet_source=sheet_source,
            )
        )
    return previews


def _read_quickbooks_debit_credit(path: Path, sheet_name: str | int) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    detection = _detect_quickbooks_debit_credit(raw)
    if detection is None:
        raise ValueError(f"Could not locate a Debit/Credit header row in '{path.name}'.")
    header_row, debit_index, credit_index = detection
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    columns = list(df.columns)
    account_index = max(debit_index - 1, 0)
    debit_column = columns[debit_index]
    credit_column = columns[credit_index]
    account_column = columns[account_index]

    out = pd.DataFrame(
        {
            "raw_account_text": df[account_column],
            "debit": pd.to_numeric(df[debit_column], errors="coerce").fillna(0.0),
            "credit": pd.to_numeric(df[credit_column], errors="coerce").fillna(0.0),
        }
    )
    out["cy_balance"] = out["debit"] - out["credit"]
    out["raw_account_text"] = out["raw_account_text"].fillna("").astype(str).str.strip()
    out["source_sheet"] = str(sheet_name)
    out["source_row"] = (df.index.to_series() + header_row + 2).astype(int)
    out["source_text_column"] = _excel_column_label(account_index)
    out["source_number_column"] = ""
    out["source_amount_column"] = f"{_excel_column_label(debit_index)}/{_excel_column_label(credit_index)}"
    return out[
        [
            "raw_account_text",
            "cy_balance",
            "source_sheet",
            "source_row",
            "source_text_column",
            "source_number_column",
            "source_amount_column",
        ]
    ]


def _read_extracted_balance_list(path: Path, sheet_name: str | int) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    first_data_row = 0
    for index, row in raw.iterrows():
        values = [clean_text(value) for value in row.tolist()[:4]]
        if len([value for value in values if value]) >= 3:
            first_data_row = int(index)
            break

    trimmed = raw.iloc[first_data_row:, :4].copy()
    trimmed.columns = ["entity_value", "bucket_account_number", "raw_account_text", "cy_balance"]
    trimmed["raw_account_text"] = trimmed["raw_account_text"].fillna("").astype(str).str.strip()
    trimmed["bucket_account_number"] = trimmed["bucket_account_number"].fillna("").astype(str).str.strip()
    trimmed["cy_balance"] = pd.to_numeric(trimmed["cy_balance"], errors="coerce").fillna(0.0)
    trimmed["source_sheet"] = str(sheet_name)
    trimmed["source_row"] = (trimmed.index.to_series() + 1).astype(int)
    trimmed["source_text_column"] = "C"
    trimmed["source_number_column"] = "B"
    trimmed["source_amount_column"] = "D"
    return trimmed[
        [
            "entity_value",
            "bucket_account_number",
            "raw_account_text",
            "cy_balance",
            "source_sheet",
            "source_row",
            "source_text_column",
            "source_number_column",
            "source_amount_column",
        ]
    ]


def read_current_workbooks(
    specs: list[WorkbookSpec],
    *,
    entity_overrides: dict[str, str] | None = None,
    known_entities: list[str] | None = None,
    client_config: ClientConfig | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for spec in specs:
        path = spec.path.expanduser()
        entity, _ = _resolve_entity(spec, path, client_config, entity_overrides, known_entities or [])
        rule = workbook_rule_for_path(client_config, path)
        sheet_name, _ = _choose_sheet(path, rule.sheet_name)
        parser_name = rule.parser or _detect_parser(path, sheet_name)[0]

        if parser_name == "quickbooks_debit_credit":
            df = _read_quickbooks_debit_credit(path, sheet_name)
            df["bucket_account_number"] = ""
        elif parser_name == "extracted_balance_list":
            df = _read_extracted_balance_list(path, sheet_name)
            if "entity_value" in df.columns:
                observed_entities = [clean_text(value) for value in df["entity_value"].unique().tolist() if clean_text(value)]
                if len(observed_entities) == 1:
                    entity = observed_entities[0]
                elif observed_entities and not spec.entity and not rule.entity and not entity_overrides:
                    entity = observed_entities[0]
        else:
            raise ValueError(f"Unsupported current-year parser '{parser_name}'.")

        df["entity"] = entity
        df["source_file"] = str(path)
        df["parser_name"] = parser_name
        df["raw_account_text"] = df["raw_account_text"].fillna("").astype(str).str.strip()
        df["cy_balance"] = pd.to_numeric(df["cy_balance"], errors="coerce").fillna(0.0)
        df = df[df["raw_account_text"] != ""].copy()
        df = df[~df["cy_balance"].apply(is_zero_balance)].copy()
        df["path_numbers"] = [
            tuple(extract_path_numbers(text, fallback=fallback))
            for text, fallback in zip(df["raw_account_text"], df["bucket_account_number"])
        ]
        df["raw_account_number"] = [
            extract_leaf_account_number(text, fallback=fallback)
            for text, fallback in zip(df["raw_account_text"], df["bucket_account_number"])
        ]
        df["parent_account_number"] = [
            extract_parent_account_number(text, fallback=fallback)
            for text, fallback in zip(df["raw_account_text"], df["bucket_account_number"])
        ]
        df["account_name"] = df["raw_account_text"].map(clean_account_name)
        df["match_key"] = df["account_name"].map(normalize_match_text)
        df["path_depth"] = df["raw_account_text"].map(path_depth)
        frames.append(df)

    if not frames:
        return pd.DataFrame(
            columns=[
                "entity",
                "raw_account_text",
                "bucket_account_number",
                "raw_account_number",
                "parent_account_number",
                "path_numbers",
                "account_name",
                "match_key",
                "path_depth",
                "cy_balance",
                "source_file",
                "source_sheet",
                "source_row",
                "source_text_column",
                "source_number_column",
                "source_amount_column",
                "parser_name",
            ]
        )

    current_df = pd.concat(frames, ignore_index=True)
    current_df.insert(0, "current_row", range(1, len(current_df) + 1))
    return current_df


def build_client_config_template(
    specs: list[WorkbookSpec],
    *,
    entity_overrides: dict[str, str] | None = None,
    known_entities: list[str] | None = None,
    client_config: ClientConfig | None = None,
    client_name: str = "",
) -> dict:
    payload = {
        "name": client_name or (client_config.name if client_config else ""),
        "entity_overrides": {},
        "workbooks": {},
    }
    previews = preview_current_workbooks(
        specs,
        entity_overrides=entity_overrides,
        known_entities=known_entities,
        client_config=client_config,
    )
    for preview in previews:
        if preview.entity:
            payload["entity_overrides"][preview.path.name] = preview.entity
        payload["workbooks"][preview.path.name] = {
            "parser": preview.parser_name,
            "sheet_name": preview.sheet_name,
            "entity": preview.entity,
        }
    return payload
