from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

from .config import entity_override_for_path, workbook_rule_for_path
from .models import ClientConfig, WorkbookSpec
from .normalize import clean_text, normalize_match_text, standard_account_number


HEADER_ALIASES = {
    "entity": {"entity", "entity fund", "entity / fund", "fund"},
    "leadsheet": {"leadsheet"},
    "account_number": {"account number", "acct no", "acct number", "account no", "account #"},
    "account_name": {"account name", "account", "description"},
    "prior_balance": {"final", "previous year rep", "prior year", "previous year", "balance"},
}

PRIOR_YEAR_COLUMNS = [
    "prior_row",
    "entity",
    "class",
    "acct_no",
    "account",
    "py_balance",
    "match_key",
    "source_file",
    "source_sheet",
    "source_row",
    "source_entity_column",
    "source_class_column",
    "source_acct_column",
    "source_account_column",
    "source_balance_column",
]


def _excel_column_label(index: int) -> str:
    return get_column_letter(int(index) + 1)


def _normalized_header(value: object) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _find_header_row(raw: pd.DataFrame) -> int:
    for index, row in raw.iterrows():
        normalized = {_normalized_header(value) for value in row.tolist()}
        if "leadsheet" in normalized and normalized & HEADER_ALIASES["account_number"] and normalized & HEADER_ALIASES["account_name"]:
            return int(index)
    raise ValueError("Could not locate a prior-year trial-balance header row.")


def _find_named_column(df: pd.DataFrame, alias_key: str, required: bool = True) -> str | None:
    aliases = HEADER_ALIASES[alias_key]
    for column in df.columns:
        if _normalized_header(column) in aliases:
            return str(column)
    if required:
        raise ValueError(f"Could not find the '{alias_key}' column.")
    return None


def _column_letter(columns: list[object], target: object) -> str:
    try:
        return _excel_column_label(columns.index(target))
    except ValueError:
        return ""


def _infer_entity_from_title(path: Path) -> str:
    workbook = pd.read_excel(path, sheet_name=0, header=None)
    title_cell = clean_text(workbook.iloc[0, 0]) if len(workbook.index) and len(workbook.columns) else ""
    if title_cell:
        title = re.sub(r"\b\d{4}\b", "", title_cell, flags=re.IGNORECASE)
        title = re.sub(r"\btrial balance\b", "", title, flags=re.IGNORECASE)
        title = re.sub(r"\b(audit|review|tax|consolidated)\b", "", title, flags=re.IGNORECASE)
        title = re.sub(r"\s*-\s*", " ", title)
        title = re.sub(r"\s+", " ", title).strip(" -")
        if title:
            return title

    stem = path.stem
    stem = re.sub(r"\b\d{4}\b", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\btrial balance\b", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\b(audit|review|tax|consolidated)\b", "", stem, flags=re.IGNORECASE)
    stem = re.sub(r"[\-()]+", " ", stem)
    return re.sub(r"\s+", " ", stem).strip()


def read_prior_workbooks(
    specs: list[WorkbookSpec],
    *,
    entity_overrides: dict[str, str] | None = None,
    client_config: ClientConfig | None = None,
) -> pd.DataFrame:
    if not specs:
        return pd.DataFrame(columns=PRIOR_YEAR_COLUMNS)

    frames: list[pd.DataFrame] = []
    for spec in specs:
        path = spec.path.expanduser()
        rule = workbook_rule_for_path(client_config, path)
        sheet_name = rule.sheet_name if rule.sheet_name is not None else 0
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        header_row = _find_header_row(raw)
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)

        entity_col = _find_named_column(df, "entity", required=False)
        leadsheet_col = _find_named_column(df, "leadsheet")
        acct_col = _find_named_column(df, "account_number")
        account_col = _find_named_column(df, "account_name")
        balance_col = _find_named_column(df, "prior_balance")
        columns = list(df.columns)

        entity_value = spec.entity or rule.entity
        if not entity_value:
            manual = ""
            if entity_overrides:
                for key in (str(path).lower(), path.name.lower(), path.stem.lower()):
                    if key in entity_overrides:
                        manual = entity_overrides[key]
                        break
            entity_value = manual or entity_override_for_path(client_config, path) or _infer_entity_from_title(path)

        out = pd.DataFrame(
            {
                "entity": df[entity_col] if entity_col else entity_value,
                "class": df[leadsheet_col],
                "acct_no": df[acct_col],
                "account": df[account_col],
                "py_balance": pd.to_numeric(df[balance_col], errors="coerce").fillna(0.0),
                "source_file": str(path),
                "source_sheet": str(sheet_name),
                "source_row": (df.index.to_series() + header_row + 2).astype(int),
                "source_entity_column": _column_letter(columns, entity_col) if entity_col else "",
                "source_class_column": _column_letter(columns, leadsheet_col),
                "source_acct_column": _column_letter(columns, acct_col),
                "source_account_column": _column_letter(columns, account_col),
                "source_balance_column": _column_letter(columns, balance_col),
            }
        )
        out["entity"] = out["entity"].fillna(entity_value).astype(str).str.strip()
        out["class"] = out["class"].fillna("").astype(str).str.strip()
        out["acct_no"] = out["acct_no"].map(standard_account_number)
        out["account"] = out["account"].fillna("").astype(str).str.strip()
        out["match_key"] = out["account"].map(normalize_match_text)
        out = out[(out["class"] != "") & (out["acct_no"] != "") & (out["account"] != "")].copy()
        frames.append(out)

    if not frames:
        return pd.DataFrame(columns=PRIOR_YEAR_COLUMNS)
    prior_df = pd.concat(frames, ignore_index=True)
    prior_df.insert(0, "prior_row", range(1, len(prior_df) + 1))
    return prior_df


def read_review_tb(path: Path, sheet_name: str | int | None = None) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name or 0, header=None)
    header_row = _find_header_row(raw)
    df = pd.read_excel(path, sheet_name=sheet_name or 0, header=header_row)

    entity_col = _find_named_column(df, "entity", required=False)
    leadsheet_col = _find_named_column(df, "leadsheet")
    acct_col = _find_named_column(df, "account_number")
    account_col = _find_named_column(df, "account_name")
    current_col = None
    for column in df.columns:
        normalized = _normalized_header(column)
        if normalized in {"current year prelim", "current year", "current balance", "final"}:
            current_col = column
            break
    if current_col is None:
        raise ValueError(f"Could not find a current-year balance column in '{path.name}'.")

    out = pd.DataFrame(
        {
            "entity": df[entity_col] if entity_col else "",
            "class": df[leadsheet_col],
            "acct_no": df[acct_col].map(standard_account_number),
            "account": df[account_col],
            "cy_balance": pd.to_numeric(df[current_col], errors="coerce").fillna(0.0),
        }
    )
    out["entity"] = out["entity"].fillna("").astype(str).str.strip()
    out["class"] = out["class"].fillna("").astype(str).str.strip()
    out["account"] = out["account"].fillna("").astype(str).str.strip()
    return out[(out["acct_no"] != "") & (out["account"] != "")].copy()
