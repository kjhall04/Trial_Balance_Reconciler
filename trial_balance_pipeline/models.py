from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass(frozen=True)
class WorkbookSpec:
    path: Path
    entity: str = ""


@dataclass(frozen=True)
class WorkbookPreview:
    path: Path
    entity: str = ""
    entity_source: str = ""
    parser_name: str = ""
    parser_source: str = ""
    sheet_name: str | int | None = None
    sheet_source: str = ""
    status: str = "ok"
    note: str = ""


@dataclass(frozen=True)
class WorkbookRule:
    parser: str = ""
    sheet_name: str | int | None = None
    entity: str = ""


@dataclass(frozen=True)
class ClientConfig:
    name: str = ""
    entity_overrides: dict[str, str] = field(default_factory=dict)
    workbook_rules: dict[str, WorkbookRule] = field(default_factory=dict)
    source_path: Path | None = None


@dataclass
class TrialBalanceBuildResult:
    current_trial_balance: pd.DataFrame
    prior_year_rows: pd.DataFrame
    updated_import: pd.DataFrame
    comparison_details: pd.DataFrame
    matched_rows: pd.DataFrame
    new_rows: pd.DataFrame
    carryforward_rows: pd.DataFrame
    renumbered_rows: pd.DataFrame
    review_queue: pd.DataFrame
    summary: pd.DataFrame
    ready_for_import: bool

