from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from trial_balance_pipeline import WorkbookSpec, build_from_workbooks, load_client_config
from trial_balance_pipeline.reporting import ACCOUNTING_NUMBER_FORMAT, format_outputs, write_import_workbook


BASE_DIR = Path(__file__).resolve().parent


def _sample_current_specs() -> list[WorkbookSpec]:
    return [
        WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "ORC TB.xlsx"),
        WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "ORE TB.xlsx"),
        WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "HR TB.xlsx"),
    ]


def _sample_prior_specs() -> list[WorkbookSpec]:
    return [
        WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "Obra Ramos Construction-Audit - 2022-Trial-Balance.xlsx"),
        WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "H&R Equipment-Tax - 2022-Trial-Balance (2).xlsx"),
    ]


def test_sample_build_keeps_carryforward_rows_in_final_import() -> None:
    cfg = load_client_config(BASE_DIR / "docs" / "client_config_template.json")
    result = build_from_workbooks(
        current_specs=_sample_current_specs(),
        prior_specs=_sample_prior_specs(),
        client_config=cfg,
        entity_overrides=cfg.entity_overrides,
        keep_zero_rows=True,
    )

    carryforward_output = result.updated_import[result.updated_import["source_kind"] == "carryforward_prior_year"]

    assert result.ready_for_import is True
    assert len(result.review_queue) == 0
    assert len(result.carryforward_rows) > 0
    assert len(carryforward_output) == len(result.carryforward_rows)
    assert len(result.updated_import) == len(result.matched_rows) + len(result.new_rows) + len(result.carryforward_rows)


def test_sample_build_stays_close_to_reference_import() -> None:
    cfg = load_client_config(BASE_DIR / "docs" / "client_config_template.json")
    result = build_from_workbooks(
        current_specs=_sample_current_specs(),
        prior_specs=_sample_prior_specs(),
        client_config=cfg,
        entity_overrides=cfg.entity_overrides,
        keep_zero_rows=True,
    )

    actual = result.updated_import[["entity", "class", "acct_no", "account"]].fillna("").copy()
    expected = pd.read_excel(
        BASE_DIR / "New_Spreadsheets" / "TB Import.xlsx",
        header=None,
        names=["entity", "class", "acct_no", "account", "py_balance", "cy_balance"],
    )[["entity", "class", "acct_no", "account"]].fillna("").copy()

    actual_keys = set(tuple(map(str, row)) for row in actual.itertuples(index=False, name=None))
    expected_keys = set(tuple(map(str, row)) for row in expected.itertuples(index=False, name=None))
    shared = len(actual_keys & expected_keys)

    assert shared >= 539


def test_import_workbook_uses_accounting_style_for_zero_and_negative_balances() -> None:
    df = pd.DataFrame(
        [
            {
                "entity": "Demo Co",
                "class": "1000",
                "acct_no": "1001",
                "account": "Cash",
                "py_balance": 125.0,
                "cy_balance": 0.0,
            },
            {
                "entity": "Demo Co",
                "class": "2000",
                "acct_no": "2001",
                "account": "Payables",
                "py_balance": -250.5,
                "cy_balance": 300.0,
            },
        ]
    )
    out_path = BASE_DIR / "tmp_formatted_import.xlsx"
    try:
        write_import_workbook(df, out_path)
        format_outputs(out_path, None)

        sheet = load_workbook(out_path).active

        assert sheet["D1"].value == 125.0
        assert sheet["E1"].value == 0.0
        assert sheet["D2"].value == -250.5
        assert sheet["E2"].value == 300.0
        assert sheet["D1"].number_format == ACCOUNTING_NUMBER_FORMAT
        assert sheet["E1"].number_format == ACCOUNTING_NUMBER_FORMAT
        assert sheet["D2"].number_format == ACCOUNTING_NUMBER_FORMAT
        assert sheet["E2"].number_format == ACCOUNTING_NUMBER_FORMAT
    finally:
        out_path.unlink(missing_ok=True)
