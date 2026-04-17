from __future__ import annotations

from pathlib import Path

import pandas as pd

from trial_balance_pipeline import WorkbookSpec, build_from_workbooks, load_client_config


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
    cfg = load_client_config(BASE_DIR / "docs" / "example_client_config.json")
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
    cfg = load_client_config(BASE_DIR / "docs" / "example_client_config.json")
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
