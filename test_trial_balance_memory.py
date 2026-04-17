from __future__ import annotations

from pathlib import Path

import pandas as pd

from trial_balance_pipeline.assistance import (
    build_memory_client_config,
    build_preflight_items,
    default_user_memory,
    known_entities_from_memory,
    remember_successful_workbooks,
)
from trial_balance_pipeline.current_year import preview_current_workbooks
from trial_balance_pipeline.models import WorkbookPreview, WorkbookSpec


BASE_DIR = Path(__file__).resolve().parent


def test_learned_memory_keeps_distinctive_workbooks_and_skips_generic_names() -> None:
    current_df = pd.DataFrame(
        [
            {
                "source_file": str(BASE_DIR / "New_Spreadsheets" / "ORC TB.xlsx"),
                "source_sheet": "Sheet1",
                "parser_name": "quickbooks_debit_credit",
                "entity": "Obra Ramos Construction, LLC",
            },
            {
                "source_file": str(BASE_DIR / "Old_Spreadsheets" / "client tb.xlsx"),
                "source_sheet": "Sheet1",
                "parser_name": "quickbooks_debit_credit",
                "entity": "APR Group, Inc.",
            },
        ]
    )

    memory = remember_successful_workbooks(default_user_memory(), current_df)

    assert "orc tb.xlsx" in memory["workbook_rules"]
    assert "orc tb" in memory["workbook_rules"]
    assert "client tb.xlsx" not in memory["workbook_rules"]
    assert "client tb" not in memory["workbook_rules"]
    assert "Obra Ramos Construction, LLC" in known_entities_from_memory(memory)
    assert "APR Group, Inc." not in known_entities_from_memory(memory)


def test_memory_client_config_can_reuse_ore_entity_without_prior_files() -> None:
    current_df = pd.DataFrame(
        [
            {
                "source_file": str(BASE_DIR / "New_Spreadsheets" / "ORE TB.xlsx"),
                "source_sheet": "Sheet1",
                "parser_name": "quickbooks_debit_credit",
                "entity": "Obra Ramos Excavation, LLC",
            }
        ]
    )
    memory = remember_successful_workbooks(default_user_memory(), current_df)
    memory_config = build_memory_client_config(memory)
    previews = preview_current_workbooks(
        [WorkbookSpec(BASE_DIR / "New_Spreadsheets" / "ORE TB.xlsx")],
        client_config=memory_config,
        known_entities=known_entities_from_memory(memory),
    )

    assert previews[0].entity == "Obra Ramos Excavation, LLC"
    assert previews[0].entity_source in {"setup file workbook rule", "setup file company match"}


def test_preflight_marks_filename_only_entities_for_confirmation() -> None:
    items = build_preflight_items(
        [
            WorkbookPreview(
                path=BASE_DIR / "New_Spreadsheets" / "ORE TB.xlsx",
                entity="ORE TB",
                entity_source="derived from file name",
                parser_name="quickbooks_debit_credit",
                sheet_name="Sheet1",
            ),
            WorkbookPreview(
                path=BASE_DIR / "New_Spreadsheets" / "HR TB.xlsx",
                entity="H&R Equipment, Inc.",
                entity_source="matched from known companies",
                parser_name="extracted_balance_list",
                sheet_name="Sheet1",
            ),
        ],
        known_entities=["Obra Ramos Construction, LLC", "H&R Equipment, Inc."],
        prior_entities=["Obra Ramos Construction, LLC", "H&R Equipment, Inc."],
    )

    assert items[0].needs_entity_confirmation is True
    assert "remember it next time" in items[0].note.lower()
    assert items[1].needs_entity_confirmation is False


def test_preflight_warns_when_prior_entities_do_not_cover_company() -> None:
    items = build_preflight_items(
        [
            WorkbookPreview(
                path=BASE_DIR / "New_Spreadsheets" / "ORE TB.xlsx",
                entity="Obra Ramos Excavation, LLC",
                entity_source="setup file workbook rule",
                parser_name="quickbooks_debit_credit",
                sheet_name="Sheet1",
            )
        ],
        known_entities=["Obra Ramos Construction, LLC", "Obra Ramos Excavation, LLC"],
        prior_entities=["Obra Ramos Construction, LLC", "H&R Equipment, Inc."],
    )

    assert items[0].needs_entity_confirmation is False
    assert "consolidated official prior-year tb workbook" in items[0].note.lower()
