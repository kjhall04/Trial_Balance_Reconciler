from __future__ import annotations

import argparse
from pathlib import Path

from .config import load_client_config, parse_entity_overrides
from .models import WorkbookSpec
from .reporting import format_outputs, write_details_workbook, write_import_workbook, write_review_workbook
from .workflow import build_from_workbooks


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--current-workbook", action="append", default=[], help="Current-year raw client workbook. Repeat for multi-entity jobs.")
    parser.add_argument(
        "--prior-workbook",
        action="append",
        default=[],
        help="Optional prior-year official trial-balance workbook. Repeat when the job has multiple entities.",
    )
    parser.add_argument("--client-config", type=str, default="", help="Optional JSON setup file with workbook parser and entity rules.")
    parser.add_argument("--entity-override", action="append", default=[], help="Optional KEY=Entity Name override.")
    parser.add_argument("--out-import", type=str, default="tb_to_import_updated.xlsx")
    parser.add_argument("--out-details", type=str, default="tb_build_details.xlsx")
    parser.add_argument("--out-review", type=str, default="tb_review_style.xlsx")
    parser.add_argument("--write-review", action="store_true")
    parser.add_argument("--review-title", type=str, default="Review Trial Balance")
    parser.add_argument("--keep-zero-rows", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.current_workbook:
        raise SystemExit("At least one --current-workbook path is required.")

    client_config = load_client_config(args.client_config) if str(args.client_config).strip() else None
    entity_overrides = parse_entity_overrides("\n".join(args.entity_override))

    result = build_from_workbooks(
        current_specs=[WorkbookSpec(Path(path).expanduser()) for path in args.current_workbook],
        prior_specs=[WorkbookSpec(Path(path).expanduser()) for path in args.prior_workbook],
        entity_overrides=entity_overrides,
        client_config=client_config,
        keep_zero_rows=args.keep_zero_rows,
    )

    out_import = Path(args.out_import).expanduser()
    out_details = Path(args.out_details).expanduser()
    write_import_workbook(result.updated_import, out_import)
    write_details_workbook(result, out_details)

    out_review = None
    if args.write_review:
        out_review = Path(args.out_review).expanduser()
        write_review_workbook(result.updated_import, out_review, args.review_title)

    format_outputs(out_import, out_details, out_review)
    print(result.summary.to_string(index=False))
    if not result.ready_for_import:
        print("\nManual review is still required. See the review_queue sheet in the details workbook.")
    elif not args.prior_workbook:
        print("\nNo prior-year official TB workbook was provided, so the build used current-year data only.")
