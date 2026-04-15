from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .models import TrialBalanceBuildResult
from .normalize import (
    account_family,
    account_suffix,
    clean_text,
    is_zero_balance,
    make_account_number,
    normalize_match_text,
    similarity,
    standard_account_number,
    token_overlap,
)


@dataclass(frozen=True)
class MatchCandidate:
    prior_row: int
    score: float
    account: str
    acct_no: str
    family: str
    leadsheet: str


KEYWORD_CLASS_RULES = {
    "property tax": "Property Taxes",
    "office lease": "Rent Expense",
    "building repairs": "Rent Expense",
    "costs and estimated earnings in excess of billings": "Deficit billings",
    "billings in excess of costs and estimated earnings": "Billings in excess of cost",
    "interest paid": "Interest Expense",
    "legal": "Professional fees",
    "accounting": "Professional fees",
    "cpa": "Professional fees",
    "payroll processing": "Payroll Expenses",
    "reconciliation discrep": "Other operating expenses",
    "other income": "Other Income",
    "non business income": "Other Income",
    "intercompany": "Members Equity",
    "equity": "Members Equity",
    "short term debt": "Short Term Debt",
    "lease liability": "Operating Lease Liabilities, less current portion",
    "field labor": "Direct Labor",
    "joint payments": "Cash and Cash equivalents",
    "payroll tax liability": "Payroll Taxes Payable",
    "direct deposit liabilities": "Payroll Taxes Payable",
    "to be reclassed": "Other direct Costs",
    "travel": "Other Operating Expense",
    "first united operating": "Cash and Cash equivalents",
    "balancing accounts": "Cash and Cash equivalents",
    "line of credit": "Notes Payable",
    "lo credit": "Notes Payable",
    "equipment lease": "Interest Expense",
    "office supply": "Office Expense",
    "fixed assets": "Fixed Assets",
    "business vehicles": "Fixed Assets",
    "equipment machinery": "Fixed Assets",
}

KEYWORD_FAMILY_RULES = {
    "to be reclassed": "5660",
    "reconciliation discrep": "8040",
    "joint payments": "1050",
    "costs and estimated earnings in excess of billings": "1450",
    "billings in excess of costs and estimated earnings": "2340",
}

REVIEW_PLACEHOLDER_CLASS = "Needs Review"
REVIEW_PLACEHOLDER_FAMILY = "9999"


def _candidate_score(current_row: pd.Series, prior_row: pd.Series) -> float:
    name_score = similarity(current_row["account_name"], prior_row["account"])
    overlap_score = token_overlap(current_row["account_name"], prior_row["account"])
    score = max(name_score, (name_score * 0.7) + (overlap_score * 0.3))

    current_number = standard_account_number(current_row.get("raw_account_number"))
    prior_number = standard_account_number(prior_row["acct_no"])
    if current_number and current_number == prior_number:
        score += 0.35
    elif account_family(current_number) and account_family(current_number) == account_family(prior_number):
        score += 0.08

    parent_number = standard_account_number(current_row.get("parent_account_number"))
    if parent_number and account_family(parent_number) == account_family(prior_number):
        score += 0.05

    if normalize_match_text(current_row["entity"]) == normalize_match_text(prior_row["entity"]):
        score += 0.05

    return min(score, 1.4)


def _family_class_lookup(
    prior_df: pd.DataFrame,
) -> tuple[dict[str, str], dict[str, str], dict[str, str], dict[str, str]]:
    family_to_class: dict[str, str] = {}
    family_to_name: dict[str, str] = {}
    prefix3_to_class: dict[str, str] = {}
    prefix2_to_class: dict[str, str] = {}
    grouped = prior_df.copy()
    grouped["family"] = grouped["acct_no"].map(account_family)
    for family, frame in grouped.groupby("family"):
        classes = Counter(frame["class"].tolist())
        names = Counter(frame["account"].tolist())
        family_to_class[family] = classes.most_common(1)[0][0]
        family_to_name[family] = names.most_common(1)[0][0]
    for prefix_len, target in ((3, prefix3_to_class), (2, prefix2_to_class)):
        grouped[f"prefix_{prefix_len}"] = grouped["family"].astype(str).str[:prefix_len]
        for prefix, frame in grouped.groupby(f"prefix_{prefix_len}"):
            if not clean_text(prefix):
                continue
            classes = Counter(frame["class"].tolist())
            target[str(prefix)] = classes.most_common(1)[0][0]
    return family_to_class, family_to_name, prefix3_to_class, prefix2_to_class


def _find_entity_matches(current_df: pd.DataFrame, prior_df: pd.DataFrame) -> tuple[dict[int, int], set[int]]:
    matches: dict[int, int] = {}
    used_prior_rows: set[int] = set()
    candidates: list[tuple[float, int, int]] = []

    for _, current_row in current_df.iterrows():
        entity_prior = prior_df[
            prior_df["entity"].map(normalize_match_text) == normalize_match_text(current_row["entity"])
        ]
        for _, prior_row in entity_prior.iterrows():
            score = _candidate_score(current_row, prior_row)
            if score >= 0.86 or (score >= 0.55 and standard_account_number(current_row["raw_account_number"]) == prior_row["acct_no"]):
                candidates.append((score, int(current_row["current_row"]), int(prior_row["prior_row"])))

    candidates.sort(reverse=True)
    for score, current_row_no, prior_row_no in candidates:
        if current_row_no in matches or prior_row_no in used_prior_rows:
            continue
        matches[current_row_no] = prior_row_no
        used_prior_rows.add(prior_row_no)
    return matches, used_prior_rows


def _best_global_candidate(current_row: pd.Series, prior_df: pd.DataFrame) -> MatchCandidate | None:
    best: MatchCandidate | None = None
    for _, prior_row in prior_df.iterrows():
        score = _candidate_score(current_row, prior_row)
        if best is None or score > best.score:
            best = MatchCandidate(
                prior_row=int(prior_row["prior_row"]),
                score=score,
                account=prior_row["account"],
                acct_no=prior_row["acct_no"],
                family=account_family(prior_row["acct_no"]),
                leadsheet=prior_row["class"],
            )
    if best and best.score >= 0.65:
        return best
    return None


def _infer_class(
    name: str,
    family: str,
    family_to_class: dict[str, str],
    prefix3_to_class: dict[str, str],
    prefix2_to_class: dict[str, str],
) -> str:
    lowered = normalize_match_text(name)
    for keyword, leadsheet in KEYWORD_CLASS_RULES.items():
        if keyword in lowered:
            return leadsheet
    if family and family in family_to_class:
        return family_to_class[family]
    if family[:3] and family[:3] in prefix3_to_class:
        return prefix3_to_class[family[:3]]
    if family[:2] and family[:2] in prefix2_to_class:
        return prefix2_to_class[family[:2]]
    return ""


def _preferred_family(current_row: pd.Series, global_candidate: MatchCandidate | None) -> str:
    raw_number = standard_account_number(current_row["raw_account_number"])
    parent_number = standard_account_number(current_row["parent_account_number"])
    raw_family = account_family(raw_number)
    parent_family = account_family(parent_number)
    candidate_family = global_candidate.family if global_candidate else ""

    if candidate_family:
        if raw_family and raw_family == candidate_family:
            return raw_family
        if parent_family and parent_family == candidate_family:
            return parent_family
        if parent_family and not raw_family:
            return parent_family
        if raw_family and not parent_family:
            return raw_family
        return candidate_family
    inferred = raw_family or parent_family
    if inferred:
        return inferred
    lowered = normalize_match_text(current_row["account_name"])
    for keyword, family in KEYWORD_FAMILY_RULES.items():
        if keyword in lowered:
            return family
    return ""


def _preferred_account_name(current_row: pd.Series, global_candidate: MatchCandidate | None) -> str:
    if global_candidate and global_candidate.score >= 0.92:
        return global_candidate.account
    return current_row["account_name"]


def _current_priority(current_row: pd.Series, global_candidate: MatchCandidate | None) -> tuple[int, bool, bool]:
    raw_number = standard_account_number(current_row["raw_account_number"])
    is_rollup = account_suffix(raw_number) == "00" or int(current_row["path_depth"]) <= 2
    can_displace = False
    priority = 70
    if global_candidate and global_candidate.score >= 0.95:
        priority = 85
        can_displace = True
    elif global_candidate and global_candidate.score >= 0.82:
        priority = 78
        can_displace = True
    if is_rollup:
        priority += 20
        can_displace = True
    return priority, is_rollup, can_displace


def _next_available_number(family: str, occupied: set[str]) -> str:
    for suffix in range(0, 10_000):
        candidate = make_account_number(family, suffix)
        if candidate not in occupied:
            return candidate
    raise ValueError(f"Could not allocate an account number in family '{family}'.")


def _sort_key(value: str) -> tuple[int, int]:
    text = standard_account_number(value)
    family = account_family(text)
    suffix = account_suffix(text)
    try:
        return int(family), int(suffix)
    except ValueError:
        return (999999, 999999)


def _safe_int(value: object) -> int | None:
    if value is None or pd.isna(value):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _append_sentence(existing: object, sentence: str) -> str:
    existing_text = clean_text(existing)
    if not existing_text:
        return sentence
    return f"{existing_text.rstrip('.')} {sentence}"


def _source_reference(file_path: object, sheet_name: object, row_number: object, *columns: object) -> str:
    parts: list[str] = []
    file_text = clean_text(file_path)
    if file_text:
        try:
            parts.append(Path(file_text).name)
        except OSError:
            parts.append(file_text)

    location_parts: list[str] = []
    sheet_text = clean_text(sheet_name)
    if sheet_text:
        location_parts.append(f"sheet {sheet_text}")
    source_row = _safe_int(row_number)
    if source_row is not None:
        location_parts.append(f"row {source_row}")
    source_columns = [clean_text(column) for column in columns if clean_text(column)]
    if source_columns:
        location_parts.append(f"col {'/'.join(source_columns)}")
    if location_parts:
        parts.append(" | ".join(location_parts))
    return " | ".join(parts)


def _current_source_metadata(current_row: pd.Series) -> dict[str, object]:
    source_row = _safe_int(current_row.get("source_row"))
    return {
        "current_source_file": clean_text(current_row.get("source_file")),
        "current_source_sheet": clean_text(current_row.get("source_sheet")),
        "current_source_row": source_row,
        "current_source_text_column": clean_text(current_row.get("source_text_column")),
        "current_source_number_column": clean_text(current_row.get("source_number_column")),
        "current_source_amount_column": clean_text(current_row.get("source_amount_column")),
        "current_source_reference": _source_reference(
            current_row.get("source_file"),
            current_row.get("source_sheet"),
            source_row,
            current_row.get("source_text_column"),
            current_row.get("source_number_column"),
            current_row.get("source_amount_column"),
        ),
    }


def _blank_prior_source_metadata() -> dict[str, object]:
    return {
        "matched_prior_entity": "",
        "matched_prior_account": "",
        "matched_prior_file": "",
        "matched_prior_sheet": "",
        "matched_prior_row": None,
        "matched_prior_entity_column": "",
        "matched_prior_class_column": "",
        "matched_prior_acct_column": "",
        "matched_prior_account_column": "",
        "matched_prior_balance_column": "",
        "matched_prior_reference": "",
    }


def _prior_source_metadata(prior_row: pd.Series | None) -> dict[str, object]:
    if prior_row is None:
        return _blank_prior_source_metadata()
    source_row = _safe_int(prior_row.get("source_row"))
    return {
        "matched_prior_entity": clean_text(prior_row.get("entity")),
        "matched_prior_account": clean_text(prior_row.get("account")),
        "matched_prior_file": clean_text(prior_row.get("source_file")),
        "matched_prior_sheet": clean_text(prior_row.get("source_sheet")),
        "matched_prior_row": source_row,
        "matched_prior_entity_column": clean_text(prior_row.get("source_entity_column")),
        "matched_prior_class_column": clean_text(prior_row.get("source_class_column")),
        "matched_prior_acct_column": clean_text(prior_row.get("source_acct_column")),
        "matched_prior_account_column": clean_text(prior_row.get("source_account_column")),
        "matched_prior_balance_column": clean_text(prior_row.get("source_balance_column")),
        "matched_prior_reference": _source_reference(
            prior_row.get("source_file"),
            prior_row.get("source_sheet"),
            source_row,
            prior_row.get("source_class_column"),
            prior_row.get("source_acct_column"),
            prior_row.get("source_account_column"),
            prior_row.get("source_balance_column"),
        ),
    }


def _normalized_confidence_score(score: object) -> float:
    try:
        return round(max(0.0, min(float(score), 1.0)), 2)
    except (TypeError, ValueError):
        return 0.0


def _confidence_for_candidate(
    *,
    source_kind: str,
    match_score: float,
    family: str,
    leadsheet: str,
    raw_account_number: str,
    desired_acct_no: str,
    has_global_candidate: bool,
) -> tuple[str, float, str]:
    normalized_score = _normalized_confidence_score(match_score)
    raw_number = standard_account_number(raw_account_number)
    desired_number = standard_account_number(desired_acct_no)

    if source_kind == "matched_prior_year":
        return "high", max(normalized_score, 0.98), "Matched directly to the prior-year trial balance."
    if source_kind == "carryforward_prior_year":
        return "high", 0.95, "No current-year row matched, so the prior-year line was carried forward."
    if not family or not leadsheet:
        return "low", min(normalized_score, 0.35), "The program could not infer a stable family and leadsheet."
    if has_global_candidate and match_score >= 0.98:
        return "high", normalized_score, "Very strong similarity to a prior-year account."
    if has_global_candidate and match_score >= 0.72:
        return "medium", max(normalized_score, 0.72), "Assigned from a similar prior-year account."
    if raw_number and desired_number and account_family(raw_number) == account_family(desired_number):
        return "medium", max(normalized_score, 0.68), "Used the current-year account family to assign the leadsheet."
    if desired_number:
        return "low", max(normalized_score, 0.45), "Account number or leadsheet was created from limited evidence."
    return "low", 0.25, "This row needs manual review."


def _mark_renumbered(row: dict[str, object], reason: str) -> None:
    row["was_renumbered"] = True
    row["confidence_level"] = "low"
    row["confidence_score"] = min(float(row.get("confidence_score", 0.45) or 0.45), 0.45)
    row["confidence_reason"] = _append_sentence(
        row.get("confidence_reason"),
        "Account number changed to avoid a numbering conflict.",
    )
    row["decision_note"] = _append_sentence(row.get("decision_note"), reason)


def _ensure_review_placeholder_assignment(row: dict[str, object]) -> None:
    if not clean_text(row.get("class")):
        row["class"] = REVIEW_PLACEHOLDER_CLASS
        row["confidence_level"] = "low"
        row["confidence_score"] = min(float(row.get("confidence_score", 0.35) or 0.35), 0.35)
        row["confidence_reason"] = _append_sentence(
            row.get("confidence_reason"),
            "Assigned a placeholder leadsheet because this row still needs manual review.",
        )
        row["decision_note"] = _append_sentence(
            row.get("decision_note"),
            "Assigned the placeholder leadsheet 'Needs Review' to keep the row in the output.",
        )

    if not clean_text(row.get("family")):
        row["family"] = REVIEW_PLACEHOLDER_FAMILY
        row["confidence_level"] = "low"
        row["confidence_score"] = min(float(row.get("confidence_score", 0.35) or 0.35), 0.35)
        row["confidence_reason"] = _append_sentence(
            row.get("confidence_reason"),
            "Account family could not be inferred from the current-year workbook alone.",
        )
        row["decision_note"] = _append_sentence(
            row.get("decision_note"),
            "Assigned a placeholder account family so the row stays in the output for review.",
        )


def build_trial_balance(
    current_df: pd.DataFrame,
    prior_df: pd.DataFrame,
    *,
    keep_zero_rows: bool = True,
) -> TrialBalanceBuildResult:
    if current_df.empty:
        raise ValueError("No current-year rows were found.")

    prior_working = prior_df.copy()
    prior_working["family"] = prior_working["acct_no"].map(account_family)
    family_to_class, family_to_name, prefix3_to_class, prefix2_to_class = _family_class_lookup(prior_working)

    entity_matches, used_prior_rows = _find_entity_matches(current_df, prior_working)

    candidate_rows: list[dict[str, object]] = []
    carryforward_records: list[dict[str, object]] = []
    review_queue_records: list[dict[str, object]] = []

    for _, current_row in current_df.iterrows():
        current_row_no = int(current_row["current_row"])
        if current_row_no in entity_matches:
            prior_row = prior_working.loc[prior_working["prior_row"] == entity_matches[current_row_no]].iloc[0]
            match_score = _candidate_score(current_row, prior_row)
            confidence_level, confidence_score, confidence_reason = _confidence_for_candidate(
                source_kind="matched_prior_year",
                match_score=match_score,
                family=account_family(prior_row["acct_no"]),
                leadsheet=prior_row["class"],
                raw_account_number=standard_account_number(current_row["raw_account_number"]),
                desired_acct_no=prior_row["acct_no"],
                has_global_candidate=True,
            )
            candidate_rows.append(
                {
                    "row_type": "current",
                    "source_kind": "matched_prior_year",
                    "status": "matched to prior year",
                    "entity": current_row["entity"],
                    "class": prior_row["class"],
                    "account": prior_row["account"],
                    "py_balance": float(prior_row["py_balance"]),
                    "cy_balance": float(current_row["cy_balance"]),
                    "desired_acct_no": prior_row["acct_no"],
                    "family": account_family(prior_row["acct_no"]),
                    "current_row": current_row_no,
                    "prior_row": int(prior_row["prior_row"]),
                    "raw_account_text": current_row["raw_account_text"],
                    "raw_account_number": standard_account_number(current_row["raw_account_number"]),
                    "path_depth": int(current_row["path_depth"]),
                    "priority": 95,
                    "can_displace": True,
                    "prefers_base": False,
                    "match_score": match_score,
                    "confidence_level": confidence_level,
                    "confidence_score": confidence_score,
                    "confidence_reason": confidence_reason,
                    "decision_note": "Reused the prior-year leadsheet and account number.",
                    "was_renumbered": False,
                    **_current_source_metadata(current_row),
                    **_prior_source_metadata(prior_row),
                }
            )
            continue

        global_candidate = _best_global_candidate(current_row, prior_working)
        global_prior_row = (
            prior_working.loc[prior_working["prior_row"] == global_candidate.prior_row].iloc[0]
            if global_candidate is not None
            else None
        )
        family = _preferred_family(current_row, global_candidate)
        inferred_class = _infer_class(
            current_row["account_name"],
            family,
            family_to_class,
            prefix3_to_class,
            prefix2_to_class,
        )
        leadsheet = (
            inferred_class
            if inferred_class and (not global_candidate or global_candidate.score < 0.9)
            else (global_candidate.leadsheet if global_candidate and global_candidate.score >= 0.72 else inferred_class)
        )
        account_name = _preferred_account_name(current_row, global_candidate)
        priority, is_rollup, can_displace = _current_priority(current_row, global_candidate)

        desired_acct_no = ""
        if global_candidate and global_candidate.score >= 0.9:
            desired_acct_no = global_candidate.acct_no
        else:
            raw_number = standard_account_number(current_row["raw_account_number"])
            desired_acct_no = raw_number if raw_number else (make_account_number(family, "00") if family else "")

        confidence_level, confidence_score, confidence_reason = _confidence_for_candidate(
            source_kind="new_current_year",
            match_score=global_candidate.score if global_candidate else 0.0,
            family=family,
            leadsheet=leadsheet,
            raw_account_number=standard_account_number(current_row["raw_account_number"]),
            desired_acct_no=desired_acct_no,
            has_global_candidate=global_candidate is not None,
        )
        if global_candidate and global_candidate.score >= 0.98:
            decision_note = "Very strong prior-year similarity guided the leadsheet and account number."
        elif global_candidate and global_candidate.score >= 0.72:
            decision_note = "A similar prior-year account guided the leadsheet and account number."
        elif family and leadsheet and standard_account_number(current_row["raw_account_number"]):
            decision_note = "Used the current-year account family to keep numbering aligned."
        elif family and leadsheet:
            decision_note = "Created a new account number from the inferred family."
        else:
            decision_note = "This row needs manual review before import."

        candidate_rows.append(
            {
                "row_type": "current",
                "source_kind": "new_current_year",
                "status": "new current-year row",
                "entity": current_row["entity"],
                "class": leadsheet,
                "account": account_name,
                "py_balance": 0.0,
                "cy_balance": float(current_row["cy_balance"]),
                "desired_acct_no": desired_acct_no,
                "family": family,
                "current_row": current_row_no,
                "prior_row": None,
                "raw_account_text": current_row["raw_account_text"],
                "raw_account_number": standard_account_number(current_row["raw_account_number"]),
                "path_depth": int(current_row["path_depth"]),
                "priority": priority,
                "can_displace": can_displace,
                "prefers_base": bool(is_rollup),
                "match_score": global_candidate.score if global_candidate else 0.0,
                "confidence_level": confidence_level,
                "confidence_score": confidence_score,
                "confidence_reason": confidence_reason,
                "decision_note": decision_note,
                "was_renumbered": False,
                **_current_source_metadata(current_row),
                **_prior_source_metadata(global_prior_row),
            }
        )
        if not family or not leadsheet:
            review_queue_records.append(
                {
                    "entity": current_row["entity"],
                    "status": "needs review",
                    "source_kind": "review_queue",
                    "current_row": current_row_no,
                    "prior_row": int(global_candidate.prior_row) if global_candidate else None,
                    "raw_account_text": current_row["raw_account_text"],
                    "raw_account_number": standard_account_number(current_row["raw_account_number"]),
                    "assigned_account": account_name,
                    "assigned_acct_no": desired_acct_no,
                    "assigned_class": leadsheet,
                    "py_balance": 0.0,
                    "cy_balance": float(current_row["cy_balance"]),
                    "family": family,
                    "match_score": global_candidate.score if global_candidate else 0.0,
                    "confidence_level": "low",
                    "confidence_score": min(confidence_score, 0.35),
                    "confidence_reason": "The program could not infer a stable family and leadsheet.",
                    "decision_note": "This line needs a manual leadsheet and account-number review before import.",
                    "reason": "The new pipeline could not infer a stable family or leadsheet for this row.",
                    "was_renumbered": False,
                    **_current_source_metadata(current_row),
                    **_prior_source_metadata(global_prior_row),
                }
            )

    for _, prior_row in prior_working[~prior_working["prior_row"].isin(used_prior_rows)].iterrows():
        confidence_level, confidence_score, confidence_reason = _confidence_for_candidate(
            source_kind="carryforward_prior_year",
            match_score=1.0,
            family=prior_row["family"],
            leadsheet=prior_row["class"],
            raw_account_number=prior_row["acct_no"],
            desired_acct_no=prior_row["acct_no"],
            has_global_candidate=True,
        )
        carryforward_records.append(
            {
                "row_type": "carryforward",
                "source_kind": "carryforward_prior_year",
                "status": "carried forward from prior year",
                "entity": prior_row["entity"],
                "class": prior_row["class"],
                "account": prior_row["account"],
                "py_balance": float(prior_row["py_balance"]),
                "cy_balance": 0.0,
                "assigned_acct_no": prior_row["acct_no"],
                "assigned_class": prior_row["class"],
                "assigned_account": prior_row["account"],
                "family": prior_row["family"],
                "current_row": None,
                "prior_row": int(prior_row["prior_row"]),
                "raw_account_text": "",
                "raw_account_number": prior_row["acct_no"],
                "match_score": 1.0,
                "confidence_level": confidence_level,
                "confidence_score": confidence_score,
                "confidence_reason": confidence_reason,
                "decision_note": "No current-year row matched, so the prior-year line was carried forward.",
                "was_renumbered": False,
                **_current_source_metadata(pd.Series(dtype=object)),
                **_prior_source_metadata(prior_row),
            }
        )

    assignments: list[dict] = []
    renumbered_records: list[dict] = []
    for entity, entity_rows_frame in pd.DataFrame(candidate_rows).groupby("entity"):
        entity_rows = entity_rows_frame.to_dict("records")
        occupied: dict[str, dict] = {}
        pending = sorted(
            entity_rows,
            key=lambda row: (
                row["priority"],
                1 if row["prefers_base"] else 0,
                -row["path_depth"],
                row["match_score"],
            ),
            reverse=True,
        )

        while pending:
            row = pending.pop(0)
            _ensure_review_placeholder_assignment(row)
            family = clean_text(row["family"])

            desired = standard_account_number(row["desired_acct_no"])
            if not desired:
                desired = make_account_number(family, "00")
            if row["prefers_base"] and not any(assigned["entity"] == entity and assigned["family"] == family and assigned["acct_no"].endswith("-00") for assigned in assignments):
                desired = make_account_number(family, "00")

            occupant = occupied.get(desired)
            if occupant is None:
                row["acct_no"] = desired
                occupied[desired] = row
                assignments.append(row)
                continue

            if row["can_displace"] and row["priority"] > occupant["priority"]:
                fallback_number = _next_available_number(family, set(occupied))
                occupant["acct_no"] = fallback_number
                occupant_reason = f"Account number changed to {fallback_number} because {row['account']} claimed {desired}."
                _mark_renumbered(occupant, occupant_reason)
                renumbered_records.append(
                    {
                        "entity": entity,
                        "account": occupant["account"],
                        "old_acct_no": desired,
                        "new_acct_no": fallback_number,
                        "reason": f"{row['account']} claimed {desired} in the fresh current-year build.",
                        "source_kind": occupant["source_kind"],
                        "current_row": occupant.get("current_row"),
                        "prior_row": occupant.get("prior_row"),
                        "confidence_level": occupant["confidence_level"],
                    }
                )
                occupied[fallback_number] = occupant
                assignments = [assigned if assigned is not occupant else occupant for assigned in assignments]
                row["acct_no"] = desired
                occupied[desired] = row
                assignments.append(row)
                continue

            fallback_number = _next_available_number(family, set(occupied))
            row["acct_no"] = fallback_number
            if fallback_number != desired:
                row_reason = f"Account number changed to {fallback_number} because {desired} was already occupied in this entity/family."
                _mark_renumbered(row, row_reason)
                renumbered_records.append(
                    {
                        "entity": entity,
                        "account": row["account"],
                        "old_acct_no": desired,
                        "new_acct_no": fallback_number,
                        "reason": f"{desired} was already occupied in this entity/family.",
                        "source_kind": row["source_kind"],
                        "current_row": row.get("current_row"),
                        "prior_row": row.get("prior_row"),
                        "confidence_level": row["confidence_level"],
                    }
                )
            occupied[fallback_number] = row
            assignments.append(row)

    output_df = pd.DataFrame(assignments)
    if output_df.empty:
        raise ValueError("No output rows were built.")

    output_df = output_df[
        [
            "entity",
            "class",
            "acct_no",
            "account",
            "py_balance",
            "cy_balance",
            "status",
            "source_kind",
            "family",
            "current_row",
            "prior_row",
            "raw_account_text",
            "raw_account_number",
            "match_score",
            "confidence_level",
            "confidence_score",
            "confidence_reason",
            "decision_note",
            "was_renumbered",
            "current_source_file",
            "current_source_sheet",
            "current_source_row",
            "current_source_text_column",
            "current_source_number_column",
            "current_source_amount_column",
            "current_source_reference",
            "matched_prior_entity",
            "matched_prior_account",
            "matched_prior_file",
            "matched_prior_sheet",
            "matched_prior_row",
            "matched_prior_entity_column",
            "matched_prior_class_column",
            "matched_prior_acct_column",
            "matched_prior_account_column",
            "matched_prior_balance_column",
            "matched_prior_reference",
        ]
    ].copy()
    if not keep_zero_rows:
        output_df = output_df[~output_df["cy_balance"].apply(is_zero_balance)].copy()
    output_df = output_df.sort_values(["entity", "acct_no", "account"], key=lambda series: series.map(_sort_key) if series.name == "acct_no" else series).reset_index(drop=True)
    output_df.attrs["multi_entity"] = output_df["entity"].astype(str).str.strip().nunique() > 1

    renumbered_df = pd.DataFrame(
        renumbered_records,
        columns=[
            "entity",
            "account",
            "old_acct_no",
            "new_acct_no",
            "reason",
            "source_kind",
            "current_row",
            "prior_row",
            "confidence_level",
        ],
    )
    matched_df = output_df[output_df["source_kind"] == "matched_prior_year"][
        [
            "entity",
            "current_row",
            "prior_row",
            "class",
            "acct_no",
            "account",
            "py_balance",
            "cy_balance",
            "match_score",
            "confidence_level",
            "confidence_score",
            "confidence_reason",
            "decision_note",
            "current_source_reference",
            "matched_prior_reference",
        ]
    ].copy()
    new_rows_df = output_df[output_df["source_kind"] == "new_current_year"][
        [
            "entity",
            "current_row",
            "prior_row",
            "raw_account_text",
            "raw_account_number",
            "class",
            "acct_no",
            "account",
            "py_balance",
            "cy_balance",
            "family",
            "match_score",
            "confidence_level",
            "confidence_score",
            "confidence_reason",
            "decision_note",
            "current_source_reference",
            "matched_prior_reference",
        ]
    ].copy()
    carryforward_details_df = pd.DataFrame(carryforward_records)
    if carryforward_details_df.empty:
        carryforward_df = pd.DataFrame(
            columns=[
                "entity",
                "prior_row",
                "assigned_class",
                "assigned_acct_no",
                "assigned_account",
                "py_balance",
                "cy_balance",
                "confidence_level",
                "confidence_score",
                "confidence_reason",
                "decision_note",
                "matched_prior_reference",
            ]
        )
    else:
        carryforward_df = carryforward_details_df[
            [
                "entity",
                "prior_row",
                "assigned_class",
                "assigned_acct_no",
                "assigned_account",
                "py_balance",
                "cy_balance",
                "confidence_level",
                "confidence_score",
                "confidence_reason",
                "decision_note",
                "matched_prior_reference",
            ]
        ].copy()
    review_queue_df = pd.DataFrame(review_queue_records)
    if not review_queue_df.empty:
        review_assignments = output_df[
            [
                "entity",
                "current_row",
                "class",
                "acct_no",
                "account",
                "family",
                "confidence_level",
                "confidence_score",
                "confidence_reason",
                "decision_note",
            ]
        ].rename(
            columns={
                "class": "resolved_class",
                "acct_no": "resolved_acct_no",
                "account": "resolved_account",
                "family": "resolved_family",
                "confidence_level": "resolved_confidence_level",
                "confidence_score": "resolved_confidence_score",
                "confidence_reason": "resolved_confidence_reason",
                "decision_note": "resolved_decision_note",
            }
        )
        review_queue_df = review_queue_df.merge(review_assignments, on=["entity", "current_row"], how="left")
        review_queue_df["assigned_class"] = review_queue_df["resolved_class"].fillna(review_queue_df["assigned_class"])
        review_queue_df["assigned_acct_no"] = review_queue_df["resolved_acct_no"].fillna(review_queue_df["assigned_acct_no"])
        review_queue_df["assigned_account"] = review_queue_df["resolved_account"].fillna(review_queue_df["assigned_account"])
        review_queue_df["family"] = review_queue_df["resolved_family"].fillna(review_queue_df["family"])
        review_queue_df["confidence_level"] = review_queue_df["resolved_confidence_level"].fillna(review_queue_df["confidence_level"])
        review_queue_df["confidence_score"] = review_queue_df["resolved_confidence_score"].fillna(review_queue_df["confidence_score"])
        review_queue_df["confidence_reason"] = review_queue_df["resolved_confidence_reason"].fillna(review_queue_df["confidence_reason"])
        review_queue_df["decision_note"] = review_queue_df["resolved_decision_note"].fillna(review_queue_df["decision_note"])
        review_queue_df = review_queue_df.drop(
            columns=[
                "resolved_class",
                "resolved_acct_no",
                "resolved_account",
                "resolved_family",
                "resolved_confidence_level",
                "resolved_confidence_score",
                "resolved_confidence_reason",
                "resolved_decision_note",
            ]
        )
    assigned_audit_df = output_df[
        [
            "entity",
            "status",
            "source_kind",
            "confidence_level",
            "confidence_score",
            "confidence_reason",
            "decision_note",
            "current_row",
            "prior_row",
            "raw_account_text",
            "raw_account_number",
            "class",
            "acct_no",
            "account",
            "py_balance",
            "cy_balance",
            "match_score",
            "family",
            "was_renumbered",
            "current_source_file",
            "current_source_sheet",
            "current_source_row",
            "current_source_text_column",
            "current_source_number_column",
            "current_source_amount_column",
            "current_source_reference",
            "matched_prior_entity",
            "matched_prior_account",
            "matched_prior_file",
            "matched_prior_sheet",
            "matched_prior_row",
            "matched_prior_entity_column",
            "matched_prior_class_column",
            "matched_prior_acct_column",
            "matched_prior_account_column",
            "matched_prior_balance_column",
            "matched_prior_reference",
        ]
    ].rename(
        columns={
            "class": "assigned_class",
            "acct_no": "assigned_acct_no",
            "account": "assigned_account",
        }
    ).copy()
    carryforward_audit_df = carryforward_details_df.copy()
    if not carryforward_audit_df.empty:
        carryforward_audit_df = carryforward_audit_df[
            [
                "entity",
                "status",
                "source_kind",
                "confidence_level",
                "confidence_score",
                "confidence_reason",
                "decision_note",
                "current_row",
                "prior_row",
                "raw_account_text",
                "raw_account_number",
                "assigned_class",
                "assigned_acct_no",
                "assigned_account",
                "py_balance",
                "cy_balance",
                "match_score",
                "family",
                "was_renumbered",
                "current_source_file",
                "current_source_sheet",
                "current_source_row",
                "current_source_text_column",
                "current_source_number_column",
                "current_source_amount_column",
                "current_source_reference",
                "matched_prior_entity",
                "matched_prior_account",
                "matched_prior_file",
                "matched_prior_sheet",
                "matched_prior_row",
                "matched_prior_entity_column",
                "matched_prior_class_column",
                "matched_prior_acct_column",
                "matched_prior_account_column",
                "matched_prior_balance_column",
                "matched_prior_reference",
            ]
        ].copy()
    comparison_columns = assigned_audit_df.columns.tolist()
    if not carryforward_audit_df.empty:
        carryforward_audit_df = carryforward_audit_df.reindex(columns=comparison_columns, fill_value="")
    if not review_queue_df.empty:
        review_queue_df = review_queue_df.reindex(columns=comparison_columns, fill_value="")
    comparison_df = (
        pd.concat(
            [frame for frame in (assigned_audit_df, carryforward_audit_df, review_queue_df) if not frame.empty],
            ignore_index=True,
        )
        if (not assigned_audit_df.empty or not carryforward_audit_df.empty or not review_queue_df.empty)
        else assigned_audit_df.copy()
    )

    confidence_counts = output_df["confidence_level"].value_counts()

    summary = pd.DataFrame(
        [
            {"metric": "current rows parsed", "value": int(len(current_df))},
            {"metric": "prior-year rows parsed", "value": int(len(prior_df))},
            {"metric": "matched to prior year", "value": int(len(matched_df))},
            {"metric": "new current-year rows", "value": int(len(new_rows_df))},
            {"metric": "carryforward prior-year rows", "value": int(len(carryforward_df))},
            {"metric": "renumbered rows", "value": int(len(renumbered_df))},
            {"metric": "review queue rows", "value": int(len(review_queue_df))},
            {"metric": "high-confidence rows", "value": int(confidence_counts.get("high", 0))},
            {"metric": "medium-confidence rows", "value": int(confidence_counts.get("medium", 0))},
            {"metric": "low-confidence rows", "value": int(confidence_counts.get("low", 0))},
            {"metric": "output row count", "value": int(len(output_df))},
            {"metric": "output current-year total", "value": round(float(output_df["cy_balance"].sum()), 2)},
        ]
    )

    return TrialBalanceBuildResult(
        current_trial_balance=current_df.copy(),
        prior_year_rows=prior_df.copy(),
        updated_import=output_df.copy(),
        comparison_details=comparison_df,
        matched_rows=matched_df,
        new_rows=new_rows_df,
        carryforward_rows=carryforward_df,
        renumbered_rows=renumbered_df,
        review_queue=review_queue_df,
        summary=summary,
        ready_for_import=review_queue_df.empty,
    )
