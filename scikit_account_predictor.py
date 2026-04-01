"""
Optional scikit-learn account prediction helpers.

This module is intentionally separate from ``guess_category.py`` so the
existing heuristic matcher stays available while we experiment with a trained
model for leadsheet and account suggestions.
"""

from __future__ import annotations

import argparse
import pickle
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Sequence

import pandas as pd

from guess_category import guess_category


DEFAULT_IMPORT_SHEET = "Trial Balance"
ACCOUNT_NUMBER_PATTERN = re.compile(r"^\d{3,6}(?:-\d{1,4})?$")
STOP_WORDS = {
    "and",
    "the",
    "of",
    "to",
    "for",
    "in",
    "on",
    "at",
    "by",
    "from",
    "current",
}
NORMALIZATION_REPLACEMENTS = {
    "&": " and ",
    "a/r": "accounts receivable",
    "a/p": "accounts payable",
    "ltd": "long term debt",
    "n/p": "notes payable",
    "#": " ",
}
HEADER_ALIASES = {
    "entity": {"entity", "entity fund", "entity / fund", "fund"},
    "leadsheet": {"leadsheet"},
    "account_number": {
        "account number",
        "acct no",
        "acct number",
        "account no",
        "account num",
        "account #",
    },
    "account_name": {"account name", "account", "description"},
    "current_balance": {
        "current year prelim",
        "cy balance",
        "current year",
        "current balance",
        "final",
        "per return",
    },
}

SKLEARN_IMPORT_ERROR: ImportError | None = None

try:
    from sklearn.compose import ColumnTransformer
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.metrics.pairwise import linear_kernel
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import OneHotEncoder
except ImportError as exc:  # pragma: no cover - exercised when dependency is absent
    ColumnTransformer = None
    TfidfVectorizer = None
    LogisticRegression = None
    linear_kernel = None
    Pipeline = None
    OneHotEncoder = None
    SKLEARN_IMPORT_ERROR = exc


@dataclass(frozen=True)
class TrainingRow:
    account_name: str
    leadsheet: str
    account_number: str
    entity: str = ""
    current_balance: float | None = None


@dataclass(frozen=True)
class LeadsheetPrediction:
    leadsheet: str
    confidence: float


@dataclass(frozen=True)
class AccountSuggestion:
    account_name: str
    account_number: str
    leadsheet: str
    entity: str
    score: float
    leadsheet_confidence: float
    text_similarity: float


def _require_scikit() -> None:
    if SKLEARN_IMPORT_ERROR is not None:
        raise ImportError(
            "scikit-learn is required for this module.\n"
            "Install it with: python -m pip install scikit-learn"
        ) from SKLEARN_IMPORT_ERROR


def _clean_string(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _clean_account_number(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (int, float)):
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return str(value).strip()
        if numeric.is_integer():
            return str(int(numeric))
    return str(value).strip()


def _normalized_header(value: object) -> str:
    text = _clean_string(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_account_text(value: object) -> str:
    text = _clean_string(value).lower()
    for source, target in NORMALIZATION_REPLACEMENTS.items():
        text = text.replace(source, target)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def token_overlap_ratio(left: str, right: str) -> float:
    left_tokens = {token for token in normalize_account_text(left).split() if token not in STOP_WORDS}
    right_tokens = {token for token in normalize_account_text(right).split() if token not in STOP_WORDS}
    if not left_tokens and not right_tokens:
        return 1.0
    return len(left_tokens & right_tokens) / max(len(left_tokens), len(right_tokens), 1)


def _coerce_float(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _balance_sign(value: float | None) -> str:
    if value is None:
        return "missing"
    if value > 0:
        return "positive"
    if value < 0:
        return "negative"
    return "zero"


def _account_number_match_ratio(series: pd.Series) -> float:
    values = [_clean_account_number(value) for value in series.tolist()]
    values = [value for value in values if value]
    if not values:
        return 0.0
    return sum(bool(ACCOUNT_NUMBER_PATTERN.match(value)) for value in values) / len(values)


def _detect_multi_entity_import(raw_df: pd.DataFrame) -> bool:
    sample = raw_df.iloc[:, :6].copy()
    if sample.shape[1] < 6:
        return False

    first = _account_number_match_ratio(sample.iloc[:, 1])
    third = _account_number_match_ratio(sample.iloc[:, 2])
    return third >= 0.50 and third > first


def _find_header_row(raw_df: pd.DataFrame) -> int | None:
    for index, row in raw_df.iterrows():
        normalized_values = {_normalized_header(value) for value in row.tolist()}
        if not normalized_values:
            continue
        if "leadsheet" not in normalized_values:
            continue
        if not ({"account name", "account"} & normalized_values):
            continue
        if not (HEADER_ALIASES["account_number"] & normalized_values):
            continue
        return int(index)
    return None


def _find_named_column(df: pd.DataFrame, alias_key: str, required: bool = True) -> str | None:
    aliases = HEADER_ALIASES[alias_key]
    for column in df.columns:
        if _normalized_header(column) in aliases:
            return str(column)
    if required:
        raise ValueError(f"Could not find a '{alias_key}' column in the workbook.")
    return None


def load_training_rows_from_import_workbook(
    path: str | Path,
    sheet_name: str = DEFAULT_IMPORT_SHEET,
    multi_entity: bool | None = None,
) -> list[TrainingRow]:
    raw_df = pd.read_excel(path, sheet_name=sheet_name, header=None)
    multi_entity = _detect_multi_entity_import(raw_df) if multi_entity is None else multi_entity

    if multi_entity:
        frame = raw_df.iloc[:, :6].copy()
        frame.columns = ["entity", "leadsheet", "account_number", "account_name", "py_balance", "current_balance"]
    else:
        frame = raw_df.iloc[:, :5].copy()
        frame.columns = ["leadsheet", "account_number", "account_name", "py_balance", "current_balance"]
        frame["entity"] = ""

    rows: list[TrainingRow] = []
    for _, record in frame.iterrows():
        account_name = _clean_string(record["account_name"])
        leadsheet = _clean_string(record["leadsheet"])
        account_number = _clean_account_number(record["account_number"])
        entity = _clean_string(record["entity"])

        if not account_name or not leadsheet or not account_number:
            continue

        rows.append(
            TrainingRow(
                account_name=account_name,
                leadsheet=leadsheet,
                account_number=account_number,
                entity=entity,
                current_balance=_coerce_float(record["current_balance"]),
            )
        )

    if not rows:
        raise ValueError(f"No labeled import rows were found in '{path}'.")

    return rows


def load_training_rows_from_review_workbook(
    path: str | Path,
    sheet_name: str | int | None = None,
) -> list[TrainingRow]:
    raw_df = pd.read_excel(path, sheet_name=sheet_name or 0, header=None)
    header_row = _find_header_row(raw_df)
    if header_row is None:
        raise ValueError(f"Could not locate a review-style header row in '{path}'.")

    df = pd.read_excel(path, sheet_name=sheet_name or 0, header=header_row)
    leadsheet_col = _find_named_column(df, "leadsheet")
    account_number_col = _find_named_column(df, "account_number")
    account_name_col = _find_named_column(df, "account_name")
    entity_col = _find_named_column(df, "entity", required=False)
    current_balance_col = _find_named_column(df, "current_balance", required=False)

    rows: list[TrainingRow] = []
    for _, record in df.iterrows():
        account_name = _clean_string(record[account_name_col])
        leadsheet = _clean_string(record[leadsheet_col])
        account_number = _clean_account_number(record[account_number_col])
        entity = _clean_string(record[entity_col]) if entity_col else ""
        current_balance = _coerce_float(record[current_balance_col]) if current_balance_col else None

        if not account_name or not leadsheet or not account_number:
            continue

        rows.append(
            TrainingRow(
                account_name=account_name,
                leadsheet=leadsheet,
                account_number=account_number,
                entity=entity,
                current_balance=current_balance,
            )
        )

    if not rows:
        raise ValueError(f"No labeled review rows were found in '{path}'.")

    return rows


def load_training_rows_from_workbook(
    path: str | Path,
    workbook_format: str = "auto",
    sheet_name: str | int | None = None,
    multi_entity: bool | None = None,
) -> list[TrainingRow]:
    normalized_format = workbook_format.strip().lower()
    if normalized_format == "import":
        return load_training_rows_from_import_workbook(
            path,
            sheet_name=sheet_name if sheet_name is not None else DEFAULT_IMPORT_SHEET,
            multi_entity=multi_entity,
        )
    if normalized_format == "review":
        return load_training_rows_from_review_workbook(path, sheet_name=sheet_name)
    if normalized_format != "auto":
        raise ValueError("workbook_format must be 'auto', 'import', or 'review'.")

    try:
        return load_training_rows_from_review_workbook(path, sheet_name=sheet_name)
    except ValueError:
        return load_training_rows_from_import_workbook(
            path,
            sheet_name=sheet_name if sheet_name is not None else DEFAULT_IMPORT_SHEET,
            multi_entity=multi_entity,
        )


def _build_feature_frame(rows: Sequence[TrainingRow]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "account_text": [normalize_account_text(row.account_name) for row in rows],
            "entity": [normalize_account_text(row.entity) for row in rows],
            "balance_sign": [_balance_sign(row.current_balance) for row in rows],
            "heuristic_bucket": [guess_category(row.account_name) for row in rows],
        }
    )


@dataclass
class ScikitAccountPredictor:
    leadsheet_pipeline: Any
    retrieval_vectorizer: Any
    retrieval_matrix: Any
    training_rows: list[TrainingRow]
    training_features: pd.DataFrame

    @classmethod
    def train(cls, training_rows: Sequence[TrainingRow]) -> "ScikitAccountPredictor":
        _require_scikit()

        cleaned_rows = [
            row
            for row in training_rows
            if normalize_account_text(row.account_name) and _clean_string(row.leadsheet)
        ]
        if len(cleaned_rows) < 5:
            raise ValueError("Need at least 5 labeled rows to train the model.")

        unique_leadsheets = {row.leadsheet for row in cleaned_rows}
        if len(unique_leadsheets) < 2:
            raise ValueError("Need at least 2 distinct leadsheets to train the model.")

        features = _build_feature_frame(cleaned_rows)
        labels = [row.leadsheet for row in cleaned_rows]

        leadsheet_pipeline = Pipeline(
            steps=[
                (
                    "preprocessor",
                    ColumnTransformer(
                        transformers=[
                            (
                                "word_tfidf",
                                TfidfVectorizer(ngram_range=(1, 2), min_df=1, lowercase=False),
                                "account_text",
                            ),
                            (
                                "char_tfidf",
                                TfidfVectorizer(
                                    analyzer="char_wb",
                                    ngram_range=(3, 5),
                                    min_df=1,
                                    lowercase=False,
                                ),
                                "account_text",
                            ),
                            ("entity", OneHotEncoder(handle_unknown="ignore"), ["entity"]),
                            ("balance_sign", OneHotEncoder(handle_unknown="ignore"), ["balance_sign"]),
                            (
                                "heuristic_bucket",
                                OneHotEncoder(handle_unknown="ignore"),
                                ["heuristic_bucket"],
                            ),
                        ]
                    ),
                ),
                (
                    "classifier",
                    LogisticRegression(
                        max_iter=4000,
                        class_weight="balanced",
                    ),
                ),
            ]
        )
        leadsheet_pipeline.fit(features, labels)

        retrieval_vectorizer = TfidfVectorizer(
            analyzer="char_wb",
            ngram_range=(3, 5),
            min_df=1,
            lowercase=False,
        )
        retrieval_matrix = retrieval_vectorizer.fit_transform(features["account_text"])

        return cls(
            leadsheet_pipeline=leadsheet_pipeline,
            retrieval_vectorizer=retrieval_vectorizer,
            retrieval_matrix=retrieval_matrix,
            training_rows=list(cleaned_rows),
            training_features=features.copy(),
        )

    @classmethod
    def train_from_workbooks(
        cls,
        paths: Sequence[str | Path],
        workbook_format: str = "auto",
        sheet_name: str | int | None = None,
        multi_entity: bool | None = None,
    ) -> "ScikitAccountPredictor":
        rows: list[TrainingRow] = []
        for path in paths:
            rows.extend(
                load_training_rows_from_workbook(
                    path,
                    workbook_format=workbook_format,
                    sheet_name=sheet_name,
                    multi_entity=multi_entity,
                )
            )
        if not rows:
            raise ValueError("No training rows were loaded from the supplied workbooks.")
        return cls.train(rows)

    def save(self, path: str | Path) -> None:
        with open(path, "wb") as handle:
            pickle.dump(self, handle)

    @classmethod
    def load(cls, path: str | Path) -> "ScikitAccountPredictor":
        with open(path, "rb") as handle:
            model = pickle.load(handle)
        if not isinstance(model, cls):
            raise TypeError("The loaded object is not a ScikitAccountPredictor.")
        return model

    def predict_leadsheets(
        self,
        account_name: str,
        entity: str = "",
        current_balance: float | None = None,
        top_k: int = 3,
    ) -> list[LeadsheetPrediction]:
        _require_scikit()

        query_row = TrainingRow(
            account_name=account_name,
            leadsheet="",
            account_number="",
            entity=entity,
            current_balance=current_balance,
        )
        query_frame = _build_feature_frame([query_row])
        classifier = self.leadsheet_pipeline.named_steps["classifier"]
        probabilities = self.leadsheet_pipeline.predict_proba(query_frame)[0]
        ranked = sorted(
            zip(classifier.classes_, probabilities),
            key=lambda item: float(item[1]),
            reverse=True,
        )
        return [
            LeadsheetPrediction(leadsheet=leadsheet, confidence=float(confidence))
            for leadsheet, confidence in ranked[: max(top_k, 1)]
        ]

    def suggest_accounts(
        self,
        account_name: str,
        entity: str = "",
        current_balance: float | None = None,
        top_k: int = 5,
        candidate_leadsheets: int = 3,
    ) -> list[AccountSuggestion]:
        _require_scikit()

        normalized_query = normalize_account_text(account_name)
        if not normalized_query:
            return []

        leadsheet_predictions = self.predict_leadsheets(
            account_name=account_name,
            entity=entity,
            current_balance=current_balance,
            top_k=max(candidate_leadsheets, 1),
        )
        leadsheet_confidence = {
            prediction.leadsheet: prediction.confidence for prediction in leadsheet_predictions
        }
        candidate_labels = {prediction.leadsheet for prediction in leadsheet_predictions}

        query_vector = self.retrieval_vectorizer.transform([normalized_query])
        raw_similarities = linear_kernel(query_vector, self.retrieval_matrix).ravel()
        query_entity = normalize_account_text(entity)
        query_sign = _balance_sign(current_balance)
        query_bucket = guess_category(account_name)

        if candidate_labels:
            candidate_indexes = [
                index
                for index, row in enumerate(self.training_rows)
                if row.leadsheet in candidate_labels
            ]
        else:
            candidate_indexes = list(range(len(self.training_rows)))

        if len(candidate_indexes) < top_k:
            candidate_indexes = list(range(len(self.training_rows)))

        ranked: list[AccountSuggestion] = []
        seen_keys: set[tuple[str, str, str, str]] = set()
        for index in candidate_indexes:
            training_row = self.training_rows[index]
            training_feature = self.training_features.iloc[index]

            lexical_similarity = float(raw_similarities[index])
            sequence_similarity = SequenceMatcher(
                None,
                normalized_query,
                training_feature["account_text"],
            ).ratio()
            overlap_similarity = token_overlap_ratio(account_name, training_row.account_name)
            text_similarity = max(
                lexical_similarity,
                (lexical_similarity + sequence_similarity + overlap_similarity) / 3.0,
            )

            bonus = 0.0
            if query_entity:
                if query_entity == training_feature["entity"]:
                    bonus += 0.05
                elif training_feature["entity"]:
                    bonus -= 0.02
            if query_sign != "missing" and query_sign == training_feature["balance_sign"]:
                bonus += 0.03
            if query_bucket == training_feature["heuristic_bucket"]:
                bonus += 0.02

            combined_score = 0.65 * text_similarity
            combined_score += 0.30 * leadsheet_confidence.get(training_row.leadsheet, 0.0)
            combined_score += bonus
            combined_score = max(0.0, min(1.0, combined_score))

            key = (
                training_row.leadsheet,
                training_row.account_number,
                training_row.account_name,
                training_row.entity,
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)

            ranked.append(
                AccountSuggestion(
                    account_name=training_row.account_name,
                    account_number=training_row.account_number,
                    leadsheet=training_row.leadsheet,
                    entity=training_row.entity,
                    score=float(combined_score),
                    leadsheet_confidence=float(leadsheet_confidence.get(training_row.leadsheet, 0.0)),
                    text_similarity=float(text_similarity),
                )
            )

        ranked.sort(key=lambda item: item.score, reverse=True)
        return ranked[: max(top_k, 1)]


def _format_predictions(predictions: Iterable[LeadsheetPrediction]) -> str:
    lines = []
    for prediction in predictions:
        lines.append(f"{prediction.leadsheet}: {prediction.confidence:.3f}")
    return "\n".join(lines)


def _format_suggestions(suggestions: Iterable[AccountSuggestion]) -> str:
    lines = []
    for suggestion in suggestions:
        entity_prefix = f"{suggestion.entity} | " if suggestion.entity else ""
        lines.append(
            f"{entity_prefix}{suggestion.leadsheet} | {suggestion.account_number} | "
            f"{suggestion.account_name} | score={suggestion.score:.3f} "
            f"text={suggestion.text_similarity:.3f} "
            f"lead={suggestion.leadsheet_confidence:.3f}"
        )
    return "\n".join(lines)


def _build_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Train and query the optional scikit account predictor.")
    subparsers = parser.add_subparsers(dest="command")

    train_parser = subparsers.add_parser("train", help="Train a predictor from labeled workbooks.")
    train_parser.add_argument("--workbook", action="append", default=[], help="Workbook to auto-detect.")
    train_parser.add_argument(
        "--import-workbook",
        action="append",
        default=[],
        help="Workbook already in import layout.",
    )
    train_parser.add_argument(
        "--review-workbook",
        action="append",
        default=[],
        help="Workbook with review-style headers like Leadsheet and Account Number.",
    )
    train_parser.add_argument("--sheet-name", default=None, help="Optional sheet name override.")
    train_parser.add_argument(
        "--model-out",
        required=True,
        help="Where to save the trained pickle model.",
    )

    predict_parser = subparsers.add_parser("predict", help="Predict leadsheets and suggest accounts.")
    predict_parser.add_argument("--model", required=True, help="Path to a saved predictor pickle.")
    predict_parser.add_argument("--account", required=True, help="Account name to score.")
    predict_parser.add_argument("--entity", default="", help="Optional entity/fund name.")
    predict_parser.add_argument("--balance", type=float, default=None, help="Optional CY balance.")
    predict_parser.add_argument("--top-k", type=int, default=5, help="Number of account suggestions to show.")

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_cli_parser()
    args = parser.parse_args(argv)

    if args.command == "train":
        _require_scikit()

        rows: list[TrainingRow] = []
        for workbook in args.workbook:
            rows.extend(load_training_rows_from_workbook(workbook, sheet_name=args.sheet_name))
        for workbook in args.import_workbook:
            rows.extend(
                load_training_rows_from_import_workbook(
                    workbook,
                    sheet_name=args.sheet_name or DEFAULT_IMPORT_SHEET,
                )
            )
        for workbook in args.review_workbook:
            rows.extend(load_training_rows_from_review_workbook(workbook, sheet_name=args.sheet_name))

        if not rows:
            parser.error("Provide at least one workbook with --workbook, --import-workbook, or --review-workbook.")

        predictor = ScikitAccountPredictor.train(rows)
        predictor.save(args.model_out)

        unique_leadsheets = len({row.leadsheet for row in rows})
        print(
            f"Trained model on {len(rows)} rows across {unique_leadsheets} leadsheets. "
            f"Saved to {args.model_out}."
        )
        return 0

    if args.command == "predict":
        _require_scikit()

        predictor = ScikitAccountPredictor.load(args.model)
        predictions = predictor.predict_leadsheets(
            account_name=args.account,
            entity=args.entity,
            current_balance=args.balance,
            top_k=min(max(args.top_k, 1), 5),
        )
        suggestions = predictor.suggest_accounts(
            account_name=args.account,
            entity=args.entity,
            current_balance=args.balance,
            top_k=max(args.top_k, 1),
        )

        print("Top leadsheets")
        print(_format_predictions(predictions) or "(none)")
        print()
        print("Suggested accounts")
        print(_format_suggestions(suggestions) or "(none)")
        return 0

    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
