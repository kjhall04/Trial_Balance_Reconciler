from __future__ import annotations

import math
import re
from difflib import SequenceMatcher


ACCOUNT_NUMBER_RE = re.compile(r"^\d{3,6}(?:-\d{1,4})?$")
ACCOUNT_NUMBER_PREFIX_RE = re.compile(r"^\d{3,6}(?:-\d{1,4})?$")
NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and math.isnan(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()


def clean_account_number(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    return re.sub(r"\.0+$", "", text)


def standard_account_number(value: object) -> str:
    text = clean_account_number(value)
    if not text:
        return ""
    if "-" in text:
        left, right = text.split("-", 1)
        if left.isdigit() and right.isdigit():
            return f"{left}-{right.zfill(2)}"
        return text
    if text.isdigit():
        if len(text) == 4:
            return f"{text}-00"
        if len(text) == 5:
            return f"{text[:4]}-00"
        if len(text) == 6:
            return f"{text[:4]}-{text[4:].zfill(2)}"
    return text


def account_family(value: object) -> str:
    text = standard_account_number(value)
    if not text:
        return ""
    if "-" in text:
        return text.split("-", 1)[0]
    if text.isdigit():
        return text[:4]
    return text


def account_suffix(value: object) -> str:
    text = standard_account_number(value)
    if not text:
        return ""
    if "-" in text:
        return text.split("-", 1)[1].zfill(2)
    return "00"


def make_account_number(family: str, suffix: int | str) -> str:
    family_text = clean_text(family)
    if not family_text:
        return ""
    return f"{family_text}-{str(suffix).zfill(2)}"


def is_zero_balance(value: object, places: int = 2) -> bool:
    try:
        amount = 0.0 if value is None else float(value)
    except (TypeError, ValueError):
        amount = 0.0
    return round(amount, places) == 0.0


def normalize_match_text(value: object) -> str:
    text = clean_text(value).replace("·", ":").replace("ｷ", ":")
    replacements = {
        "&": " and ",
        "a/r": " accounts receivable ",
        "a/p": " accounts payable ",
        "n/p": " notes payable ",
        "#": " ",
    }
    lowered = text.lower()
    for source, target in replacements.items():
        lowered = lowered.replace(source, target)
    lowered = NON_ALNUM_RE.sub(" ", lowered)
    return re.sub(r"\s+", " ", lowered).strip()


def split_account_segments(value: object) -> list[str]:
    text = clean_text(value).replace("·", ":").replace("ｷ", ":")
    if not text:
        return []
    pieces = [piece.strip(" :-") for piece in text.split(":")]
    return [piece for piece in pieces if piece]


def segment_account_number(segment: str) -> str:
    text = clean_text(segment)
    if not text:
        return ""
    first_token = text.split(" ", 1)[0].strip(" :-")
    return clean_account_number(first_token) if ACCOUNT_NUMBER_PREFIX_RE.fullmatch(first_token) else ""


def extract_path_numbers(value: object, fallback: object = "") -> list[str]:
    numbers: list[str] = []
    for segment in split_account_segments(value):
        number = segment_account_number(segment)
        if number:
            numbers.append(number)
    fallback_number = clean_account_number(fallback)
    if fallback_number and not numbers:
        numbers.append(fallback_number)
    return numbers


def extract_leaf_account_number(value: object, fallback: object = "") -> str:
    numbers = extract_path_numbers(value, fallback=fallback)
    return numbers[-1] if numbers else ""


def extract_parent_account_number(value: object, fallback: object = "") -> str:
    numbers = extract_path_numbers(value, fallback=fallback)
    return numbers[-2] if len(numbers) >= 2 else ""


def clean_account_name(value: object) -> str:
    cleaned_segments: list[str] = []
    for segment in split_account_segments(value):
        without_number = re.sub(r"^\d{3,6}(?:-\d{1,4})?\s*", "", segment)
        without_number = without_number.strip(" :-")
        if without_number:
            cleaned_segments.append(without_number)
    return ":".join(cleaned_segments)


def path_depth(value: object) -> int:
    return len(split_account_segments(value))


def similarity(left: object, right: object) -> float:
    return SequenceMatcher(None, normalize_match_text(left), normalize_match_text(right)).ratio()


def token_overlap(left: object, right: object) -> float:
    left_tokens = set(normalize_match_text(left).split())
    right_tokens = set(normalize_match_text(right).split())
    if not left_tokens and not right_tokens:
        return 1.0
    return len(left_tokens & right_tokens) / max(len(left_tokens), len(right_tokens), 1)

