from __future__ import annotations

from .normalize import normalize_match_text, similarity
from .matching import build_trial_balance
from .models import ClientConfig, TrialBalanceBuildResult, WorkbookSpec
from .current_year import read_current_workbooks
from .prior_year import read_prior_workbooks


def build_from_workbooks(
    *,
    current_specs: list[WorkbookSpec],
    prior_specs: list[WorkbookSpec],
    entity_overrides: dict[str, str] | None = None,
    client_config: ClientConfig | None = None,
    keep_zero_rows: bool = True,
) -> TrialBalanceBuildResult:
    prior_df = read_prior_workbooks(
        prior_specs,
        entity_overrides=entity_overrides,
        client_config=client_config,
    )
    known_entities = [entity for entity in prior_df["entity"].astype(str).unique().tolist() if str(entity).strip()]
    current_df = read_current_workbooks(
        current_specs,
        entity_overrides=entity_overrides,
        known_entities=known_entities,
        client_config=client_config,
    )
    current_entities = [entity for entity in current_df["entity"].astype(str).unique().tolist() if str(entity).strip()]
    prior_entity_map: dict[str, str] = {}
    for prior_entity in known_entities:
        best_entity = prior_entity
        best_score = -1.0
        for current_entity in current_entities:
            score = similarity(prior_entity, current_entity)
            if normalize_match_text(prior_entity) in normalize_match_text(current_entity) or normalize_match_text(current_entity) in normalize_match_text(prior_entity):
                score += 0.1
            if score > best_score:
                best_entity = current_entity
                best_score = score
        if best_score >= 0.72:
            prior_entity_map[prior_entity] = best_entity
    if prior_entity_map:
        prior_df = prior_df.copy()
        prior_df["entity"] = prior_df["entity"].map(lambda value: prior_entity_map.get(value, value))
    return build_trial_balance(current_df, prior_df, keep_zero_rows=keep_zero_rows)
