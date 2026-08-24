from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence


@dataclass(frozen=True)
class TaxRateRule:
    from_year: int
    multiplier: float


# Starting defaults for the global tax multiplier rules, editable from the
# "Налоговые множители" window and shared by every task that needs an after-tax value.
DEFAULT_TAX_RULES: tuple[TaxRateRule, ...] = (
    TaxRateRule(from_year=0, multiplier=0.8),
    TaxRateRule(from_year=2025, multiplier=0.75),
)


def tax_multiplier_for_year(year: int, tax_rules: Sequence[TaxRateRule]) -> float:
    selected: float | None = None
    for rule in sorted(tax_rules, key=lambda item: item.from_year):
        if year >= rule.from_year:
            selected = rule.multiplier

    if selected is None:
        raise ValueError("tax_rules must contain a rule with from_year <= the requested year")

    return float(selected)
