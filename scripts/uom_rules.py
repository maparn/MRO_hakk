"""Shared diagnostics for the AMOS UOM export. These checks never apply a conversion."""

from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from math import isfinite


def normalize_key(value):
    return "" if value is None else str(value).strip().upper()


def numeric_factor(value):
    try:
        amount = Decimal(str(value).strip().replace(",", "."))
    except InvalidOperation:
        return None
    if not amount.is_finite():
        return None
    excel_value = float(amount)
    if not isfinite(excel_value) or (amount and excel_value == 0):
        return None
    return amount


def _factor_key(value):
    # Keep malformed tokens distinct. Decimal compares 1, 1.0 and "1,00" equally.
    number = numeric_factor(value)
    return number if number is not None else normalize_key(value)


def rule_signature(rule):
    return (
        normalize_key(rule["partno"]),
        normalize_key(rule["measure_unit_ext"]),
        normalize_key(rule["measure_unit"]),
        _factor_key(rule["rate"]),
        _factor_key(rule["exchange_base"]),
    )


def annotate_rules(rules):
    """Annotate every source row; a second external unit is a legitimate alternative."""
    signatures = [rule_signature(rule) for rule in rules]
    repetitions = Counter(sig for sig in signatures if all(sig[:3]))
    definitions = defaultdict(set)
    for part, external, internal, rate, base in signatures:
        if part and external:
            definitions[part, external].add((internal, rate, base))

    for rule, sig in zip(rules, signatures):
        part, external, internal, rate, base = sig
        issues = []
        if not all((part, external, internal)):
            issues.append("MISSING_UOM_KEY")
        if not all(isinstance(x, Decimal) and x > 0 for x in (rate, base)):
            issues.append("INVALID_UOM_FACTOR")
        if repetitions[sig] > 1:
            issues.append("REPEATED_UOM_RULE")
        if len(definitions.get((part, external), ())) > 1:
            issues.append("CONFLICTING_UOM_RULES")
        rule["exact_rule_repetitions"] = repetitions[sig]
        rule["rule_issues"] = tuple(issues)
