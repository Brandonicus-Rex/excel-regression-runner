"""Rules for comparing an expected-result cell with an actual API response.

The actual response is flattened into ``key=value`` lines (see ``flatten``),
for example::

    status=active
    plan=basic
    features=voicemail
    features=mms
    remaining.data_mb=150
    notifications=Data usage at 80% of allowance

An expected-result cell holds one rule per line. Every rule must pass.

    status=active        at least one line matches (wildcards * and ? allowed)
    !features=sms        no line may match ("must not contain")
    once:notifications=* exactly one line must match (catches duplicates)

Matching is case-insensitive. An empty expected cell means "don't check".
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from fnmatch import fnmatchcase
from typing import Any

# Rule prefixes. Kept short so they're easy to type in a spreadsheet cell.
ABSENT_PREFIX = "!"
ONCE_PREFIX = "once:"


@dataclass(frozen=True)
class Rule:
    """One line of an expected-result cell.

    Attributes:
        kind: ``"match"``, ``"absent"`` or ``"once"``.
        pattern: Wildcard pattern to compare against each actual line.
        raw: The original text, used in failure messages.
    """

    kind: str
    pattern: str
    raw: str


@dataclass
class CheckResult:
    """Outcome of checking one expected cell against the actual lines."""

    passed: bool
    failures: list[str] = field(default_factory=list)


# ---- Flattening responses ------------------------------------------------

def flatten(data: Any, prefix: str = "") -> list[str]:
    """Turn parsed JSON into ``key=value`` lines that rules can match.

    - Nested objects use dotted keys: ``{"remaining": {"sms": 5}}`` -> ``remaining.sms=5``
    - Lists of plain values repeat the key: ``{"features": ["a", "b"]}`` ->
      ``features=a`` and ``features=b``, so rules can check list membership.
    - Lists of objects use an index: ``items.0.name=...``
    - None becomes an empty value.
    """
    lines: list[str] = []
    if isinstance(data, dict):
        for key, value in data.items():
            lines.extend(flatten(value, f"{prefix}.{key}" if prefix else str(key)))
    elif isinstance(data, list):
        for i, item in enumerate(data):
            if isinstance(item, (dict, list)):
                lines.extend(flatten(item, f"{prefix}.{i}"))
            else:
                lines.extend(flatten(item, prefix))
    else:
        value = "" if data is None else str(data)
        # Booleans print as true/false so they read the same as JSON.
        if isinstance(data, bool):
            value = str(data).lower()
        lines.append(f"{prefix}={value}")
    return lines


# ---- Parsing and evaluating rules ----------------------------------------

def parse_rules(expected: str | None) -> list[Rule]:
    """Split an expected-result cell into rules, one per non-blank line."""
    rules: list[Rule] = []
    for raw in (expected or "").splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith(ABSENT_PREFIX):
            rules.append(Rule("absent", line[len(ABSENT_PREFIX):].strip(), line))
        elif line.lower().startswith(ONCE_PREFIX):
            rules.append(Rule("once", line[len(ONCE_PREFIX):].strip(), line))
        else:
            rules.append(Rule("match", line, line))
    return rules


def _matches(pattern: str, line: str) -> bool:
    """Case-insensitive wildcard match of a whole line."""
    return fnmatchcase(line.lower(), pattern.lower())


def evaluate(expected: str | None, actual_lines: list[str]) -> CheckResult:
    """Check every rule in ``expected`` against the actual response lines.

    All rules are evaluated (not just up to the first failure) so a failed
    row reports every problem at once.
    """
    failures: list[str] = []
    for rule in parse_rules(expected):
        hits = sum(_matches(rule.pattern, line) for line in actual_lines)
        if rule.kind == "match" and hits == 0:
            failures.append(f"no line matches '{rule.pattern}'")
        elif rule.kind == "absent" and hits > 0:
            failures.append(f"unwanted line matches '{rule.pattern}'")
        elif rule.kind == "once" and hits != 1:
            failures.append(f"expected exactly 1 line matching '{rule.pattern}', found {hits}")

    # When plain (non-wildcard) expectations fail, a word diff makes it
    # obvious what changed, e.g. "missing: suspended | unexpected: active".
    if failures:
        literal = [r.pattern for r in parse_rules(expected) if r.kind == "match" and not _has_wildcard(r.pattern)]
        if literal:
            missing, unexpected = word_diff(" ".join(literal), " ".join(actual_lines))
            if missing:
                failures.append("missing words: " + ", ".join(missing))
            if unexpected and missing:
                # Only show unexpected words alongside missing ones; on their own
                # they're just the rest of the response and add noise.
                failures.append("unexpected words: " + ", ".join(unexpected[:10]))
    return CheckResult(passed=not failures, failures=failures)


def _has_wildcard(pattern: str) -> bool:
    return any(ch in pattern for ch in "*?[")


# ---- Word diff -----------------------------------------------------------

def _words(text: str) -> list[str]:
    """Split on whitespace, '=' and line breaks, lowercased."""
    return [w for w in text.replace("=", " ").lower().split() if w]


def word_diff(expected: str, actual: str) -> tuple[list[str], list[str]]:
    """Compare two texts word by word, ignoring order and case.

    Returns:
        (missing, unexpected): words in ``expected`` but not ``actual``, and
        words in ``actual`` but not ``expected``. Repeated words are counted,
        so "sms sms" vs "sms" reports one missing "sms".
    """
    exp, act = Counter(_words(expected)), Counter(_words(actual))
    missing = list((exp - act).elements())
    unexpected = list((act - exp).elements())
    return missing, unexpected
