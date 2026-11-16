"""Unit tests for the comparison rules."""

from runner.validators import evaluate, flatten, parse_rules, word_diff

LINES = [
    "status=active",
    "features=voicemail",
    "features=mms",
    "remaining.data_mb=150",
    "notifications=Data usage at 80% of allowance",
]


# ---- flatten() -----------------------------------------------------------

def test_flatten_nested_lists_and_none():
    data = {"a": {"b": 1}, "tags": ["x", "y"], "items": [{"n": 1}], "empty": None, "flag": True}
    assert flatten(data) == ["a.b=1", "tags=x", "tags=y", "items.0.n=1", "empty=", "flag=true"]


# ---- parse_rules() -------------------------------------------------------

def test_parse_rules_kinds_and_blank_lines():
    rules = parse_rules("status=active\n\n!features=sms\nONCE: notifications=*")
    assert [(r.kind, r.pattern) for r in rules] == [
        ("match", "status=active"), ("absent", "features=sms"), ("once", "notifications=*")]


# ---- evaluate() ----------------------------------------------------------

def test_empty_expected_always_passes():
    assert evaluate("", LINES).passed


def test_match_is_case_insensitive_with_wildcards():
    assert evaluate("STATUS=Active\nnotifications=*80%*", LINES).passed


def test_absent_rule():
    assert evaluate("!features=sms", LINES).passed
    result = evaluate("!features=mms", LINES)
    assert not result.passed and "unwanted" in result.failures[0]


def test_once_rule_catches_duplicates():
    assert evaluate("once:features=mms", LINES).passed
    result = evaluate("once:features=*", LINES)  # two feature lines
    assert not result.passed and "found 2" in result.failures[0]


def test_every_failure_is_reported():
    result = evaluate("status=suspended\n!features=mms", LINES)
    assert len([f for f in result.failures if "matches" in f]) == 2


def test_failure_includes_word_diff_for_literal_rules():
    result = evaluate("status=suspended", LINES)
    assert "missing words: suspended" in result.failures


# ---- word_diff() ---------------------------------------------------------

def test_word_diff_counts_repeats_and_ignores_case():
    missing, unexpected = word_diff("SMS sms data", "sms voice")
    assert sorted(missing) == ["data", "sms"]
    assert unexpected == ["voice"]
