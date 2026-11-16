"""Runs test cases: send the request, capture the response, check expectations."""

from __future__ import annotations

import logging

import requests

from runner.actions import ACTIONS
from runner.client import ApiClient
from runner.validators import evaluate, flatten
from runner.workbook import Outcome, TestCase

log = logging.getLogger(__name__)


def run_case(client: ApiClient, case: TestCase) -> Outcome:
    """Run one test case and return its outcome. Never raises.

    A failure in one row (unknown action, network error, bad response) is
    recorded on that row so the rest of the run continues.
    """
    action = ACTIONS.get(case.action)
    if action is None:
        known = ", ".join(sorted(ACTIONS))
        return Outcome(case, None, "", False, [f"unknown action '{case.action}' (known: {known})"])

    try:
        response = action(client, case.phone, case.params)
    except requests.RequestException as exc:
        return Outcome(case, None, "", False, [f"request failed: {exc}"])

    # Flatten JSON into key=value lines for both display and rule matching.
    # Non-JSON bodies (e.g. an HTML error page) are shown as plain text.
    try:
        lines = flatten(response.json())
    except ValueError:
        lines = [response.text.strip()] if response.text.strip() else []
    actual_text = "\n".join(lines)

    details: list[str] = []
    if case.expected_status and response.status_code not in case.expected_status:
        wanted = " or ".join(str(code) for code in case.expected_status)
        details.append(f"expected status {wanted}, got {response.status_code}")
    details.extend(evaluate(case.expected_response, lines).failures)

    passed = not details
    log.info("%s %s!%s %s", "PASS" if passed else "FAIL", case.sheet, case.row, case.case_id)
    return Outcome(case, response.status_code, actual_text, passed, details)
