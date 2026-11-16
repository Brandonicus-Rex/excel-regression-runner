"""End-to-end: run the sample workbook through the CLI against the mock API."""

import socket
import threading
import time

import pytest
import requests
import uvicorn
from openpyxl import load_workbook

from mock_server.app import API_PASSWORD, API_USER, app
from runner import cli
from samples.make_sample import build_sample


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="module")
def base_url():
    """Run the mock API in a background thread for this module's tests."""
    port = _free_port()
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    url = f"http://127.0.0.1:{port}"
    for _ in range(50):
        try:
            if requests.get(f"{url}/health", timeout=0.5).ok:
                break
        except requests.ConnectionError:
            time.sleep(0.1)
    else:
        raise RuntimeError("Mock server did not start")
    yield url
    server.should_exit = True
    thread.join(timeout=5)


@pytest.fixture
def env(base_url, monkeypatch):
    """Point the runner at the mock API and start from an empty database."""
    monkeypatch.setenv("API_USERNAME", API_USER)
    monkeypatch.setenv("API_PASSWORD", API_PASSWORD)
    requests.post(f"{base_url}/api/v1/_reset", auth=(API_USER, API_PASSWORD), timeout=5)
    return base_url


def test_sample_workbook_all_pass(env, tmp_path, capsys):
    src = build_sample(tmp_path / "tests.xlsx")
    out = tmp_path / "results.xlsx"
    code = cli.main([str(src), "--base-url", env, "--output", str(out)])
    assert code == 0, capsys.readouterr().out

    wb = load_workbook(out)
    assert wb.sheetnames[0] == "Summary"
    results = [c.value for ws in ("Lifecycle", "Usage & Notifications", "Negative") for c in wb[ws]["J"][1:]]
    assert results and all(r == "PASS" for r in results)
    # The input workbook is never modified.
    assert load_workbook(src)["Lifecycle"]["J2"].value is None


def test_cleanup_rows_make_rounds_repeatable(env, tmp_path, capsys):
    """Run the full sample twice against the same server, with no reset in between.

    The cleanup rows close every subscriber the sheets create, so the second
    round's "create" rows must not hit 409 Conflict.
    """
    src = build_sample(tmp_path / "tests.xlsx")
    for round_no in (1, 2):
        out = tmp_path / f"round{round_no}.xlsx"
        code = cli.main([str(src), "--base-url", env, "--output", str(out)])
        assert code == 0, f"round {round_no}: {capsys.readouterr().out}"


def test_cleanup_rows_pass_when_nothing_to_clean(env, tmp_path):
    """Cleanup on a fresh server gets 404 (never created) and still passes."""
    src = build_sample(tmp_path / "tests.xlsx")
    out = tmp_path / "results.xlsx"
    assert cli.main([str(src), "--sheet", "Negative", "--rows", "2", "--base-url", env, "--output", str(out)]) == 0
    ws = load_workbook(out)["Negative"]
    assert (ws["A2"].value, ws["H2"].value, ws["J2"].value) == ("NG-00", 404, "PASS")


def test_setup_rows_recover_from_an_interrupted_round(env, tmp_path, capsys):
    """A round that stopped halfway leaves open accounts; the next full round still passes.

    Without the setup rows at the top, the "create" rows would fail with 409.
    """
    src = build_sample(tmp_path / "tests.xlsx")
    # Interrupted round: only the create rows ran, so nothing was cleaned up.
    for sheet, rows in (("Lifecycle", "3"), ("Usage & Notifications", "4,8"), ("Negative", "6")):
        assert cli.main([str(src), "--sheet", sheet, "--rows", rows, "--base-url", env,
                         "--output", str(tmp_path / "partial.xlsx")]) == 0
    out = tmp_path / "results.xlsx"
    code = cli.main([str(src), "--base-url", env, "--output", str(out)])
    assert code == 0, capsys.readouterr().out


def test_status_list_failure_message(env, tmp_path):
    src = build_sample(tmp_path / "tests.xlsx")
    wb = load_workbook(src)
    wb["Negative"]["F3"] = "200,201"  # NG-01 really returns 404
    wb.save(src)
    out = tmp_path / "results.xlsx"
    assert cli.main([str(src), "--sheet", "Negative", "--rows", "3", "--base-url", env, "--output", str(out)]) == 1
    assert "expected status 200 or 201, got 404" in load_workbook(out)["Negative"]["K3"].value


def test_failures_are_reported_and_exit_code_is_1(env, tmp_path, capsys):
    src = build_sample(tmp_path / "tests.xlsx")
    wb = load_workbook(src)
    wb["Lifecycle"]["G3"] = "status=suspended"  # LC-01: wrong on purpose
    wb.save(src)

    out = tmp_path / "results.xlsx"
    code = cli.main([str(src), "--sheet", "Lifecycle", "--rows", "3", "--base-url", env, "--output", str(out)])
    assert code == 1
    ws = load_workbook(out)["Lifecycle"]
    assert ws["J3"].value == "FAIL"
    assert "missing words: suspended" in ws["K3"].value
    assert "FAIL Lifecycle!3" in capsys.readouterr().out


def test_setup_error_exit_code_2(env, tmp_path):
    src = build_sample(tmp_path / "tests.xlsx")
    assert cli.main([str(src), "--sheet", "Missing", "--base-url", env]) == 2
    assert cli.main([str(src), "--rows", "abc", "--base-url", env]) == 2
