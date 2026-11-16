"""Maps the spreadsheet's Action column to API calls.

Each test row names an action (e.g. ``suspend``) and a phone number, plus
optional ``key=value`` parameters. Adding a new action is one function and
one entry in ``ACTIONS``.
"""

from __future__ import annotations

from typing import Callable

import requests

from runner.client import ApiClient

BASE = "/api/v1/subscribers"

# An action receives the client, the subscriber's phone number and the row's
# parameters, and returns the raw HTTP response.
Action = Callable[[ApiClient, str, dict[str, str]], requests.Response]


def _split_list(value: str | None) -> list[str]:
    """'voicemail, mms' -> ['voicemail', 'mms']"""
    return [v.strip() for v in (value or "").split(",") if v.strip()]


def create(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    body = {"phone_number": phone, "plan": params.get("plan", "basic"),
            "features": _split_list(params.get("features"))}
    return client.request("POST", BASE, json=body)


def get(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("GET", f"{BASE}/{phone}")


def suspend(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("POST", f"{BASE}/{phone}/suspend")


def restore(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("POST", f"{BASE}/{phone}/restore")


def add_feature(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("POST", f"{BASE}/{phone}/features", json={"add": _split_list(params.get("feature"))})


def remove_feature(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("POST", f"{BASE}/{phone}/features", json={"remove": _split_list(params.get("feature"))})


def usage(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    # Amount is sent as given; the API validates it, so a bad value becomes
    # a 422 the test can assert on rather than a crash here.
    body = {"type": params.get("type", ""), "amount": params.get("amount", "0")}
    return client.request("POST", f"{BASE}/{phone}/usage", json=body)


def close(client: ApiClient, phone: str, params: dict[str, str]) -> requests.Response:
    return client.request("POST", f"{BASE}/{phone}/close")


ACTIONS: dict[str, Action] = {
    "create": create,
    "get": get,
    "suspend": suspend,
    "restore": restore,
    "add_feature": add_feature,
    "remove_feature": remove_feature,
    "usage": usage,
    "close": close,
}
