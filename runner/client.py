"""HTTP client for the API under test: timeouts, retries, config from env."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class ClientConfig:
    """Connection settings. Credentials come from the environment, never code."""

    base_url: str
    username: str | None = None
    password: str | None = None
    timeout: float = 10.0
    retries: int = 2

    @classmethod
    def from_env(cls, base_url: str | None = None) -> "ClientConfig":
        """Build a config from environment variables (see ``.env.example``).

        Args:
            base_url: Optional override for ``API_BASE_URL``.

        Raises:
            RuntimeError: If no base URL is available.
        """
        url = base_url or os.getenv("API_BASE_URL")
        if not url:
            raise RuntimeError("API_BASE_URL is not set. Copy .env.example to .env or pass --base-url.")
        return cls(
            base_url=url.rstrip("/"),
            username=os.getenv("API_USERNAME"),
            password=os.getenv("API_PASSWORD"),
            timeout=float(os.getenv("API_TIMEOUT", "10")),
            retries=int(os.getenv("API_RETRIES", "2")),
        )


class ApiClient:
    """Reusable HTTP session. Use as a context manager to close connections."""

    def __init__(self, config: ClientConfig) -> None:
        self.config = config
        self.session = requests.Session()
        if config.username and config.password:
            self.session.auth = (config.username, config.password)
        self.session.headers["Accept"] = "application/json"

        # Retry transient gateway errors on idempotent methods only. POST is
        # excluded because repeating it could apply a change twice.
        retry = Retry(total=config.retries, backoff_factor=0.5, status_forcelist=(502, 503, 504),
                      allowed_methods=("GET", "PUT", "DELETE", "HEAD"))
        adapter = HTTPAdapter(max_retries=retry)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

    def request(self, method: str, path: str, json: dict | None = None) -> requests.Response:
        """Send a request relative to the base URL and return the response.

        HTTP error statuses are returned, not raised, because test cases often
        expect them (e.g. a 404 or 409).
        """
        url = f"{self.config.base_url}/{path.lstrip('/')}"
        response = self.session.request(method, url, json=json, timeout=self.config.timeout)
        log.info("%s %s -> %s", method, path, response.status_code)
        return response

    def close(self) -> None:
        self.session.close()

    def __enter__(self) -> "ApiClient":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
