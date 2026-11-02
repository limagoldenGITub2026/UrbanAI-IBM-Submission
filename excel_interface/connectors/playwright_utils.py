"""
playwright_utils.py — Shared Playwright helper for browser-based connectors.

Provides:
  is_playwright_available() -> bool
  get_page()               -> context manager yielding a playwright Page
"""

from __future__ import annotations

import contextlib
import logging
from typing import Generator

logger = logging.getLogger(__name__)

# Chrome-like user agent to avoid bot detection
_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


def is_playwright_available() -> bool:
    """Return True if the playwright package and Chromium browser are importable."""
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
        return True
    except ImportError:
        return False


INSTALL_NOTE = (
    "Playwright לא מותקן. להתקנה:\n"
    "  pip install playwright\n"
    "  python -m playwright install chromium"
)


@contextlib.contextmanager
def get_page(timeout_ms: int = 30_000) -> Generator:
    """Context manager that yields a Playwright Page (headless Chromium).

    Usage
    -----
        with get_page() as page:
            page.goto("https://...")
            text = page.inner_text("h1")

    Parameters
    ----------
    timeout_ms:
        Default navigation/wait timeout in milliseconds (default 30s).
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            user_agent=_USER_AGENT,
            locale="he-IL",
            viewport={"width": 1280, "height": 800},
            ignore_https_errors=True,
        )
        context.set_default_timeout(timeout_ms)
        page = context.new_page()
        try:
            yield page
        finally:
            context.close()
            browser.close()
