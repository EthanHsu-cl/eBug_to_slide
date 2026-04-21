import os
import sys
from pathlib import Path

import browser_cookie3
import requests

EBUG_BASE = "https://ecl.cyberlink.com/Ebug/eBugHandle"
CONTENT_URL = EBUG_BASE + "/HandleMainEbugContent.asp"
DOMAIN = ".cyberlink.com"

ENV_PATH = Path(__file__).parent / ".env"

# All browsers supported by browser-cookie3
SUPPORTED_BROWSERS = ["brave", "chrome", "edge", "safari", "firefox", "chromium", "opera", "vivaldi"]
AUTO_ORDER = ["brave", "chrome", "edge", "safari"]


# ---------------------------------------------------------------------------
# .env helpers (browser preference persistence)
# ---------------------------------------------------------------------------

def _load_env() -> None:
    if not ENV_PATH.exists():
        return
    with open(ENV_PATH, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            if key.strip() not in os.environ:
                os.environ[key.strip()] = value.strip()


def load_browser_preference() -> str:
    _load_env()
    return os.environ.get("EBUG_BROWSER", "auto")


def save_browser_preference(browser: str) -> None:
    lines: list[str] = []
    found = False
    if ENV_PATH.exists():
        with open(ENV_PATH, encoding="utf-8") as fh:
            for line in fh:
                if line.strip().startswith("EBUG_BROWSER="):
                    lines.append(f"EBUG_BROWSER={browser}\n")
                    found = True
                else:
                    lines.append(line)
    if not found:
        lines.append(f"EBUG_BROWSER={browser}\n")
    with open(ENV_PATH, "w", encoding="utf-8") as fh:
        fh.writelines(lines)
    os.environ["EBUG_BROWSER"] = browser


# ---------------------------------------------------------------------------
# Cookie extraction
# ---------------------------------------------------------------------------

def _try_browser(name: str) -> "requests.cookies.RequestsCookieJar | None":
    loader = getattr(browser_cookie3, name, None)
    if loader is None:
        return None
    try:
        cj = loader(domain_name=DOMAIN)
        if any(True for _ in cj):
            return cj
    except Exception:
        pass
    return None


def get_cookies(browser: str = "auto") -> "requests.cookies.RequestsCookieJar":
    """Return a CookieJar for ecl.cyberlink.com from the user's browser."""
    candidates = AUTO_ORDER if browser == "auto" else [browser]
    for name in candidates:
        cj = _try_browser(name)
        if cj is not None:
            return cj

    print(
        "ERROR: Could not extract cookies for ecl.cyberlink.com from any browser.\n"
        f"Tried: {', '.join(candidates)}\n"
        "Make sure you are logged in at https://ecl.cyberlink.com in one of those browsers.",
        file=sys.stderr,
    )
    sys.exit(1)


# ---------------------------------------------------------------------------
# HTTP fetch
# ---------------------------------------------------------------------------

def fetch_bug(bug_code: str, browser: str = "auto") -> tuple[requests.Session, str]:
    """
    Fetch the eBug content page and return (session, html_text).
    The main HandleMainEbug.asp page is a frameset; the real content is in
    HandleMainEbugContent.asp served by the HandleMain frame.
    """
    # Resolve "auto" through the saved preference first
    if browser == "auto":
        browser = load_browser_preference()

    cj = get_cookies(browser)
    session = requests.Session()
    session.cookies.update(cj)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Referer": f"{EBUG_BASE}/HandleMainEbug.asp?BugCode={bug_code}",
        }
    )

    url = f"{CONTENT_URL}?BugCode={bug_code}&IsFromMail="
    resp = session.get(url, timeout=30)
    resp.raise_for_status()

    # Detect redirect to login page
    if "ecl.cyberlink.com/login" in resp.url or "Login" in resp.url:
        print(
            "ERROR: Redirected to login page. Please log into https://ecl.cyberlink.com "
            f"in your {browser} browser and try again.",
            file=sys.stderr,
        )
        sys.exit(1)

    return session, resp.text


def fetch_image(session: requests.Session, url: str) -> bytes:
    """Download an image using the authenticated session."""
    resp = session.get(url, timeout=30)
    resp.raise_for_status()
    return resp.content
