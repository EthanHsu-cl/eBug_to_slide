import getpass
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

_KEYRING_SERVICE = "eBug-to-slide"
_KEYRING_USER_KEY = "_ntlm_username"


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


def _env_get(key: str, default: str = "") -> str:
    _load_env()
    return os.environ.get(key, default)


def _env_set(key: str, value: str) -> None:
    """Write or update a key=value line in the .env file."""
    lines: list[str] = []
    found = False
    if ENV_PATH.exists():
        with open(ENV_PATH, encoding="utf-8") as fh:
            for line in fh:
                if line.strip().startswith(f"{key}="):
                    lines.append(f"{key}={value}\n")
                    found = True
                else:
                    lines.append(line)
    if not found:
        lines.append(f"{key}={value}\n")
    with open(ENV_PATH, "w", encoding="utf-8") as fh:
        fh.writelines(lines)
    os.environ[key] = value


def load_browser_preference() -> str:
    return _env_get("EBUG_BROWSER", "auto")


def save_browser_preference(browser: str) -> None:
    _env_set("EBUG_BROWSER", browser)


def load_output_dir() -> str:
    return _env_get("EBUG_OUTPUT_DIR", "")


def save_output_dir(path: str) -> None:
    _env_set("EBUG_OUTPUT_DIR", path)


def load_last_bug_code() -> str:
    return _env_get("EBUG_LAST_BUG_CODE", "")


def save_last_bug_code(code: str) -> None:
    _env_set("EBUG_LAST_BUG_CODE", code)


_DEFAULT_OLLAMA_MODEL = "gemma4:e2b-it-q4_K_M"


def load_ollama_model() -> str:
    return _env_get("EBUG_OLLAMA_MODEL", _DEFAULT_OLLAMA_MODEL)


def save_ollama_model(model: str) -> None:
    _env_set("EBUG_OLLAMA_MODEL", model)


# ---------------------------------------------------------------------------
# NTLM credential storage (macOS Keychain / Windows Credential Manager via keyring)
# ---------------------------------------------------------------------------

# Sentinel: None = not yet fetched; ("", "") = fetched but user gave nothing
_ntlm_cache: tuple[str, str] | None = None
_ntlm_prompted: bool = False   # true once we've shown the prompt this run


def _get_ntlm_credentials() -> tuple[str, str] | None:
    """
    Return (username, password) or None if no credentials are available.
    - Prompts at most once per run (in-process cache).
    - Persists to macOS Keychain so subsequent runs skip the prompt entirely.
    - Returns None if the user left the prompt blank, suppressing NTLM for
      the rest of the run without re-prompting.
    """
    import keyring

    global _ntlm_cache, _ntlm_prompted

    # Return cached result (even if ("", "")) so we never prompt twice.
    if _ntlm_cache is not None:
        return _ntlm_cache if all(_ntlm_cache) else None

    # Try Keychain first.
    username = keyring.get_password(_KEYRING_SERVICE, _KEYRING_USER_KEY)
    if username:
        password = keyring.get_password(_KEYRING_SERVICE, username)
        if password:
            _ntlm_cache = (username, password)
            return _ntlm_cache

    # Prompt once.
    _ntlm_prompted = True
    print(
        "\nImage download requires your CyberLink network credentials "
        "(stored securely in system credential storage)."
    )
    try:
        username = input("Username (e.g. CYBERLINK\\yourname or yourname@cyberlink.com): ").strip()
        password = getpass.getpass("Password: ")
    except (EOFError, KeyboardInterrupt):
        username = ""
        password = ""

    if username and password:
        keyring.set_password(_KEYRING_SERVICE, _KEYRING_USER_KEY, username)
        keyring.set_password(_KEYRING_SERVICE, username, password)
        print("Credentials saved. You won't be prompted again.\n")
        _ntlm_cache = (username, password)
        return _ntlm_cache

    # User left prompt blank — cache the empty tuple so we don't ask again.
    _ntlm_cache = ("", "")
    return None


def _evict_keychain_credentials() -> None:
    """Remove NTLM credentials from Keychain only (keeps in-process cache intact)."""
    import keyring
    username = keyring.get_password(_KEYRING_SERVICE, _KEYRING_USER_KEY)
    if username:
        try:
            keyring.delete_password(_KEYRING_SERVICE, username)
        except Exception:
            pass
        try:
            keyring.delete_password(_KEYRING_SERVICE, _KEYRING_USER_KEY)
        except Exception:
            pass


def clear_ntlm_credentials() -> None:
    """Remove stored NTLM credentials from Keychain and in-process cache."""
    global _ntlm_cache, _ntlm_prompted
    _ntlm_cache = None
    _ntlm_prompted = False
    _evict_keychain_credentials()
    print("Credentials cleared from system credential storage.")


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


def fetch_image(session: requests.Session, url: str, browser: str = "auto") -> bytes:
    """
    Download an image using NTLM Windows auth (credentials stored in system credential storage).
    Falls back to the cookie-authenticated session if NTLM is not needed.
    Raises on failure so the caller can fall back to a local copy.
    """
    from requests_ntlm import HttpNtlmAuth

    # Try the plain cookie session first (avoids a credential prompt for non-NTLM endpoints).
    resp = session.get(url, timeout=30)
    if resp.status_code != 401:
        resp.raise_for_status()
        return resp.content

    # Server requires Windows auth — use NTLM with Keychain-stored credentials.
    creds = _get_ntlm_credentials()
    if creds is None:
        raise RuntimeError("No NTLM credentials provided")

    username, password = creds
    resp = session.get(url, auth=HttpNtlmAuth(username, password), timeout=30)
    if resp.status_code == 401:
        # Wrong credentials — evict from Keychain so user is re-prompted next run.
        # Keep in-process cache so we don't prompt again for remaining images this run.
        _evict_keychain_credentials()
        print(
            "\nERROR: NTLM authentication failed. Stored credentials have been cleared.\n"
            "Re-run the tool and enter the correct credentials when prompted.",
            file=sys.stderr,
        )
        raise RuntimeError("NTLM auth failed")
    resp.raise_for_status()
    return resp.content
