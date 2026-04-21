import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from scraper import fetch_image

# Matches {type:Current, step:1, file:image1.png}
# "step:" may be wrapped in <b> tags in the raw HTML, but after text extraction it's a plain string.
_IMAGE_TAG_RE = re.compile(
    r"\{type:\s*(\w+)\s*,\s*step:\s*(\d+)\s*,\s*file:\s*([^\}]+?)\s*\}"
)

# Upload File section uses this URL pattern for image downloads
_DOWNLOAD_HREF_RE = re.compile(r"DownloadeBugFile\.ashx", re.IGNORECASE)

TYPE_ORDER = {"Current": 1, "Reference": 2, "Proposal": 3}


@dataclass
class ImageStep:
    step_num: int       # sequential number from the {step:N} tag
    text: str           # step line text with the {…} tag stripped
    img_type: str       # "Current", "Reference", or "Proposal"
    image_file: str     # filename, e.g. "image1.png"
    image_data: bytes = field(default=b"", repr=False)


@dataclass
class BugData:
    code: str
    title: str
    version_info: str   # e.g. "v8.20" — used in version box
    image_steps: list[ImageStep]
    section_texts: dict[str, str] = field(default_factory=dict)  # {"Current": "...", ...}


# ---------------------------------------------------------------------------
# HTML → text helpers
# ---------------------------------------------------------------------------

def _html_to_text(tag) -> str:
    """
    Extract text treating <br> as newline and all other inline tags as
    transparent (concatenated without separator).

    The eBug page wraps "step:" in <b> tags inside the {type:...} metadata,
    so using BeautifulSoup's get_text("\n") would split the tag across lines.
    This function avoids that by only inserting newlines at actual <br> tags.
    """
    html = str(tag)
    html = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    return BeautifulSoup(html, "lxml").get_text("")


# ---------------------------------------------------------------------------
# eBug-specific field extractors
# ---------------------------------------------------------------------------

def _find_short_description(soup: BeautifulSoup) -> str:
    """
    The short description lives in a <td colspan=5> like:
      <font size=2><b>Short Description: </b></font>
      <font color='red'><b>TITLE TEXT</b></font>
    """
    for td in soup.find_all("td"):
        if "Short Description" not in td.get_text():
            continue
        red = td.find("font", attrs={"color": re.compile(r"^red$", re.IGNORECASE)})
        if red:
            return red.get_text("", strip=True)
    return ""


def _find_version(soup: BeautifulSoup) -> str:
    """
    Version lives inline in a <td>:
      Version: <font color='darkblue'>8.20</font>
    The next sibling td is Build No, not the value we want.
    """
    for td in soup.find_all("td"):
        raw = td.get_text()
        if re.match(r"\s*Version\s*:", raw, re.IGNORECASE):
            blue = td.find("font", attrs={"color": re.compile(r"darkblue", re.IGNORECASE)})
            if blue:
                val = blue.get_text("", strip=True)
                if val:
                    return f"v{val}"
    return ""


def _find_repro_text(soup: BeautifulSoup) -> tuple[str, str]:
    """
    Return (text, raw_html) of the repro/description block.
    The raw HTML is needed to extract section texts from <Current>...</Current> markers.
    """
    for td in soup.find_all("td"):
        classes = td.get("class") or []
        if "NoLine2" not in classes and "noline2" not in [c.lower() for c in classes]:
            continue
        text = _html_to_text(td)
        if _IMAGE_TAG_RE.search(text):
            return text, str(td)

    # Fallback: any container whose text matches the image tag pattern
    for tag in soup.find_all(["td", "div", "pre"]):
        text = _html_to_text(tag)
        if _IMAGE_TAG_RE.search(text):
            return text, str(tag)

    return "", ""


_SECTION_OPEN_RE = re.compile(r"^\s*<(Current|Reference|Proposal)>\s*$", re.IGNORECASE)
_SECTION_CLOSE_RE = re.compile(r"^\s*</(Current|Reference|Proposal)>\s*$", re.IGNORECASE)


def _find_section_texts(repro_text: str) -> dict[str, str]:
    """
    Extract description text for each section type from the processed repro text.
    The text contains HTML-decoded markers like literal <Current>...</Current>.
    Some sections have no closing tag (Reference, Proposal), so we collect until
    the next opening tag or end of text.
    """
    texts: dict[str, str] = {}
    current_section: str | None = None
    section_lines: list[str] = []

    for line in repro_text.splitlines():
        if m := _SECTION_OPEN_RE.match(line):
            if current_section and section_lines:
                texts.setdefault(current_section, " ".join(section_lines))
            current_section = m.group(1).capitalize()
            section_lines = []
        elif _SECTION_CLOSE_RE.match(line):
            if current_section and section_lines:
                texts.setdefault(current_section, " ".join(section_lines))
            current_section = None
            section_lines = []
        elif current_section:
            clean = _IMAGE_TAG_RE.sub("", line).strip()
            if clean:
                section_lines.append(clean)

    if current_section and section_lines:
        texts.setdefault(current_section, " ".join(section_lines))

    return texts


def _build_image_url_map(soup: BeautifulSoup) -> dict[str, str]:
    """
    The Upload File table lists images as:
      <a href="https://ecl.cyberlink.com/dc/support/DownloadeBugFile.ashx?d=ID">
        <u>image1.png</u>
      </a>
    Build filename.lower() → absolute download URL.
    """
    url_map: dict[str, str] = {}
    for a in soup.find_all("a", href=_DOWNLOAD_HREF_RE):
        href = a.get("href", "")
        filename = a.get_text(strip=True).lower()
        if filename and href:
            url_map[filename] = href
    return url_map


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse(
    html: str,
    bug_code: str,
    session: requests.Session,
    page_url: str,
    debug_html_path: str | None = None,
    browser: str = "auto",
) -> BugData:
    if debug_html_path:
        with open(debug_html_path, "w", encoding="utf-8") as fh:
            fh.write(html)

    soup = BeautifulSoup(html, "lxml")

    title = _find_short_description(soup) or bug_code
    version_info = _find_version(soup) or "PX, fix in X/X"
    img_url_map = _build_image_url_map(soup)

    repro_text, repro_raw = _find_repro_text(soup)
    if not repro_text:
        debug_path = debug_html_path or f"{bug_code}_debug.html"
        print(
            f"WARNING: Could not find repro steps.\n"
            f"Raw HTML saved to {debug_path} for inspection.",
            file=sys.stderr,
        )
        if not debug_html_path:
            with open(debug_path, "w", encoding="utf-8") as fh:
                fh.write(html)
        return BugData(code=bug_code, title=title,
                       version_info=version_info, image_steps=[])

    section_texts = _find_section_texts(repro_text)

    image_steps: list[ImageStep] = []
    missing: list[tuple[str, str]] = []   # (filename, download_url) for failed downloads
    local_dir = Path(f"images/{bug_code}")
    base_url = page_url.rsplit("/", 1)[0] + "/"

    for line in repro_text.splitlines():
        m = _IMAGE_TAG_RE.search(line)
        if not m:
            continue

        img_type = m.group(1)
        step_num = int(m.group(2))
        image_file = m.group(3).strip()
        clean_text = _IMAGE_TAG_RE.sub("", line).strip()

        img_url = (
            img_url_map.get(image_file.lower())
            or f"{base_url}images/{bug_code}/{image_file}"
        )

        # 1. Try local copy first (user may have manually saved images here)
        local_path = local_dir / image_file
        if local_path.exists():
            image_data = local_path.read_bytes()
        else:
            # 2. Try downloading
            image_data = b""
            try:
                image_data = fetch_image(session, img_url, browser)
            except Exception:
                missing.append((image_file, img_url))

        image_steps.append(ImageStep(
            step_num=step_num,
            text=clean_text,
            img_type=img_type,
            image_file=image_file,
            image_data=image_data,
        ))

    if missing:
        local_dir.mkdir(parents=True, exist_ok=True)
        print(
            f"\n{len(missing)} image(s) could not be downloaded automatically "
            f"(the download endpoint requires Windows auth).\n"
            f"To embed them, open each link in Brave, save to:\n"
            f"  {local_dir.resolve()}/\n"
            f"then re-run the tool.\n",
            file=sys.stderr,
        )
        for filename, url in missing:
            print(f"  {filename}  →  {url}", file=sys.stderr)
        print(file=sys.stderr)

    image_steps.sort(key=lambda s: (TYPE_ORDER.get(s.img_type, 99), s.step_num))

    return BugData(
        code=bug_code,
        title=title,
        version_info=version_info,
        image_steps=image_steps,
        section_texts=section_texts,
    )
