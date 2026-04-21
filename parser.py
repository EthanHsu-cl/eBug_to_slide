import re
import sys
from dataclasses import dataclass, field
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from scraper import fetch_image

# Matches {type:Current, step:1, file:image1.png} with flexible whitespace
_IMAGE_TAG_RE = re.compile(
    r"\{type:\s*(\w+)\s*,\s*step:\s*(\d+)\s*,\s*file:\s*([^\}]+?)\s*\}"
)


@dataclass
class ImageStep:
    step_num: int       # overall line number in the repro steps
    text: str           # step text with the {…} tag stripped
    img_type: str       # "Current", "Reference", or "Proposal"
    image_file: str     # filename, e.g. "image1.png"
    image_data: bytes = field(default=b"", repr=False)


@dataclass
class BugData:
    code: str
    title: str
    version_info: str
    image_steps: list[ImageStep]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _text_of(tag) -> str:
    return tag.get_text(" ", strip=True) if tag else ""


def _find_field_value(soup: BeautifulSoup, *labels: str) -> str:
    """
    Search for a table cell whose text matches any of *labels* and return
    the text of the adjacent/sibling cell.  Works for both <th>/<td> pairs
    and <td label> + <td value> layouts common in old ASP pages.
    """
    for label in labels:
        pattern = re.compile(label, re.IGNORECASE)
        cell = soup.find(lambda t: t.name in ("td", "th") and pattern.search(t.get_text()))
        if cell:
            sibling = cell.find_next_sibling("td")
            if sibling:
                return sibling.get_text(" ", strip=True)
    return ""


def _find_repro_block(soup: BeautifulSoup) -> str:
    """
    Return the raw text of the Repro Steps block. Tries several heuristics
    common to ASP-based bug-tracker pages.
    """
    # 1. Look for a label cell that says "Repro" / "Steps" and take its sibling/next block
    for label_text in ("Repro Step", "Repro", "Steps to Reproduce", "How to Reproduce"):
        pattern = re.compile(label_text, re.IGNORECASE)
        cell = soup.find(lambda t: t.name in ("td", "th", "div", "span", "label")
                         and pattern.search(t.get_text()))
        if cell:
            # Look for a following <td> or <div> containing numbered steps
            parent_row = cell.find_parent("tr")
            if parent_row:
                next_row = parent_row.find_next_sibling("tr")
                if next_row:
                    return next_row.get_text("\n", strip=True)
            sibling = cell.find_next_sibling()
            if sibling:
                return sibling.get_text("\n", strip=True)

    # 2. Fall back: find any block of text that contains the image tag pattern
    for tag in soup.find_all(["td", "div", "pre", "textarea"]):
        text = tag.get_text("\n")
        if _IMAGE_TAG_RE.search(text):
            return text

    return ""


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse(
    html: str,
    bug_code: str,
    session: requests.Session,
    page_url: str,
    debug_html_path: str | None = None,
) -> BugData:
    """
    Parse the eBug content page and return a BugData object.
    Downloads images via *session*.  If *debug_html_path* is given, the raw
    HTML is saved there regardless of parse success.
    """
    if debug_html_path:
        with open(debug_html_path, "w", encoding="utf-8") as fh:
            fh.write(html)

    soup = BeautifulSoup(html, "lxml")

    # --- Title ---
    title = (
        _find_field_value(soup, "Title", "Subject", "Summary", "Short Description")
        or _text_of(soup.find("title"))
        or bug_code
    )

    # --- Version / fix-version ---
    version_info = _find_field_value(
        soup, "Fix In", "Fix Version", "Fixed In", "Target Version", "Version"
    ) or "PX, fix in X/X"

    # --- Repro steps block ---
    repro_text = _find_repro_block(soup)
    if not repro_text:
        print(
            f"WARNING: Could not find repro steps in the page.\n"
            f"Raw HTML saved to {debug_html_path or bug_code + '_debug.html'} for inspection.",
            file=sys.stderr,
        )
        if not debug_html_path:
            path = f"{bug_code}_debug.html"
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(html)

    # --- Parse image-tagged steps ---
    image_steps: list[ImageStep] = []
    base_url = page_url.rsplit("/", 1)[0] + "/"

    # Build a map from image filename → <img> src found in the page
    img_srcs: dict[str, str] = {}
    for img_tag in soup.find_all("img"):
        src = img_tag.get("src", "")
        if src:
            filename = src.rsplit("/", 1)[-1].split("?")[0]
            img_srcs[filename.lower()] = urljoin(base_url, src)

    for line in repro_text.splitlines():
        m = _IMAGE_TAG_RE.search(line)
        if not m:
            continue

        img_type, step_num_str, image_file = m.group(1), m.group(2), m.group(3).strip()
        clean_text = _IMAGE_TAG_RE.sub("", line).strip()
        step_num = int(step_num_str)

        # Resolve image URL: prefer <img> src found in page, else construct URL
        img_url = img_srcs.get(image_file.lower()) or _guess_image_url(
            base_url, bug_code, image_file
        )

        image_data = b""
        if img_url:
            try:
                image_data = fetch_image(session, img_url)
            except Exception as exc:
                print(f"WARNING: Could not download {image_file} from {img_url}: {exc}", file=sys.stderr)

        image_steps.append(
            ImageStep(
                step_num=step_num,
                text=clean_text,
                img_type=img_type,
                image_file=image_file,
                image_data=image_data,
            )
        )

    return BugData(
        code=bug_code,
        title=title,
        version_info=version_info,
        image_steps=image_steps,
    )


def _guess_image_url(base_url: str, bug_code: str, filename: str) -> str:
    """Try a few common ASP attachment URL patterns."""
    candidates = [
        f"{base_url}images/{bug_code}/{filename}",
        f"{base_url}Upload/{bug_code}/{filename}",
        f"{base_url}Attachments/{bug_code}/{filename}",
        f"{base_url}{filename}",
    ]
    return candidates[0]  # caller will try and fall back gracefully
