#!/usr/bin/env python3
"""
eBug to Slide — converts a CyberLink eBug report into a PowerPoint presentation.

Usage:
    python main.py <url_or_bug_code> [--output path/to/out.pptx] [--browser auto|chrome|brave|edge|safari] [--debug]

Examples:
    python main.py PRP265213-0053
    python main.py https://ecl.cyberlink.com/Ebug/eBugHandle/HandleMainEbug.asp?BugCode=PRP265213-0053
    python main.py PRP265213-0053 --output ~/Desktop/my_report.pptx --browser brave
"""

import argparse
import os
import re
import sys

EBUG_BASE = "https://ecl.cyberlink.com/Ebug/eBugHandle"
BUG_CODE_RE = re.compile(r"BugCode=([A-Z0-9\-]+)", re.IGNORECASE)


def _parse_bug_code(url_or_code: str) -> str:
    m = BUG_CODE_RE.search(url_or_code)
    if m:
        return m.group(1)
    # Assume it's a bare code like PRP265213-0053
    return url_or_code.strip()


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Convert a CyberLink eBug report to a PowerPoint presentation."
    )
    ap.add_argument(
        "url_or_code",
        help="eBug URL or bug code (e.g. PRP265213-0053)",
    )
    ap.add_argument(
        "--output", "-o",
        default=None,
        help="Output .pptx path (default: ./<bug_code>.pptx)",
    )
    ap.add_argument(
        "--browser", "-b",
        default="auto",
        choices=["auto", "brave", "chrome", "edge", "safari", "firefox", "chromium", "opera", "vivaldi"],
        help="Browser to extract session cookies from (default: auto, or saved preference)",
    )
    ap.add_argument(
        "--save-browser",
        action="store_true",
        help="Save --browser as the default for future runs (stored in .env)",
    )
    ap.add_argument(
        "--debug",
        action="store_true",
        help="Save raw HTML to <bug_code>_debug.html for inspection",
    )
    args = ap.parse_args()

    bug_code = _parse_bug_code(args.url_or_code)
    output_path = args.output or f"{bug_code}.pptx"
    debug_html = f"{bug_code}_debug.html" if args.debug else None

    content_url = f"{EBUG_BASE}/HandleMainEbugContent.asp?BugCode={bug_code}&IsFromMail="

    from scraper import fetch_bug, save_browser_preference
    if args.save_browser and args.browser != "auto":
        save_browser_preference(args.browser)
        print(f"Browser preference saved: {args.browser}")

    print(f"Fetching bug {bug_code} ...")

    session, html = fetch_bug(bug_code, browser=args.browser)

    print("Parsing page ...")
    from parser import parse
    bug_data = parse(html, bug_code, session, content_url, debug_html_path=debug_html)

    if not bug_data.image_steps:
        print(
            "No image-tagged repro steps found.\n"
            "Run with --debug to save the raw HTML for inspection.",
            file=sys.stderr,
        )
        sys.exit(1)

    print(f"Found {len(bug_data.image_steps)} image step(s). Generating slides ...")
    from slide_gen import generate
    generate(bug_data, output_path)


if __name__ == "__main__":
    main()
