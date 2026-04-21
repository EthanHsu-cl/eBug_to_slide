#!/usr/bin/env python3
"""
eBug to Slide — converts a CyberLink eBug report into a PowerPoint presentation.

Usage:
    python main.py <url_or_bug_code> [--output path/to/out.pptx] [--browser brave] [--debug]

Defaults are read from .env in the project folder:
    EBUG_BROWSER=brave
    EBUG_OUTPUT_DIR=~/Desktop/Reports

Save defaults with --save-browser and --save-output-dir, or edit .env directly.
"""

import argparse
import os
import re
import sys
from pathlib import Path

EBUG_BASE = "https://ecl.cyberlink.com/Ebug/eBugHandle"
BUG_CODE_RE = re.compile(r"BugCode=([A-Z0-9\-]+)", re.IGNORECASE)


def _parse_bug_code(url_or_code: str) -> str:
    m = BUG_CODE_RE.search(url_or_code)
    if m:
        return m.group(1)
    return url_or_code.strip()


def _resolve_output(args_output: str | None, bug_code: str, saved_dir: str) -> str:
    if args_output:
        return args_output
    base = Path(saved_dir).expanduser() if saved_dir else Path(".")
    return str(base / f"{bug_code}.pptx")


def main() -> None:
    from scraper import (
        clear_ntlm_credentials,
        load_browser_preference,
        load_output_dir,
        save_browser_preference,
        save_output_dir,
    )

    ap = argparse.ArgumentParser(
        description="Convert a CyberLink eBug report to a PowerPoint presentation.",
        epilog=(
            "Persistent defaults are stored in .env next to main.py.\n"
            "You can also edit .env directly:\n"
            "  EBUG_BROWSER=brave\n"
            "  EBUG_OUTPUT_DIR=~/Desktop/Reports"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "url_or_code",
        nargs="?",
        help="eBug URL or bug code (e.g. PRP265213-0053)",
    )
    ap.add_argument(
        "--output", "-o",
        default=None,
        help="Output .pptx path. Overrides EBUG_OUTPUT_DIR. "
             f"Default: <EBUG_OUTPUT_DIR>/<bug_code>.pptx (currently: '{load_output_dir() or '.'}/')",
    )
    ap.add_argument(
        "--save-output-dir",
        metavar="DIR",
        default=None,
        help="Save a directory as the default output location (stored in .env)",
    )
    ap.add_argument(
        "--browser", "-b",
        default="auto",
        choices=["auto", "brave", "chrome", "edge", "safari", "firefox", "chromium", "opera", "vivaldi"],
        help=f"Browser to extract session cookies from "
             f"(default: auto, saved: '{load_browser_preference()}')",
    )
    ap.add_argument(
        "--save-browser",
        action="store_true",
        help="Save --browser as the default for future runs (stored in .env)",
    )
    ap.add_argument(
        "--clear-credentials",
        action="store_true",
        help="Remove stored NTLM credentials from macOS Keychain and exit",
    )
    ap.add_argument(
        "--debug",
        action="store_true",
        help="Save raw HTML to <bug_code>_debug.html for inspection",
    )
    args = ap.parse_args()

    if args.clear_credentials:
        clear_ntlm_credentials()
        return

    if not args.url_or_code:
        ap.error("url_or_code is required")

    # --- Persist preferences ---
    if args.save_browser and args.browser != "auto":
        save_browser_preference(args.browser)
        print(f"Browser preference saved: {args.browser}")

    if args.save_output_dir is not None:
        save_output_dir(args.save_output_dir)
        print(f"Output directory saved: {args.save_output_dir}")

    # --- Resolve paths ---
    bug_code = _parse_bug_code(args.url_or_code)
    output_path = _resolve_output(args.output, bug_code, load_output_dir())
    debug_html = f"{bug_code}_debug.html" if args.debug else None

    # Ensure output directory exists
    output_dir = Path(output_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)

    content_url = f"{EBUG_BASE}/HandleMainEbugContent.asp?BugCode={bug_code}&IsFromMail="

    print(f"Fetching bug {bug_code} ...")
    from scraper import fetch_bug
    session, html = fetch_bug(bug_code, browser=args.browser)

    print("Parsing page ...")
    from parser import parse
    bug_data = parse(html, bug_code, session, content_url, debug_html_path=debug_html, browser=args.browser)

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
