#!/usr/bin/env python3
"""
eBug to Slide — converts a CyberLink eBug report into a PowerPoint presentation.

Usage:
    python main.py <url_or_bug_code> [--output path/to/out.pptx] [--browser brave] [--debug]
    python main.py bugs.txt          # process a list of bug codes from a file
    python main.py                   # re-run the last used bug code

Supported list file formats: .txt (one per line), .json (array), .yaml/.yml (list)

Defaults are read from .env in the project folder:
    EBUG_BROWSER=brave
    EBUG_OUTPUT_DIR=~/Desktop/Reports
    EBUG_LAST_BUG_CODE=PRP265213-0053

Save defaults with --save-browser and --save-output-dir, or edit .env directly.
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

EBUG_BASE = "https://ecl.cyberlink.com/Ebug/eBugHandle"
BUG_CODE_RE = re.compile(r"BugCode=([A-Z0-9\-]+)", re.IGNORECASE)

_LIST_EXTENSIONS = {".txt", ".json", ".yaml", ".yml"}


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


def _load_bug_codes_from_file(path: str) -> list[str]:
    """Parse a .txt, .json, .yaml, or .yml file into a list of bug codes."""
    p = Path(path)
    suffix = p.suffix.lower()
    text = p.read_text(encoding="utf-8")

    if suffix == ".txt":
        return [
            ln.strip()
            for ln in text.splitlines()
            if ln.strip() and not ln.strip().startswith("#")
        ]

    if suffix == ".json":
        data = json.loads(text)
        if isinstance(data, list):
            return [str(x).strip() for x in data if x]
        raise ValueError("JSON file must contain a top-level array of bug codes")

    if suffix in (".yaml", ".yml"):
        try:
            import yaml
        except ImportError:
            sys.exit("pyyaml is required for YAML input. Run: pip install pyyaml")
        data = yaml.safe_load(text)
        if isinstance(data, list):
            return [str(x).strip() for x in data if x]
        raise ValueError("YAML file must contain a top-level list of bug codes")

    raise ValueError(f"Unsupported file type '{suffix}'. Use .txt, .json, .yaml, or .yml")


def _run_single(
    bug_code: str,
    output_path: str,
    browser: str,
    debug: bool,
    ai_refine: bool = False,
    ollama_model: str = "",
) -> bool:
    """Fetch, parse, and generate a slide for one bug code. Returns True on success."""
    debug_html = f"{bug_code}_debug.html" if debug else None
    content_url = f"{EBUG_BASE}/HandleMainEbugContent.asp?BugCode={bug_code}&IsFromMail="

    print(f"Fetching bug {bug_code} ...")
    from scraper import fetch_bug
    session, html = fetch_bug(bug_code, browser=browser)

    print("Parsing page ...")
    from parser import parse
    bug_data = parse(html, bug_code, session, content_url, debug_html_path=debug_html, browser=browser)

    if ai_refine:
        from refiner import refine_bug_data
        refine_bug_data(bug_data, ollama_model)

    if not bug_data.image_steps:
        print(
            f"[{bug_code}] No image-tagged repro steps found. "
            "Run with --debug to save the raw HTML for inspection.",
            file=sys.stderr,
        )
        return False

    output_dir = Path(output_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Found {len(bug_data.image_steps)} image step(s). Generating slides ...")
    from slide_gen import generate
    generate(bug_data, output_path)
    return True


def main() -> None:
    if len(sys.argv) == 1:
        from gui import launch_gui
        launch_gui()
        return

    from scraper import (
        clear_ntlm_credentials,
        load_browser_preference,
        load_last_bug_code,
        load_ollama_model,
        load_output_dir,
        save_browser_preference,
        save_last_bug_code,
        save_ollama_model,
        save_output_dir,
    )

    ap = argparse.ArgumentParser(
        description="Convert a CyberLink eBug report to a PowerPoint presentation.",
        epilog=(
            "Persistent defaults are stored in .env next to main.py.\n"
            "You can also edit .env directly:\n"
            "  EBUG_BROWSER=brave\n"
            "  EBUG_OUTPUT_DIR=~/Desktop/Reports\n"
            "  EBUG_LAST_BUG_CODE=PRP265213-0053"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument(
        "url_or_code",
        nargs="?",
        help=(
            "eBug URL, bug code (e.g. PRP265213-0053), or path to a list file "
            f"(.txt/.json/.yaml/.yml). Omit to reuse last code "
            f"(currently: '{load_last_bug_code() or 'none'}')"
        ),
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
        help="Remove stored NTLM credentials from system credential storage and exit",
    )
    ap.add_argument(
        "--debug",
        action="store_true",
        help="Save raw HTML to <bug_code>_debug.html for inspection",
    )
    ap.add_argument(
        "--ai-refine",
        action="store_true",
        help="Refine title and section text with a local Ollama model before generating slides",
    )
    ap.add_argument(
        "--ollama-model",
        metavar="MODEL",
        default=None,
        help=(
            "Ollama model to use for --ai-refine (overrides saved preference). "
            f"Saved: '{load_ollama_model()}'"
        ),
    )
    ap.add_argument(
        "--set-ollama-model",
        metavar="MODEL",
        default=None,
        help="Save an Ollama model name as the default for --ai-refine (stored in .env)",
    )
    args = ap.parse_args()

    if args.clear_credentials:
        clear_ntlm_credentials()
        return

    # --- Persist preferences ---
    if args.save_browser and args.browser != "auto":
        save_browser_preference(args.browser)
        print(f"Browser preference saved: {args.browser}")

    if args.save_output_dir is not None:
        save_output_dir(args.save_output_dir)
        print(f"Output directory saved: {args.save_output_dir}")

    if args.set_ollama_model is not None:
        save_ollama_model(args.set_ollama_model)
        print(f"Ollama model preference saved: {args.set_ollama_model}")

    # --- Resolve input ---
    raw_input = args.url_or_code

    # Fall back to last bug code when nothing is provided
    if not raw_input:
        last = load_last_bug_code()
        if not last:
            ap.error("url_or_code is required (no last bug code saved in .env)")
        print(f"No input provided — reusing last bug code: {last}")
        raw_input = last

    effective_model = args.ollama_model if args.ollama_model else load_ollama_model()

    # Check if the input is a list file
    p = Path(raw_input)
    if p.suffix.lower() in _LIST_EXTENSIONS and p.exists():
        bug_codes = _load_bug_codes_from_file(raw_input)
        if not bug_codes:
            sys.exit(f"No bug codes found in {raw_input}")
        print(f"Processing {len(bug_codes)} bug code(s) from {p.name} ...")
        saved_dir = load_output_dir()
        browser = load_browser_preference() if args.browser == "auto" else args.browser
        failures = []
        for code in bug_codes:
            output_path = _resolve_output(None, code, saved_dir)
            ok = _run_single(code, output_path, browser, args.debug, args.ai_refine, effective_model)
            if ok:
                save_last_bug_code(code)
            else:
                failures.append(code)
        if failures:
            print(f"\nFailed: {', '.join(failures)}", file=sys.stderr)
            sys.exit(1)
        return

    # Single bug code / URL
    bug_code = _parse_bug_code(raw_input)
    output_path = _resolve_output(args.output, bug_code, load_output_dir())
    browser = load_browser_preference() if args.browser == "auto" else args.browser

    ok = _run_single(bug_code, output_path, browser, args.debug, args.ai_refine, effective_model)
    if not ok:
        sys.exit(1)
    save_last_bug_code(bug_code)


if __name__ == "__main__":
    main()
