#!/usr/bin/env python3
"""
eBug to Slide — converts a CyberLink eBug report into a PowerPoint presentation.

Usage:
    python main.py <bug_code> [--output path/to/out.pptx] [--browser brave] [--debug]
    python main.py BUG1 BUG2 BUG3        # multiple codes → one combined .pptx
    python main.py "BUG1, BUG2, BUG3"   # comma-separated in quotes
    python main.py bugs.txt              # list file; codes may be one per line,
                                         # comma-separated, or space-separated
    python main.py                       # re-run the last used bug code

Multiple bug codes are combined into a single .pptx by default.
Use --separate to generate one file per bug code instead.

Supported list file formats: .txt, .json (array), .yaml/.yml (list)

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


def _parse_bug_codes_from_string(s: str) -> list[str]:
    """Split a raw string that may contain one or more bug codes separated by
    commas, spaces, or a mix of both, and return a list of parsed bug codes."""
    parts = re.split(r"[,\s]+", s.strip())
    return [_parse_bug_code(p) for p in parts if p.strip()]


def _combined_filename(codes: list[str]) -> str:
    """Build a filename stem for a combined multi-bug PPTX.

    When codes share a common prefix (e.g. 'PRP265213-'), abbreviates trailing
    codes to just their unique suffix:
        PRP265213-0053, PRP265213-0052, PRP265213-0051 → PRP265213-0053_0052_0051
    Falls back to joining full codes (or truncating with '_and_N_more') when
    no meaningful common prefix exists.
    """
    if len(codes) == 1:
        return codes[0]

    prefix = os.path.commonprefix(codes)
    # Snap back to the last '-' so we keep whole numeric segments intact
    if "-" in prefix:
        prefix = prefix[: prefix.rfind("-") + 1]

    if prefix.endswith("-"):
        # Clean segment boundary found — keep first code in full, strip prefix from rest
        parts = [codes[0]] + [c[len(prefix):] for c in codes[1:]]
        return "_".join(parts)

    # No common prefix — join in full, or truncate for long lists
    if len(codes) <= 3:
        return "_".join(codes)
    return f"{codes[0]}_and_{len(codes) - 1}_more"


def _load_bug_codes_from_file(path: str) -> list[str]:
    """Parse a .txt, .json, .yaml, or .yml file into a list of bug codes.

    .txt supports one code per line, comma-separated codes on a line, and
    space-separated codes on a line — or any mix of the above.
    """
    p = Path(path)
    suffix = p.suffix.lower()
    text = p.read_text(encoding="utf-8")

    if suffix == ".txt":
        codes: list[str] = []
        for ln in text.splitlines():
            ln = ln.strip()
            if not ln or ln.startswith("#"):
                continue
            # Split each line by comma and/or whitespace so all common formats work:
            # "BUG001"  |  "BUG001, BUG002"  |  "BUG001 BUG002"
            parts = re.split(r"[,\s]+", ln)
            codes.extend(p for p in parts if p)
        return codes

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


def _fetch_and_parse(
    bug_code: str,
    browser: str,
    debug: bool,
    ai_refine: bool = False,
    ollama_model: str = "",
):
    """Fetch and parse one bug code. Returns BugData or None on failure."""
    from scraper import fetch_bug
    from parser import parse

    debug_html = str(Path("debug") / f"{bug_code}_debug.html") if debug else None
    content_url = f"{EBUG_BASE}/HandleMainEbugContent.asp?BugCode={bug_code}&IsFromMail="

    print(f"Fetching bug {bug_code} ...")
    session, html = fetch_bug(bug_code, browser=browser)

    print("Parsing page ...")
    bug_data = parse(html, bug_code, session, content_url, debug_html_path=debug_html, browser=browser)

    if ai_refine:
        from refiner import refine_bug_data
        refine_bug_data(bug_data, ollama_model)

    return bug_data


def _run_single(
    bug_code: str,
    output_path: str,
    browser: str,
    debug: bool,
    ai_refine: bool = False,
    ollama_model: str = "",
) -> bool:
    """Fetch, parse, and generate a slide for one bug code. Returns True on success."""
    bug_data = _fetch_and_parse(bug_code, browser, debug, ai_refine, ollama_model)

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


def _run_combined(
    bug_codes: list[str],
    output_path: str,
    browser: str,
    debug: bool,
    ai_refine: bool = False,
    ollama_model: str = "",
) -> bool:
    """Fetch all bug codes and generate a single combined PPTX. Returns True on full success."""
    from slide_gen import generate_combined

    all_bug_data = []
    failures = []

    for bug_code in bug_codes:
        bug_data = _fetch_and_parse(bug_code, browser, debug, ai_refine, ollama_model)
        if not bug_data.image_steps:
            print(
                f"[{bug_code}] No image-tagged repro steps found. "
                "Run with --debug to save the raw HTML for inspection.",
                file=sys.stderr,
            )
            failures.append(bug_code)
        else:
            all_bug_data.append(bug_data)

    if not all_bug_data:
        return False

    output_dir = Path(output_path).parent
    output_dir.mkdir(parents=True, exist_ok=True)

    total = sum(len(d.image_steps) for d in all_bug_data)
    print(f"\nFound {total} image step(s) across {len(all_bug_data)} bug(s). Generating combined slides ...")
    generate_combined(all_bug_data, output_path)
    return not failures


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
        save_cookies_string,
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
        "url_or_codes",
        nargs="*",
        metavar="url_or_code",
        help=(
            "eBug URL(s), bug code(s), or path to a list file (.txt/.json/.yaml/.yml). "
            "Multiple codes may be space-separated on the command line or "
            "comma-separated inside a single quoted string. "
            f"Omit to reuse last code (currently: '{load_last_bug_code() or 'none'}')"
        ),
    )
    ap.add_argument(
        "--separate",
        action="store_true",
        help=(
            "Generate a separate .pptx for each bug code. "
            "Default when multiple codes are given: combine all into one file."
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
        "--cookies",
        metavar="COOKIE_STRING",
        default=None,
        help=(
            "Semicolon-separated cookie string copied from browser devtools "
            "(Network tab → any ecl.cyberlink.com request → Cookie header). "
            "Saved to .env as EBUG_COOKIES for future runs. "
            "Pass an empty string to clear: --cookies \"\""
        ),
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
    if args.cookies is not None:
        save_cookies_string(args.cookies)
        if args.cookies:
            print("Cookie string saved to .env (EBUG_COOKIES).")
        else:
            print("Cookie string cleared from .env.")

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
    raw_inputs: list[str] = args.url_or_codes

    if not raw_inputs:
        last = load_last_bug_code()
        if not last:
            ap.error("url_or_code is required (no last bug code saved in .env)")
        print(f"No input provided — reusing last bug code: {last}")
        raw_inputs = [last]

    effective_model = args.ollama_model if args.ollama_model else load_ollama_model()
    browser = load_browser_preference() if args.browser == "auto" else args.browser
    saved_dir = load_output_dir()

    # Flatten all tokens (file paths, URLs, single codes, comma-separated) into codes
    bug_codes: list[str] = []
    for token in raw_inputs:
        p = Path(token)
        if p.suffix.lower() in _LIST_EXTENSIONS and p.exists():
            codes = _load_bug_codes_from_file(token)
            if not codes:
                sys.exit(f"No bug codes found in {token}")
            bug_codes.extend(codes)
        else:
            # Handle comma-separated codes within a single token (spaces handled by shell nargs)
            for part in re.split(r",+", token):
                part = part.strip()
                if part:
                    bug_codes.append(_parse_bug_code(part))

    if not bug_codes:
        ap.error("No bug codes resolved from the provided input")

    if len(bug_codes) == 1 or args.separate:
        # One file per bug code
        failures = []
        for code in bug_codes:
            out = args.output if len(bug_codes) == 1 else None
            output_path = _resolve_output(out, code, saved_dir)
            ok = _run_single(code, output_path, browser, args.debug, args.ai_refine, effective_model)
            if ok:
                save_last_bug_code(code)
            else:
                failures.append(code)
        if failures:
            print(f"\nFailed: {', '.join(failures)}", file=sys.stderr)
            sys.exit(1)
    else:
        # Multiple bugs → one combined file (default)
        name = _combined_filename(bug_codes)
        output_path = _resolve_output(args.output, name, saved_dir)
        ok = _run_combined(bug_codes, output_path, browser, args.debug, args.ai_refine, effective_model)
        if ok:
            save_last_bug_code(bug_codes[-1])
        else:
            sys.exit(1)


if __name__ == "__main__":
    main()
