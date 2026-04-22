# Claude Code Instructions

## Running Scripts

Always activate the virtual environment before running any Python script:

```bash
source venv/bin/activate
python main.py  # or any other script
```

## Making Changes

When adding features or modifying existing behavior, update [README.md](README.md) to reflect the changes before considering the task complete.

## Python Version

Requires Python 3.10+. The codebase uses `str | None` union syntax and other 3.10+ features.

## File Responsibilities

| File | Role |
| --- | --- |
| [main.py](main.py) | CLI entry point — argument parsing, orchestration |
| [scraper.py](scraper.py) | Cookie extraction, HTTP fetch, `.env` preferences |
| [parser.py](parser.py) | HTML parsing, image downloading |
| [slide_gen.py](slide_gen.py) | PPTX generation via python-pptx |

Keep logic in the appropriate module; don't add scraping logic to `slide_gen.py`, etc.

## Template File

`Template/New Layout.pptx` is referenced by exact path in [slide_gen.py](slide_gen.py). Do not rename or move it.

## Generated Files — Do Not Commit

- `.env` — auto-created; stores user preferences (`EBUG_BROWSER`, `EBUG_OUTPUT_DIR`, `EBUG_LAST_BUG_CODE`)
- `images/` — downloaded screenshots; re-fetched on each run
- `*_debug.html` — produced by `--debug` flag; temporary inspection files
