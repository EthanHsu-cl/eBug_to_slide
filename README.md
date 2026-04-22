# eBug to Slide

Converts a CyberLink eBug report into a PowerPoint presentation automatically.  
Each repro step that references a screenshot becomes one slide — the image fills the main area and the step text appears in the annotation panel.

---

## How it works

The tool:

1. Reads your existing browser session (no credentials stored) to authenticate with the eBug system
2. Fetches the bug report from `HandleMainEbugContent.asp`
3. Parses repro steps for image tags in the format `{type:Current, step:1, file:image1.png}`
4. Downloads each referenced screenshot
5. Generates a `.pptx` based on the slide template — one slide per image-tagged step

### Repro step format

Steps without an image tag are skipped. Steps with a tag produce one slide:

```text
Repro Steps:
1. Launch PMO
2. Switch to AI Marketing Post
3. Select Use Your Image / Create with AI
4. Select a result {type:Current, step:1, file:image1.png}   ← slide 1
5. Enter edit room  {type:Current, step:2, file:image2.png}   ← slide 2
6. Click AI Advisor to open panel
7. Edit on canvas to make it dirty {type:Current, step:3, file:image3.png}  ← slide 3
```

Supported types: `Current`, `Reference`, `Proposal`

### Slide layout (from template)

| Area | Position | Purpose |
| --- | --- | --- |
| Title | top-left | Bug short description |
| Type label | below title | `Current: <version>`, `Reference: …`, `Proposal: …` |
| Version info | top-right | Fix version / sprint info |
| Main image | center (large) | Screenshot, scaled to fit |
| Step annotation | off right edge* | Step text from the repro list |

\* The annotation box is intentionally positioned just past the slide's right edge — this matches the UX team's template design.

---

## Setup

```bash
# 1. Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Make sure you're logged into https://ecl.cyberlink.com in your browser
```

### Requirements

- Python 3.10+
- One of: Brave, Chrome, Edge, Safari, Firefox, Chromium, Opera, or Vivaldi  
  (must be signed in to `ecl.cyberlink.com`)

---

## Usage

```bash
source venv/bin/activate

# Basic — auto-detects your browser
python main.py PRP265213-0053

# Full URL also works
python main.py "https://ecl.cyberlink.com/Ebug/eBugHandle/HandleMainEbug.asp?BugCode=PRP265213-0053"

# Re-run the last used bug code (no argument needed)
python main.py

# Specify output path
python main.py PRP265213-0053 --output ~/Desktop/my_report.pptx

# Save a default output directory so you never have to specify --output again
python main.py --save-output-dir ~/Desktop/Reports

# Specify browser explicitly
python main.py PRP265213-0053 --browser brave

# Save browser preference so you never have to specify it again
python main.py PRP265213-0053 --browser brave --save-browser

# Save raw HTML for debugging if parsing fails
python main.py PRP265213-0053 --debug

# Remove stored NTLM credentials from macOS Keychain
python main.py --clear-credentials
```

### Batch processing

Pass a file containing multiple bug codes to process them all in one run.  
Each bug produces its own `.pptx` in the configured output directory.

**`.txt`** — one bug code per line (`#` lines are ignored):

```text
# Sprint 42
PRP265213-0053
PRP265213-0054
PRP265213-0055
```

**`.json`** — top-level array:

```json
["PRP265213-0053", "PRP265213-0054", "PRP265213-0055"]
```

**`.yaml` / `.yml`** — top-level list (requires `pyyaml`, included in `requirements.txt`):

```yaml
- PRP265213-0053
- PRP265213-0054
- PRP265213-0055
```

```bash
python main.py bugs.txt
python main.py bugs.json
python main.py bugs.yaml
```

### Options

| Flag | Short | Default | Description |
| --- | --- | --- | --- |
| `--output` | `-o` | `<EBUG_OUTPUT_DIR>/<bug_code>.pptx` | Output file path (single bug only) |
| `--save-output-dir DIR` | | | Save DIR as the default output location (stored in `.env`) |
| `--browser` | `-b` | `auto` (or saved preference) | Browser to read cookies from |
| `--save-browser` | | | Persist `--browser` to `.env` for future runs |
| `--clear-credentials` | | | Remove stored NTLM credentials from macOS Keychain and exit |
| `--debug` | | | Save raw HTML to `<bug_code>_debug.html` for inspection |

### Browser auto-detection order

`brave` → `chrome` → `edge` → `safari`

Use `--save-browser` to set a permanent default and skip detection entirely.

---

## Persistent settings (`.env`)

Defaults are stored in `.env` next to `main.py`. You can edit this file directly or use the CLI flags above to update it.

| Key | Set via | Description |
| --- | --- | --- |
| `EBUG_BROWSER` | `--save-browser` | Default browser for cookie extraction |
| `EBUG_OUTPUT_DIR` | `--save-output-dir` | Default output directory |
| `EBUG_LAST_BUG_CODE` | auto-saved | Last successfully processed bug code; used when no argument is given |

---

## File structure

```text
eBug_to_slide/
├── main.py               # CLI entry point
├── scraper.py            # Cookie extraction + HTTP fetch
├── parser.py             # HTML parsing + image download
├── slide_gen.py          # PPTX generation (python-pptx)
├── requirements.txt      # Python dependencies
├── Template/
│   └── New Layout.pptx   # Slide template (do not rename or move)
└── .env                  # Auto-created; stores preferences and last bug code
```

---

## Troubleshooting

**"Could not extract cookies from any browser"**  
Log into `https://ecl.cyberlink.com` in your browser, then retry. Use `--browser <name>` to target a specific browser.

**"Redirected to login page"**  
Your session has expired. Log in again in the browser.

**"No image-tagged repro steps found"**  
Run with `--debug` to save the raw HTML. Open `<bug_code>_debug.html` and check that repro steps are present and use the `{type:…, step:…, file:…}` format.

**Images fail to download (401)**  
The eBug page uses cookie auth, but the image download endpoint (`DownloadeBugFile.ashx`) uses Windows Authentication — a separate layer that only the browser can handle transparently. When this happens, the tool prints the exact download URLs and a target folder path. Open each URL in Brave (which handles auth automatically), save the file with its original name, then re-run. The tool checks `images/<bug_code>/` for local copies before attempting a download.

**Images missing from slides**  
The tool downloads images from the Upload File table in the eBug report. If a download fails, check the warning output — it will show the exact URL attempted. Run with `--debug` to inspect the raw HTML.
