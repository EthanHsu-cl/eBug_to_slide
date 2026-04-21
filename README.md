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
5. Generates a `.pptx` based on `UX ppt template.pptx` — one slide per image-tagged step

### Repro step format

Steps without an image tag are skipped. Steps with a tag produce one slide:

```
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
|---|---|---|
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

# Specify output path
python main.py PRP265213-0053 --output ~/Desktop/my_report.pptx

# Specify browser explicitly
python main.py PRP265213-0053 --browser brave

# Save browser preference so you never have to specify it again
python main.py PRP265213-0053 --browser brave --save-browser

# Save raw HTML for debugging if parsing fails
python main.py PRP265213-0053 --debug
```

### Options

| Flag | Short | Default | Description |
|---|---|---|---|
| `--output` | `-o` | `./<bug_code>.pptx` | Output file path |
| `--browser` | `-b` | `auto` (or saved preference) | Browser to read cookies from |
| `--save-browser` | | | Persist `--browser` to `.env` for future runs |
| `--debug` | | | Save raw HTML to `<bug_code>_debug.html` |

### Browser auto-detection order

`brave` → `chrome` → `edge` → `safari`

Use `--save-browser` to set a permanent default and skip detection entirely.

---

## File structure

```
eBug_to_slide/
├── main.py               # CLI entry point
├── scraper.py            # Cookie extraction + HTTP fetch
├── parser.py             # HTML parsing + image download
├── slide_gen.py          # PPTX generation (python-pptx)
├── requirements.txt      # Python dependencies
├── UX ppt template.pptx  # Slide template (do not rename)
└── .env                  # Auto-created; stores browser preference
```

---

## Troubleshooting

**"Could not extract cookies from any browser"**  
Log into `https://ecl.cyberlink.com` in your browser, then retry. Use `--browser <name>` to target a specific browser.

**"Redirected to login page"**  
Your session has expired. Log in again in the browser.

**"No image-tagged repro steps found"**  
Run with `--debug` to save the raw HTML. Open `<bug_code>_debug.html` and check that repro steps are present and use the `{type:…, step:…, file:…}` format.

**Images missing from slides**  
The tool tries to resolve image URLs from `<img>` tags in the page and from common attachment paths. If images are hosted at an unexpected path, open `<bug_code>_debug.html`, find the `<img>` src values, and let the developer know the URL pattern so `parser.py` can be updated.
