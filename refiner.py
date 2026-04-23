"""
refiner.py — Optional AI text refinement via a local Ollama model.

Call refine_bug_data() to improve title and section_texts in-place.
All functions degrade gracefully: if Ollama is unreachable or the model
errors out, the original text is preserved and a warning is printed to stderr.
"""

import json
import sys

import requests

from parser import BugData

OLLAMA_BASE = "http://localhost:11434"
_CONNECT_TIMEOUT = 5    # seconds — fail fast if Ollama isn't running
_READ_TIMEOUT    = 120  # seconds — allow time for slower local models

_SYSTEM_PROMPT = (
    "You are a seasoned QA Engineer writing bug reports for cross-functional RD and QA teams. "
    "Your job is to make existing bug report text clear and professional so any manager can "
    "understand it at a glance — without any prior technical context.\n\n"
    "Rules:\n"
    "- Fix all grammatical and spelling errors.\n"
    "- Use simple, direct language. Avoid complex vocabulary or long sentences.\n"
    "- Preserve every technical fact: product names, version numbers, feature names, "
    "file names, UI element names, and step-by-step details must not be changed.\n"
    "- Do NOT add explanations, bullet points, or headers.\n"
    "- Return ONLY the refined text with no introduction, no commentary, no quotation marks."
)


def check_ollama_available(model: str) -> bool:
    """Return True if Ollama is reachable and the requested model is present."""
    try:
        resp = requests.get(f"{OLLAMA_BASE}/api/tags", timeout=_CONNECT_TIMEOUT)
        resp.raise_for_status()
    except requests.ConnectionError:
        print(
            f"WARNING: Ollama is not running at {OLLAMA_BASE}. Skipping AI refinement.",
            file=sys.stderr,
        )
        return False
    except requests.Timeout:
        print(
            f"WARNING: Ollama did not respond at {OLLAMA_BASE} (timeout). Skipping AI refinement.",
            file=sys.stderr,
        )
        return False
    except requests.RequestException as exc:
        print(f"WARNING: Could not reach Ollama: {exc}. Skipping AI refinement.", file=sys.stderr)
        return False

    data = resp.json()
    models = data.get("models", [])
    base_name = model.split(":")[0]
    if not any(
        m.get("name", "").startswith(base_name) or m.get("name", "") == model
        for m in models
    ):
        available = [m.get("name", "") for m in models]
        print(
            f"WARNING: Model '{model}' not found in Ollama. "
            f"Available: {available}. Skipping AI refinement.",
            file=sys.stderr,
        )
        return False

    return True


def refine_text(text: str, model: str) -> str:
    """Send text to Ollama for refinement. Returns the refined string."""
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": text},
        ],
        "stream": True,
        "options": {"temperature": 0.3},
    }
    resp = requests.post(
        f"{OLLAMA_BASE}/api/chat",
        json=payload,
        timeout=(_CONNECT_TIMEOUT, _READ_TIMEOUT),
        stream=True,
    )
    resp.raise_for_status()

    parts: list[str] = []
    for line in resp.iter_lines():
        if not line:
            continue
        chunk = json.loads(line)
        content = chunk.get("message", {}).get("content", "")
        if content:
            parts.append(content)
        if chunk.get("done"):
            break

    result = "".join(parts).strip()
    if not result:
        raise ValueError("Ollama returned an empty response")
    return result


def _strip_module_prefix(title: str) -> str:
    """Return only the description part after 'Module: ', or the full title if not found."""
    _, sep, rest = title.partition(": ")
    return rest.strip() if sep else title


def refine_bug_data(bug_data: BugData, model: str) -> None:
    """Refine bug_data.title and section_texts values in-place using Ollama.

    The module prefix (e.g. 'VideoStudio: ') is stripped from the title before
    refinement so only the short description appears on the slide.
    """
    if not check_ollama_available(model):
        # Still strip the module prefix even when AI is unavailable.
        bug_data.title = _strip_module_prefix(bug_data.title)
        return

    print(f"Refining text with Ollama model '{model}' ...")

    try:
        bug_data.title = refine_text(_strip_module_prefix(bug_data.title), model)
    except Exception as exc:
        print(f"WARNING: AI refinement failed for title: {exc}", file=sys.stderr)
        bug_data.title = _strip_module_prefix(bug_data.title)

    for key in ("Current", "Reference", "Proposal"):
        if not bug_data.section_texts.get(key):
            continue
        try:
            bug_data.section_texts[key] = refine_text(bug_data.section_texts[key], model)
        except Exception as exc:
            print(f"WARNING: AI refinement failed for '{key}' section: {exc}", file=sys.stderr)
