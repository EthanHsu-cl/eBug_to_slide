import io
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk


def launch_gui() -> None:
    """Launch the GUI. Called from main.py when no CLI arguments are given."""
    from scraper import (
        load_browser_preference,
        load_last_bug_code,
        load_output_dir,
        save_browser_preference,
        save_last_bug_code,
        save_output_dir,
    )
    from main import _load_bug_codes_from_file, _parse_bug_code, _resolve_output, _run_single

    root = tk.Tk()
    root.title("eBug to Slide")
    root.resizable(False, False)

    pad = {"padx": 8, "pady": 4}

    # ── Bug code ──────────────────────────────────────────────────────────────
    tk.Label(root, text="Bug code:").grid(row=0, column=0, sticky="e", **pad)
    bug_var = tk.StringVar(value=load_last_bug_code())
    tk.Entry(root, textvariable=bug_var, width=38).grid(
        row=0, column=1, columnspan=2, sticky="w", **pad
    )

    # ── .txt file ─────────────────────────────────────────────────────────────
    tk.Label(root, text="OR .txt file:").grid(row=1, column=0, sticky="e", **pad)
    file_var = tk.StringVar()
    tk.Label(root, textvariable=file_var, fg="#555", width=30, anchor="w").grid(
        row=1, column=1, sticky="w", **pad
    )

    def browse_file() -> None:
        path = filedialog.askopenfilename(
            title="Select bug code list",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            file_var.set(path)
            bug_var.set("")

    tk.Button(root, text="Browse…", command=browse_file).grid(
        row=1, column=2, sticky="w", **pad
    )

    # ── Output directory ──────────────────────────────────────────────────────
    tk.Label(root, text="Output dir:").grid(row=2, column=0, sticky="e", **pad)
    outdir_var = tk.StringVar(value=load_output_dir() or ".")
    tk.Label(root, textvariable=outdir_var, fg="#555", width=30, anchor="w").grid(
        row=2, column=1, sticky="w", **pad
    )

    def browse_outdir() -> None:
        path = filedialog.askdirectory(title="Select output directory")
        if path:
            outdir_var.set(path)

    tk.Button(root, text="Browse…", command=browse_outdir).grid(
        row=2, column=2, sticky="w", **pad
    )

    # ── Browser ───────────────────────────────────────────────────────────────
    tk.Label(root, text="Browser:").grid(row=3, column=0, sticky="e", **pad)
    browser_choices = [
        "auto", "brave", "chrome", "edge", "safari",
        "firefox", "chromium", "opera", "vivaldi",
    ]
    browser_var = tk.StringVar(value=load_browser_preference())
    ttk.Combobox(
        root, textvariable=browser_var, values=browser_choices,
        state="readonly", width=16,
    ).grid(row=3, column=1, sticky="w", **pad)

    # ── Generate button ───────────────────────────────────────────────────────
    gen_btn = tk.Button(root, text="Generate Slides", width=20)
    gen_btn.grid(row=4, column=0, columnspan=3, pady=8)

    # ── Log area ──────────────────────────────────────────────────────────────
    log = scrolledtext.ScrolledText(
        root, height=12, width=58, state="disabled", wrap="word"
    )
    log.grid(row=5, column=0, columnspan=3, padx=8, pady=(0, 8))
    log.tag_config("err", foreground="red")

    def _log(text: str, tag: str = "") -> None:
        log.config(state="normal")
        log.insert("end", text + "\n", tag)
        log.see("end")
        log.config(state="disabled")

    class _Writer(io.TextIOBase):
        def __init__(self, tag: str = "") -> None:
            self._tag = tag
            self._buf = ""

        def write(self, s: str) -> int:
            self._buf += s
            while "\n" in self._buf:
                line, self._buf = self._buf.split("\n", 1)
                if line:
                    root.after(0, _log, line, self._tag)
            return len(s)

        def flush(self) -> None:
            if self._buf:
                root.after(0, _log, self._buf, self._tag)
                self._buf = ""

    def on_generate() -> None:
        code = bug_var.get().strip()
        file_path = file_var.get().strip()
        out_dir = outdir_var.get().strip() or "."
        browser = browser_var.get()

        if not code and not file_path:
            _log("Error: enter a bug code or select a .txt file.", "err")
            return

        gen_btn.config(state="disabled")
        _log("─" * 50)

        def run() -> None:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = _Writer()
            sys.stderr = _Writer("err")
            try:
                save_output_dir(out_dir)
                save_browser_preference(browser)
                if file_path:
                    codes = _load_bug_codes_from_file(file_path)
                    for c in codes:
                        out = _resolve_output(None, c, out_dir)
                        ok = _run_single(c, out, browser, False)
                        if ok:
                            save_last_bug_code(c)
                else:
                    bug_code = _parse_bug_code(code)
                    out = _resolve_output(None, bug_code, out_dir)
                    ok = _run_single(bug_code, out, browser, False)
                    if ok:
                        save_last_bug_code(bug_code)
            except Exception as exc:
                root.after(0, _log, f"Error: {exc}", "err")
            finally:
                sys.stdout.flush()
                sys.stderr.flush()
                sys.stdout = old_out
                sys.stderr = old_err
                root.after(0, lambda: gen_btn.config(state="normal"))

        threading.Thread(target=run, daemon=True).start()

    gen_btn.config(command=on_generate)
    root.mainloop()
