import io
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk


def launch_gui() -> None:
    """Launch the GUI. Called from main.py when no CLI arguments are given."""
    from scraper import (
        clear_ntlm_credentials,
        load_browser_preference,
        load_cookies_string,
        load_last_bug_code,
        load_ntlm_username,
        load_output_dir,
        save_browser_preference,
        save_cookies_string,
        save_last_bug_code,
        save_output_dir,
        set_ntlm_credentials,
    )
    from main import (
        _combined_filename,
        _load_bug_codes_from_file,
        _parse_bug_codes_from_string,
        _resolve_output,
        _run_combined,
        _run_single,
    )

    root = tk.Tk()
    root.title("eBug to Slide")
    root.resizable(False, False)

    pad = {"padx": 8, "pady": 4}

    # ── Bug code(s) ───────────────────────────────────────────────────────────
    tk.Label(root, text="Bug code(s):").grid(row=0, column=0, sticky="e", **pad)
    bug_var = tk.StringVar(value=load_last_bug_code())
    tk.Entry(root, textvariable=bug_var, width=38).grid(
        row=0, column=1, columnspan=2, sticky="w", **pad
    )
    tk.Label(
        root, text="comma- or space-separated for multiple", fg="#888", font=("", 8),
    ).grid(row=1, column=1, columnspan=2, sticky="w", padx=(8, 0), pady=0)

    # ── .txt file ─────────────────────────────────────────────────────────────
    tk.Label(root, text="OR .txt file:").grid(row=2, column=0, sticky="e", **pad)
    file_var = tk.StringVar()
    tk.Label(root, textvariable=file_var, fg="#555", width=30, anchor="w").grid(
        row=2, column=1, sticky="w", **pad
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
        row=2, column=2, sticky="w", **pad
    )

    # ── Output directory ──────────────────────────────────────────────────────
    tk.Label(root, text="Output dir:").grid(row=3, column=0, sticky="e", **pad)
    outdir_var = tk.StringVar(value=load_output_dir() or ".")
    tk.Label(root, textvariable=outdir_var, fg="#555", width=30, anchor="w").grid(
        row=3, column=1, sticky="w", **pad
    )

    def browse_outdir() -> None:
        path = filedialog.askdirectory(title="Select output directory")
        if path:
            outdir_var.set(path)

    tk.Button(root, text="Browse…", command=browse_outdir).grid(
        row=3, column=2, sticky="w", **pad
    )

    # ── Browser ───────────────────────────────────────────────────────────────
    tk.Label(root, text="Browser:").grid(row=4, column=0, sticky="e", **pad)
    browser_choices = [
        "auto", "brave", "chrome", "edge", "safari",
        "firefox", "chromium", "opera", "vivaldi",
    ]
    browser_var = tk.StringVar(value=load_browser_preference())
    ttk.Combobox(
        root, textvariable=browser_var, values=browser_choices,
        state="readonly", width=16,
    ).grid(row=4, column=1, sticky="w", **pad)

    # ── Cookie string (Windows fallback) ──────────────────────────────────────
    tk.Label(root, text="Cookie string:").grid(row=5, column=0, sticky="e", **pad)
    cookie_var = tk.StringVar(value=load_cookies_string())
    tk.Entry(root, textvariable=cookie_var, width=30).grid(
        row=5, column=1, sticky="w", **pad
    )

    def clear_cookies() -> None:
        cookie_var.set("")

    tk.Button(root, text="Clear", command=clear_cookies).grid(
        row=5, column=2, sticky="w", **pad
    )

    # ── Windows NTLM credentials (image auth fallback) ────────────────────────
    tk.Label(root, text="Win username:").grid(row=6, column=0, sticky="e", **pad)
    ntlm_user_var = tk.StringVar(value=load_ntlm_username())
    tk.Entry(root, textvariable=ntlm_user_var, width=30).grid(
        row=6, column=1, sticky="w", **pad
    )

    tk.Label(root, text="Win password:").grid(row=7, column=0, sticky="e", **pad)
    ntlm_pass_var = tk.StringVar()
    tk.Entry(root, textvariable=ntlm_pass_var, width=30, show="*").grid(
        row=7, column=1, sticky="w", **pad
    )

    def clear_creds() -> None:
        ntlm_user_var.set("")
        ntlm_pass_var.set("")
        clear_ntlm_credentials()

    tk.Button(root, text="Clear", command=clear_creds).grid(
        row=6, column=2, sticky="w", **pad
    )

    # ── Combine checkbox ──────────────────────────────────────────────────────
    combine_var = tk.BooleanVar(value=True)
    tk.Checkbutton(
        root, text="Combine multiple bugs into one file", variable=combine_var,
    ).grid(row=8, column=0, columnspan=3, sticky="w", padx=8, pady=(2, 0))

    # ── Generate button ───────────────────────────────────────────────────────
    gen_btn = tk.Button(root, text="Generate Slides", width=20)
    gen_btn.grid(row=9, column=0, columnspan=3, pady=8)

    # ── Log area ──────────────────────────────────────────────────────────────
    log = scrolledtext.ScrolledText(
        root, height=12, width=58, state="disabled", wrap="word"
    )
    log.grid(row=10, column=0, columnspan=3, padx=8, pady=(0, 8))
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
        cookie_str = cookie_var.get().strip()
        combine = combine_var.get()

        if not code and not file_path:
            _log("Error: enter a bug code or select a .txt file.", "err")
            return

        # Pre-populate NTLM credentials before the worker thread starts so
        # fetch_image can use them without calling input() / getpass (which
        # crash in a console=False windowed exe).
        set_ntlm_credentials(ntlm_user_var.get().strip(), ntlm_pass_var.get())

        gen_btn.config(state="disabled")
        _log("─" * 50)

        def run() -> None:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = _Writer()
            sys.stderr = _Writer("err")
            try:
                save_output_dir(out_dir)
                save_browser_preference(browser)
                save_cookies_string(cookie_str)

                codes = (
                    _load_bug_codes_from_file(file_path)
                    if file_path
                    else _parse_bug_codes_from_string(code)
                )
                if not codes:
                    root.after(0, _log, "Error: no bug codes found.", "err")
                    return

                if combine and len(codes) > 1:
                    name = _combined_filename(codes)
                    out = _resolve_output(None, name, out_dir)
                    ok = _run_combined(codes, out, browser, False)
                    if ok:
                        save_last_bug_code(codes[-1])
                else:
                    for c in codes:
                        out = _resolve_output(None, c, out_dir)
                        ok = _run_single(c, out, browser, False)
                        if ok:
                            save_last_bug_code(c)
            except Exception as exc:
                root.after(0, _log, f"Error: {exc}", "err")
            finally:
                sys.stdout.flush()
                sys.stderr.flush()
                sys.stdout = old_out
                sys.stderr = old_err
                root.after(0, lambda: ntlm_pass_var.set(""))
                root.after(0, lambda: gen_btn.config(state="normal"))

        threading.Thread(target=run, daemon=True).start()

    gen_btn.config(command=on_generate)
    root.mainloop()
