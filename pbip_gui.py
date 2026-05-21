#!/usr/bin/env python3
"""
pbip_gui.py — Tkinter GUI wrapper for pbip_extract.py

Run:
    python pbip_gui.py
"""

import contextlib
import io
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Allow running from any working directory
sys.path.insert(0, str(Path(__file__).parent))
import pbip_extract as pe


def _auto_filename(pbip_path: str, copilot: bool) -> str:
    project_name = Path(pbip_path).name.replace(".pbip", "").replace("_", " ").title()
    suffix = "copilot_kb.txt" if copilot else "docs.md"
    return f"{project_name.replace(' ', '_')}_{suffix}"


def _run_extraction(
    pbip_path: str,
    output_dir: str,
    copilot: bool,
    log: scrolledtext.ScrolledText,
    run_btn: ttk.Button,
    root: tk.Tk,
):
    def append(text: str):
        def _do():
            log.configure(state="normal")
            log.insert(tk.END, text)
            log.see(tk.END)
            log.configure(state="disabled")
        root.after(0, _do)

    class LogWriter(io.TextIOBase):
        def write(self, s):
            if s:
                append(s)
            return len(s)

    writer = LogWriter()
    try:
        pbip_root = Path(pbip_path)
        project_name = pbip_root.name.replace(".pbip", "").replace("_", " ").title()

        with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
            model_path, model_format = pe.find_semantic_model(pbip_root)
            if not model_path:
                append("ERROR: No semantic model found (.bim or TMDL definition folder).\n")
                return

            print(f"Found semantic model ({model_format}): {model_path}")

            parser = pe.TMSLParser(model_path) if model_format == "TMSL" else pe.TMLDParser(model_path)

            report_data = None
            report_def_dir = pe.find_report_definition_dir(pbip_root)
            if report_def_dir:
                print(f"Found modern report definition: {report_def_dir}")
                try:
                    report_data = pe.parse_report_definition(report_def_dir)
                except Exception as e:
                    print(f"WARNING: Could not parse modern report definition: {e}")

            if not report_data or not report_data.get("pages"):
                report_json_path = pe.find_report_json(pbip_root)
                if report_json_path:
                    print(f"Found legacy report.json: {report_json_path}")
                    try:
                        report_data = pe.parse_report(report_json_path)
                    except Exception as e:
                        print(f"WARNING: Could not parse report.json: {e}")
                else:
                    print("No report found — skipping report structure section.")

            if report_data and report_data.get("pages"):
                total_visuals = sum(len(p.get("visuals", [])) for p in report_data["pages"])
                total_fields = sum(
                    len(v.get("fields", []))
                    for p in report_data["pages"]
                    for v in p.get("visuals", [])
                )
                print(f"  Pages: {len(report_data['pages'])}, Visuals: {total_visuals}, Field bindings: {total_fields}")

            tables = parser.tables()
            n_measures = sum(len(parser.get_measures(t)) for t in tables)

            filename = _auto_filename(pbip_path, copilot)
            output_path = Path(output_dir) / filename

            if copilot:
                content = pe.render_copilot_kb(project_name, parser, report_data, model_format)
            else:
                content = pe.render_markdown(project_name, parser, report_data, model_format)

            output_path.write_text(content, encoding="utf-8")

            print(f"\nDone. Written to: {output_path.resolve()}")
            print(f"  Tables:        {len(tables)}")
            print(f"  Measures:      {n_measures}")
            print(f"  Relationships: {len(parser.relationships())}")

        append("\n--- Done ---\n")
    except Exception as exc:
        append(f"\nERROR: {exc}\n")
    finally:
        root.after(0, lambda: run_btn.configure(state="normal"))


def main():
    root = tk.Tk()
    root.title("PBIP Extract")
    root.resizable(True, True)

    frame = ttk.Frame(root, padding=12)
    frame.grid(sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)

    # --- PBIP folder ---
    ttk.Label(frame, text="PBIP folder:").grid(row=0, column=0, sticky="w", pady=4)
    pbip_var = tk.StringVar()
    ttk.Entry(frame, textvariable=pbip_var, width=60).grid(row=0, column=1, sticky="ew", padx=6)

    def browse_pbip():
        path = filedialog.askdirectory(title="Select .pbip project folder")
        if path:
            pbip_var.set(path)
            if not output_dir_var.get():
                output_dir_var.set(path)

    ttk.Button(frame, text="Browse…", command=browse_pbip).grid(row=0, column=2, padx=4)

    # --- Output folder ---
    ttk.Label(frame, text="Output folder:").grid(row=1, column=0, sticky="w", pady=4)
    output_dir_var = tk.StringVar()
    ttk.Entry(frame, textvariable=output_dir_var, width=60).grid(row=1, column=1, sticky="ew", padx=6)

    def browse_output_dir():
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            output_dir_var.set(path)

    ttk.Button(frame, text="Browse…", command=browse_output_dir).grid(row=1, column=2, padx=4)

    # --- Mode ---
    copilot_var = tk.BooleanVar(value=False)
    mode_frame = ttk.Frame(frame)
    mode_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=6)
    ttk.Label(mode_frame, text="Output mode:").pack(side="left")
    ttk.Radiobutton(mode_frame, text="Markdown docs", variable=copilot_var, value=False).pack(side="left", padx=8)
    ttk.Radiobutton(mode_frame, text="Copilot knowledge base", variable=copilot_var, value=True).pack(side="left")

    # --- Filename preview ---
    preview_var = tk.StringVar(value="Output file: —")
    ttk.Label(frame, textvariable=preview_var, foreground="gray").grid(row=3, column=0, columnspan=3, sticky="w")

    def _update_preview(*_):
        pbip = pbip_var.get().strip()
        out_dir = output_dir_var.get().strip()
        if pbip and out_dir:
            filename = _auto_filename(pbip, copilot_var.get())
            preview_var.set(f"Output file: {Path(out_dir) / filename}")
        elif pbip:
            preview_var.set(f"Output file: {_auto_filename(pbip, copilot_var.get())}")
        else:
            preview_var.set("Output file: —")

    pbip_var.trace_add("write", _update_preview)
    output_dir_var.trace_add("write", _update_preview)
    copilot_var.trace_add("write", _update_preview)

    # --- Run button ---
    run_btn = ttk.Button(frame, text="Generate")
    run_btn.grid(row=4, column=0, columnspan=3, pady=8)

    def on_run():
        pbip = pbip_var.get().strip()
        if not pbip:
            messagebox.showwarning("Missing input", "Please select a PBIP folder first.")
            return
        out_dir = output_dir_var.get().strip()
        if not out_dir:
            messagebox.showwarning("Missing output", "Please select an output folder first.")
            return
        run_btn.configure(state="disabled")
        log.configure(state="normal")
        log.delete("1.0", tk.END)
        log.configure(state="disabled")
        threading.Thread(
            target=_run_extraction,
            args=(pbip, out_dir, copilot_var.get(), log, run_btn, root),
            daemon=True,
        ).start()

    run_btn.configure(command=on_run)

    # --- Log area ---
    log = scrolledtext.ScrolledText(frame, state="disabled", height=16, wrap="word", font=("Courier", 10))
    log.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=4)
    frame.rowconfigure(5, weight=1)

    root.mainloop()


if __name__ == "__main__":
    main()
