#!/usr/bin/env python3
"""
pbip_gui.py - small Tkinter window around pbip_extract.py (standard library only).

Run:  python pbip_gui.py     (on Windows: double-click, or `py pbip_gui.py`)
"""

import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pbip_extract as pe  # noqa: E402


def main():
    root = tk.Tk()
    root.title("PBIP Extract")
    frame = ttk.Frame(root, padding=12)
    frame.grid(sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)

    pbip_var, out_var, counts_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
    query_var = tk.BooleanVar(value=True)
    copilot_var = tk.BooleanVar(value=True)

    def pick_pbip():
        path = filedialog.askopenfilename(title="Select .pbip file", filetypes=[("Power BI project", "*.pbip")])
        if path:
            pbip_var.set(path)
            if not out_var.get():
                out_var.set(str(Path(path).parent))

    def pick_folder():
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            out_var.set(path)

    def pick_counts():
        path = filedialog.askopenfilename(title="Select the pasted row-count result")
        if path:
            counts_var.set(path)

    rows = ((".pbip file:", pbip_var, pick_pbip), ("Output folder:", out_var, pick_folder),
            ("Row counts file (optional):", counts_var, pick_counts))
    for row, (label, var, command) in enumerate(rows):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=var, width=70).grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(frame, text="Browse...", command=command).grid(row=row, column=2)

    ttk.Checkbutton(frame, variable=copilot_var,
                    text="Name the output <project>_copilot_kb.txt (knowledge file for Copilot / Claude / ChatGPT); "
                         "off = <project>_docs.md"
                    ).grid(row=3, column=0, columnspan=3, sticky="w", pady=4)
    ttk.Checkbutton(frame, variable=query_var,
                    text="Also write the *_rowcount_query.dax helper (run it in Power BI Desktop to get row counts)"
                    ).grid(row=4, column=0, columnspan=3, sticky="w", pady=4)
    run_button = ttk.Button(frame, text="Generate documentation")
    run_button.grid(row=5, column=0, columnspan=3, pady=8)
    log = scrolledtext.ScrolledText(frame, state="disabled", height=14, wrap="word")
    log.grid(row=6, column=0, columnspan=3, sticky="nsew")
    frame.rowconfigure(6, weight=1)

    def write(message):                       # called from the worker thread
        def append():
            log.configure(state="normal")
            log.insert(tk.END, str(message) + "\n")
            log.see(tk.END)
            log.configure(state="disabled")
        root.after(0, append)

    def work(pbip, out_dir, counts, with_query, copilot):
        try:
            pe.run(pbip, out_dir, counts or None, with_query, copilot, log=write)
            write("Done.")
        except Exception as e:                # show every problem in the window, never a silent crash
            write("ERROR: %s" % e)
        finally:
            root.after(0, lambda: run_button.configure(state="normal"))

    def on_run():
        pbip, out_dir = pbip_var.get().strip(), out_var.get().strip()
        if not pbip or not out_dir:
            messagebox.showwarning("Missing input", "Select a .pbip file and an output folder first.")
            return
        if not Path(out_dir).is_dir():
            messagebox.showwarning("Missing folder", "The output folder does not exist.")
            return
        run_button.configure(state="disabled")
        log.configure(state="normal")
        log.delete("1.0", tk.END)
        log.configure(state="disabled")
        threading.Thread(target=work, args=(pbip, out_dir, counts_var.get().strip(), query_var.get(), copilot_var.get()),
                         daemon=True).start()

    run_button.configure(command=on_run)
    root.mainloop()


if __name__ == "__main__":
    main()
