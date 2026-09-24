"""Tkinter front end for merge.py.

Pick one or more .rar (or .zip) files, choose the output .zip, press Run.
All site folders from all selected archives are merged together.

Usage: uv run --with openpyxl python merge_gui.py
"""
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from merge import MergeError, merge


class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NetAct CR Merge Tool")
        self.geometry("820x560")
        self.minsize(600, 420)
        self.inputs: list[Path] = []
        self.log_queue: queue.Queue = queue.Queue()
        self.output_auto = True  # output path follows the first input until edited

        pad = {"padx": 8, "pady": 4}

        # --- Input files
        box_in = ttk.LabelFrame(self, text="Input (.rar files)")
        box_in.pack(fill="both", expand=False, **pad)
        self.listbox = tk.Listbox(box_in, height=6, selectmode="extended")
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        sb = ttk.Scrollbar(box_in, orient="vertical", command=self.listbox.yview)
        sb.pack(side="left", fill="y", pady=8)
        self.listbox.config(yscrollcommand=sb.set)
        btns = ttk.Frame(box_in)
        btns.pack(side="left", fill="y", padx=8, pady=8)
        ttk.Button(btns, text="Add files…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="Remove", command=self.remove_selected).pack(fill="x", pady=2)
        ttk.Button(btns, text="Clear", command=self.clear_files).pack(fill="x", pady=2)

        # --- Output file
        box_out = ttk.LabelFrame(self, text="Output (.zip)")
        box_out.pack(fill="x", **pad)
        self.output_var = tk.StringVar()
        entry = ttk.Entry(box_out, textvariable=self.output_var)
        entry.pack(side="left", fill="x", expand=True, padx=8, pady=8)
        entry.bind("<Key>", lambda e: setattr(self, "output_auto", False))
        ttk.Button(box_out, text="Browse…", command=self.choose_output).pack(side="left", padx=8)

        # --- Run
        bar = ttk.Frame(self)
        bar.pack(fill="x", **pad)
        self.run_btn = ttk.Button(bar, text="Run", command=self.run_merge)
        self.run_btn.pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=200)
        self.progress.pack(side="left", padx=12)
        self.status = ttk.Label(bar, text="Select input files.")
        self.status.pack(side="left")

        # --- Log
        box_log = ttk.LabelFrame(self, text="Log")
        box_log.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(box_log, height=12, wrap="none", state="disabled")
        self.log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        lsb = ttk.Scrollbar(box_log, orient="vertical", command=self.log.yview)
        lsb.pack(side="left", fill="y", pady=8)
        self.log.config(yscrollcommand=lsb.set)

        self.after(100, self.drain_log)

    # ----- input list
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .rar files",
            filetypes=[("RAR archives", "*.rar"), ("RAR / ZIP archives", "*.rar *.zip"),
                       ("All files", "*")],
        )
        for p in map(Path, paths):
            if p not in self.inputs:
                self.inputs.append(p)
                self.listbox.insert("end", str(p))
        self.update_default_output()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.inputs[i]
        self.update_default_output()

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.inputs.clear()
        self.update_default_output()

    # ----- output
    def update_default_output(self):
        if not self.output_auto:
            return
        if self.inputs:
            first = self.inputs[0]
            self.output_var.set(str(first.with_name(first.stem + "_merged.zip")))
        else:
            self.output_var.set("")

    def choose_output(self):
        current = Path(self.output_var.get()) if self.output_var.get() else None
        path = filedialog.asksaveasfilename(
            title="Save merged output as",
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip"), ("RAR archive", "*.rar")],
            initialdir=str(current.parent) if current else None,
            initialfile=current.name if current else "merged.zip",
        )
        if path:
            self.output_var.set(path)
            self.output_auto = False

    # ----- run
    def run_merge(self):
        if not self.inputs:
            messagebox.showwarning("No input", "Add at least one .rar file.")
            return
        out = self.output_var.get().strip()
        if not out:
            messagebox.showwarning("No output", "Choose an output file.")
            return
        out = Path(out)
        if out.suffix.lower() not in (".zip", ".rar"):
            out = out.with_suffix(".zip")
            self.output_var.set(str(out))
        if out.exists() and not messagebox.askyesno("Overwrite?", f"{out.name} exists. Overwrite it?"):
            return

        self.clear_log()
        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self.status.config(text="Merging…")
        threading.Thread(target=self.worker, args=(list(self.inputs), out), daemon=True).start()

    def worker(self, inputs, out):
        try:
            merge(inputs, out, log=self.log_queue.put)
            self.log_queue.put(("done", out))
        except MergeError as e:
            self.log_queue.put(("error", str(e)))
        except Exception as e:  # show unexpected failures instead of dying silently
            self.log_queue.put(("error", f"{type(e).__name__}: {e}"))

    def finish(self, kind, value):
        self.progress.stop()
        self.run_btn.config(state="normal")
        if kind == "done":
            self.status.config(text=f"Done: {value.name}")
            messagebox.showinfo("Finished", f"Merged file saved to:\n{value}")
        else:
            self.status.config(text="Failed")
            self.write_log(f"\nERROR: {value}")
            messagebox.showerror("Error", value)

    # ----- log
    def drain_log(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            if isinstance(msg, tuple):
                self.finish(*msg)
            else:
                self.write_log(msg)
        self.after(100, self.drain_log)

    def write_log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    MergeApp().mainloop()
