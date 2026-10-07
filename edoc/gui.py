"""Native desktop interface; Tk stays on the main thread, Excel on a worker."""

from __future__ import annotations

import json
import os
import queue
import threading
import tkinter as tk
from dataclasses import asdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .config import Mapping, Profile, ValidationError, load_profile, save_profile
from .engine import discover_reports, preview, run_batch
from .excel import ExcelSession


def settings_path() -> Path:
    return Path(os.getenv("APPDATA", str(Path.home()))) / "EDocGenerator" / "last-profile.json"


class MappingDialog(tk.Toplevel):
    def __init__(self, parent, mapping: Mapping, source_sheets: list[str], target_sheets: list[str]):
        super().__init__(parent)
        self.title("Cell mapping")
        self.resizable(False, False)
        self.transient(parent)
        self.result = None
        self.variables = {}
        panel = ttk.Frame(self, padding=20)
        panel.pack(fill="both", expand=True)
        rows = [("Field name", "label"), ("Report worksheet", "source_sheet"),
                ("Read from cell", "source_cell"), ("Approval worksheet", "target_sheet"),
                ("Write into cell", "target_cell"), ("Copy mode", "mode")]
        for index, (label, name) in enumerate(rows):
            ttk.Label(panel, text=label).grid(row=index, column=0, sticky="w", pady=6, padx=(0, 20))
            variable = tk.StringVar(value=getattr(mapping, name))
            self.variables[name] = variable
            if name in {"source_sheet", "target_sheet", "mode"}:
                values = source_sheets if name == "source_sheet" else target_sheets if name == "target_sheet" else ["value", "text"]
                widget = ttk.Combobox(panel, textvariable=variable, values=values, width=30,
                                      state="readonly" if name == "mode" else "normal")
            else:
                widget = ttk.Entry(panel, textvariable=variable, width=33)
            widget.grid(row=index, column=1, sticky="ew")
        self.required = tk.BooleanVar(value=mapping.required)
        self.formula = tk.BooleanVar(value=mapping.allow_formula_overwrite)
        ttk.Checkbutton(panel, text="Required: fail this report if the value is blank", variable=self.required).grid(row=6, columnspan=2, sticky="w", pady=(12, 3))
        ttk.Checkbutton(panel, text="Allow replacing a formula in this destination cell", variable=self.formula).grid(row=7, columnspan=2, sticky="w")
        ttk.Label(panel, text="Value: keep numbers/dates. Text: copy Excel's displayed value,\nincluding units and leading zeros. Use the top-left cell of a merge.",
                  style="Muted.TLabel").grid(row=8, columnspan=2, sticky="w", pady=14)
        actions = ttk.Frame(panel)
        actions.grid(row=9, columnspan=2, sticky="e")
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(actions, text="Save mapping", command=self.save, style="Accent.TButton").pack(side="left")
        self.bind("<Return>", lambda _: self.save())
        self.bind("<Escape>", lambda _: self.destroy())
        self.grab_set()
        self.wait_window()

    def save(self):
        try:
            mapping = Mapping(**{key: variable.get() for key, variable in self.variables.items()},
                              required=self.required.get(), allow_formula_overwrite=self.formula.get())
            mapping.validate()
            self.result = mapping
            self.destroy()
        except ValidationError as exc:
            messagebox.showerror("Check mapping", str(exc), parent=self)


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"eDoc Generator {__version__}")
        width = min(1160, max(960, self.winfo_screenwidth() - 100))
        height = min(850, max(680, self.winfo_screenheight() - 110))
        self.geometry(f"{width}x{height}")
        self.minsize(960, 680)
        self.configure(background="#eef2f7")
        self.profile = Profile()
        self.source_sheets = []
        self.target_sheets = []
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.busy = False
        self.close_after_task = False
        self.last_run = None
        self.vars = {}
        self.configure_style()
        header = ttk.Frame(self, padding=(28, 22, 28, 22), style="Header.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="eDoc Generator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="Calibration approval documents  /  Excel templates  /  PDF certificates", style="HeaderSubtitle.TLabel").pack(anchor="w", pady=(6, 0))
        toolbar = ttk.Frame(self, padding=(28, 14, 28, 14), style="Toolbar.TFrame")
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Load profile", command=self.load).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Save profile", command=self.save).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Help", command=self.help).pack(side="right")
        self.tabs = ttk.Notebook(self)
        self.setup_tab = ttk.Frame(self.tabs, padding=20)
        self.mapping_tab = ttk.Frame(self.tabs, padding=20)
        self.run_tab = ttk.Frame(self.tabs, padding=20)
        self.tabs.add(self.setup_tab, text="01   Files and folders")
        self.tabs.add(self.mapping_tab, text="02   Cell mappings")
        self.tabs.add(self.run_tab, text="03   Generate and results")
        self.build_setup()
        self.build_mappings()
        self.build_run()
        self.status = tk.StringVar(value="Choose your Excel approval template and input folder to begin.")
        ttk.Label(self, textvariable=self.status, style="Status.TLabel", padding=(28, 13)).pack(side="bottom", fill="x")
        self.tabs.pack(fill="both", expand=True, padx=24, pady=(0, 4))
        if settings_path().is_file():
            try:
                self.apply_profile(load_profile(settings_path()))
                self.status.set("Last mapping profile restored. Check folders before generating.")
            except ValidationError:
                self.status.set("Previous settings could not be loaded. Choose files and create your mappings.")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self.poll)

    def configure_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), background="white", foreground="#20344a")
        style.configure("TFrame", background="white")
        style.configure("TLabel", background="white")
        style.configure("Header.TFrame", background="#142a43")
        style.configure("Toolbar.TFrame", background="#eef2f7")
        style.configure("Title.TLabel", font=("Segoe UI", 25, "bold"), background="#142a43", foreground="white")
        style.configure("HeaderSubtitle.TLabel", background="#142a43", foreground="#bdcee0")
        style.configure("Heading.TLabel", font=("Segoe UI", 13, "bold"), foreground="#142a43")
        style.configure("Muted.TLabel", foreground="#617286")
        style.configure("Status.TLabel", foreground="#52657a", background="#eef2f7")
        style.configure("TButton", padding=(14, 9), background="white", bordercolor="#d6dfe9", relief="flat")
        style.map("TButton", background=[("active", "#eaf0f7"), ("disabled", "#edf1f5")], foreground=[("disabled", "#8b97a5")])
        style.configure("Accent.TButton", background="#235daa", foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#194c91"), ("disabled", "#aab8ca")], foreground=[("disabled", "white")])
        style.configure("TEntry", padding=(9, 8), fieldbackground="#fbfcfe", bordercolor="#ccd7e4")
        style.configure("TCombobox", padding=(7, 7), fieldbackground="#fbfcfe", bordercolor="#ccd7e4")
        style.configure("TNotebook", background="#eef2f7", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(17, 11), background="#e5ebf3", font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "white"), ("active", "#f8fafc")], foreground=[("selected", "#235daa")])
        style.configure("Treeview", rowheight=34, background="white", fieldbackground="white", bordercolor="#dce4ed")
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#edf2f8", foreground="#20344a", padding=10)
        style.map("Treeview", background=[("selected", "#e3edfb")], foreground=[("selected", "#16385f")])
        style.configure("Horizontal.TProgressbar", background="#235daa", troughcolor="#e8eef5", borderwidth=0)

    def path_row(self, parent, row: int, label: str, key: str, file=False):
        variable = tk.StringVar()
        self.vars[key] = variable
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=9, padx=(0, 16))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=9)
        ttk.Button(parent, text="Browse…", command=lambda: self.browse(key, file)).grid(row=row, column=2, padx=(10, 0))

    def build_setup(self):
        # A scrolling setup form keeps all controls reachable on laptop displays.
        self.setup_canvas = tk.Canvas(self.setup_tab, background="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.setup_tab, orient="vertical", command=self.setup_canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.setup_canvas.pack(side="left", fill="both", expand=True)
        self.setup_canvas.configure(yscrollcommand=scrollbar.set)
        panel = ttk.Frame(self.setup_canvas, padding=(0, 0, 16, 12))
        form_window = self.setup_canvas.create_window((0, 0), window=panel, anchor="nw")
        panel.bind("<Configure>", lambda _: self.setup_canvas.configure(scrollregion=self.setup_canvas.bbox("all")))
        self.setup_canvas.bind("<Configure>", lambda event: self.setup_canvas.itemconfigure(form_window, width=event.width))
        self.bind_all("<MouseWheel>", self.scroll_setup, add="+")
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text="Choose the files once. Reuse your saved profile next time.", style="Heading.TLabel").grid(row=0, columnspan=3, sticky="w", pady=(0, 12))
        self.path_row(panel, 1, "Approval Excel template", "template", file=True)
        self.path_row(panel, 2, "Input report folder", "input_dir")
        self.path_row(panel, 3, "Output folder", "output_dir")
        ttk.Separator(panel).grid(row=4, columnspan=3, sticky="ew", pady=14)
        ttk.Label(panel, text="Certificate source").grid(row=5, column=0, sticky="w")
        self.vars["certificate_mode"] = tk.StringVar(value="report")
        choices = ttk.Frame(panel)
        choices.grid(row=5, column=1, columnspan=2, sticky="w")
        for label, value in [("Export each Excel report", "report"), ("Matching PDF", "pdf"), ("Approval only", "approval_only")]:
            ttk.Radiobutton(choices, text=label, value=value, variable=self.vars["certificate_mode"]).pack(side="left", padx=(0, 14), pady=4)
        self.path_row(panel, 6, "Certificate PDF folder", "certificate_dir")
        ttk.Label(panel, text="Matching PDF mode: report ABC.xlsx pairs with ABC.pdf.\nWith subfolders enabled, PDF subfolders must match the report subfolders.", style="Muted.TLabel").grid(row=7, column=1, columnspan=2, sticky="w", pady=(0, 12))
        self.vars["approval_sheet"] = tk.StringVar()
        self.vars["report_sheet"] = tk.StringVar()
        for row, label, key in [(8, "Approval PDF worksheet", "approval_sheet"), (9, "Report PDF worksheet", "report_sheet")]:
            ttk.Label(panel, text=label).grid(row=row, column=0, sticky="w", pady=7)
            ttk.Entry(panel, textvariable=self.vars[key]).grid(row=row, column=1, sticky="ew")
        ttk.Label(panel, text="Leave worksheet names blank to export the whole workbook.\nPDF layout follows Excel's saved print areas, margins and page settings.", style="Muted.TLabel").grid(row=10, column=1, columnspan=2, sticky="w", pady=7)
        self.vars["recursive"] = tk.BooleanVar(value=False)
        self.vars["keep_components"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(panel, text="Include reports in input subfolders", variable=self.vars["recursive"]).grid(row=11, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(panel, text="Keep approval and certificate PDFs alongside final.pdf", variable=self.vars["keep_components"]).grid(row=12, columnspan=3, sticky="w")

    def build_mappings(self):
        ttk.Label(self.mapping_tab, text="Map any number of values. No code changes needed.", style="Heading.TLabel").pack(anchor="w")
        actions = ttk.Frame(self.mapping_tab)
        actions.pack(fill="x", pady=12)
        for label, command in [("Inspect worksheets", self.inspect), ("Add mapping", self.add_mapping),
                               ("Edit selected", self.edit_mapping), ("Remove selected", self.remove_mapping),
                               ("Preview sample", self.preview_sample)]:
            ttk.Button(actions, text=label, command=command).pack(side="left", padx=(0, 7))
        columns = ("label", "source", "target", "mode", "required", "formula")
        table = ttk.Frame(self.mapping_tab)
        self.mapping_tree = ttk.Treeview(table, columns=columns, show="headings", selectmode="browse", height=6)
        for key, heading, width in [("label", "Field", 130), ("source", "Report sheet!cell", 210),
                                    ("target", "Approval sheet!cell", 210), ("mode", "Mode", 65),
                                    ("required", "Required", 75), ("formula", "Replace formula", 100)]:
            self.mapping_tree.heading(key, text=heading)
            self.mapping_tree.column(key, width=width, minwidth=50)
        self.mapping_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.mapping_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.mapping_tree.configure(yscrollcommand=scrollbar.set)
        self.mapping_tree.bind("<Double-1>", lambda _: self.edit_mapping())
        ttk.Label(self.mapping_tab, text="Use exact worksheet names and one cell address per mapping.\nText mode preserves displayed units / leading zeros. Value mode copies numbers and dates.\nPreview reads a sample without writing or exporting anything.", style="Muted.TLabel").pack(side="bottom", anchor="w", pady=(14, 0))
        table.pack(fill="both", expand=True)

    def build_run(self):
        actions = ttk.Frame(self.run_tab)
        actions.pack(fill="x", pady=(0, 12))
        ttk.Button(actions, text="Generate documents", command=self.generate, style="Accent.TButton").pack(side="left", padx=(0, 10))
        self.cancel_button = ttk.Button(actions, text="Stop after current report", command=self.cancel, state="disabled")
        self.cancel_button.pack(side="left")
        self.output_button = ttk.Button(actions, text="Open output folder", command=self.open_output)
        self.output_button.pack(side="right")
        self.progress = ttk.Progressbar(self.run_tab, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 12))
        table = ttk.Frame(self.run_tab)
        self.run_tree = ttk.Treeview(table, columns=("source", "status", "details"), show="headings", height=6)
        for key, title, width in [("source", "Report", 200), ("status", "Status", 90), ("details", "Output or error", 580)]:
            self.run_tree.heading(key, text=title)
            self.run_tree.column(key, width=width)
        self.run_tree.tag_configure("failed", foreground="#ab2431")
        self.run_tree.tag_configure("success", foreground="#167047")
        self.run_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(table, orient="vertical", command=self.run_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.run_tree.configure(yscrollcommand=scrollbar.set)
        self.run_tree.bind("<Double-1>", self.show_result)
        ttk.Label(self.run_tab, text="Every batch gets a new output folder. Successful reports contain a filled approval workbook,\nfinal.pdf and audit.json. batch.json lists successes, failures and unprocessed reports.\nDouble-click a result for its full path or error.", style="Muted.TLabel").pack(side="bottom", anchor="w", pady=(14, 0))
        table.pack(fill="both", expand=True)

    def scroll_setup(self, event):
        if self.tabs.select() == str(self.setup_tab):
            self.setup_canvas.yview_scroll(int(-event.delta / 120), "units")

    def browse(self, key, file):
        if self.busy:
            return
        path = filedialog.askopenfilename(title="Select approval Excel template", filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls *.xlsb")]) if file else filedialog.askdirectory(title=f"Choose {key.replace('_', ' ')}", mustexist=key != "output_dir")
        if path:
            self.vars[key].set(path)
            if key in {"template", "input_dir"}:
                self.source_sheets = []
                self.target_sheets = []

    def current_profile(self) -> Profile:
        return Profile(**{key: var.get() for key, var in self.vars.items()},
                       mappings=[Mapping(**asdict(mapping)) for mapping in self.profile.mappings])

    def apply_profile(self, profile):
        self.profile = profile
        for key, variable in self.vars.items():
            variable.set(getattr(profile, key))
        self.source_sheets = []
        self.target_sheets = []
        self.refresh_mappings()

    def refresh_mappings(self):
        self.mapping_tree.delete(*self.mapping_tree.get_children())
        for index, mapping in enumerate(self.profile.mappings):
            self.mapping_tree.insert("", "end", iid=str(index), values=(mapping.label,
                f"{mapping.source_sheet}!{mapping.source_cell}", f"{mapping.target_sheet}!{mapping.target_cell}",
                mapping.mode, "Yes" if mapping.required else "No", "Yes" if mapping.allow_formula_overwrite else "No"))

    def selected_mapping(self):
        selected = self.mapping_tree.selection()
        return int(selected[0]) if selected else None

    def add_mapping(self):
        if self.busy:
            return
        mapping = Mapping(source_sheet=self.source_sheets[0] if self.source_sheets else "",
                          target_sheet=self.target_sheets[0] if self.target_sheets else "")
        self.mapping_dialog(mapping)

    def edit_mapping(self):
        if self.busy:
            return
        index = self.selected_mapping()
        if index is not None:
            self.mapping_dialog(self.profile.mappings[index], index)

    def mapping_dialog(self, mapping, index=None):
        dialog = MappingDialog(self, mapping, self.source_sheets, self.target_sheets)
        if dialog.result is None:
            return
        proposed = list(self.profile.mappings)
        if index is None:
            proposed.append(dialog.result)
        else:
            proposed[index] = dialog.result
        try:
            Profile(mappings=proposed).validate_mappings()
        except ValidationError as exc:
            messagebox.showerror("Check mappings", str(exc), parent=self)
            return
        self.profile.mappings = proposed
        self.refresh_mappings()

    def remove_mapping(self):
        if self.busy:
            return
        index = self.selected_mapping()
        if index is not None:
            del self.profile.mappings[index]
            self.refresh_mappings()

    def sample_file(self):
        return filedialog.askopenfilename(title="Choose one sample calibration report", initialdir=self.vars["input_dir"].get() or None,
                                         filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls *.xlsb")])

    def inspect(self):
        if self.busy:
            return
        template = Path(self.vars["template"].get())
        if not template.is_file():
            messagebox.showerror("Select template", "Choose the approval template first.", parent=self)
            return
        sample = self.sample_file()
        if not sample:
            return
        def work():
            with ExcelSession() as session:
                with session.open(Path(sample)) as book:
                    source_sheets = session.sheets(book)
                with session.open(template) as book:
                    target_sheets = session.sheets(book)
            return {"source_sheets": source_sheets, "target_sheets": target_sheets}
        self.start_task("inspect", work)

    def preview_sample(self):
        if self.busy:
            return
        try:
            profile = self.current_profile()
            profile.validate_mappings()
            if not Path(profile.template).is_file():
                raise ValidationError("Choose an approval template first.")
            sample = self.sample_file()
            if sample:
                self.start_task("preview", lambda: preview(profile, Path(sample)))
        except ValidationError as exc:
            messagebox.showerror("Check setup", str(exc), parent=self)

    def save(self):
        if self.busy:
            return
        try:
            profile = self.current_profile()
            profile.validate_mappings()
            path = filedialog.asksaveasfilename(title="Save reusable profile", defaultextension=".json", filetypes=[("Mapping profile", "*.json")])
            if path:
                save_profile(Path(path), profile)
                self.status.set(f"Profile saved: {path}")
        except (ValidationError, OSError) as exc:
            messagebox.showerror("Could not save profile", str(exc), parent=self)

    def load(self):
        if self.busy:
            return
        path = filedialog.askopenfilename(title="Load reusable profile", filetypes=[("Mapping profile", "*.json")])
        if path:
            try:
                self.apply_profile(load_profile(Path(path)))
                self.status.set("Profile loaded. Check paths, then preview a sample.")
            except ValidationError as exc:
                messagebox.showerror("Could not load profile", str(exc), parent=self)

    def generate(self):
        if self.busy:
            return
        try:
            profile = self.current_profile()
            profile.validate()
            files = discover_reports(profile)
            if not files:
                raise ValidationError("No Excel reports found in the selected folder.")
            save_profile(settings_path(), profile)
            self.run_tree.delete(*self.run_tree.get_children())
            self.progress.configure(value=0, maximum=len(files))
            self.tabs.select(self.run_tab)
            self.start_task("batch", lambda: run_batch(profile, progress=self.events.put, cancel=self.cancel_event))
        except (ValidationError, OSError) as exc:
            messagebox.showerror("Check setup", str(exc), parent=self)

    def start_task(self, name, work):
        self.cancel_event.clear()
        self.busy = True
        self.set_controls_busy(True)
        self.cancel_button.configure(state="normal" if name == "batch" else "disabled")
        self.status.set("Reading worksheets…" if name == "inspect" else "Checking sample…" if name == "preview" else "Starting Excel and processing reports…")
        def worker():
            try:
                value = work()
                self.events.put({"kind": "complete", "task": name, "value": value})
            except Exception as exc:
                self.events.put({"kind": "error", "message": str(exc) or type(exc).__name__})
        threading.Thread(target=worker, name="edoc-excel-worker", daemon=False).start()

    def set_controls_busy(self, busy):
        def visit(widget):
            for child in widget.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Entry, ttk.Checkbutton, ttk.Radiobutton, ttk.Combobox)):
                    child.state(["disabled"] if busy else ["!disabled"])
                visit(child)
        visit(self)
        self.cancel_button.configure(state="disabled")

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind = event["kind"]
                if kind == "starting":
                    self.status.set(f"Processing {event['index']}/{event['total']}: {Path(event['source']).name}")
                elif kind == "result":
                    details = event["output"] if event["status"] == "success" else event["error"]
                    self.run_tree.insert("", "end", values=(Path(event["source"]).name, event["status"], details), tags=(event["status"],))
                    self.progress.configure(value=event["index"])
                elif kind in {"complete", "error"}:
                    self.busy = False
                    self.set_controls_busy(False)
                    if self.close_after_task:
                        self.destroy()
                        return
                    if kind == "error":
                        self.status.set("Task stopped. See the error details.")
                        messagebox.showerror("Could not complete task", event["message"], parent=self)
                    else:
                        self.task_complete(event["task"], event["value"])
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def task_complete(self, task, value):
        if task == "inspect":
            self.source_sheets = value["source_sheets"]
            self.target_sheets = value["target_sheets"]
            self.status.set("Worksheets loaded. Add mappings using the worksheet dropdowns.")
            messagebox.showinfo("Worksheets", "Report: " + ", ".join(self.source_sheets) + "\n\nApproval: " + ", ".join(self.target_sheets), parent=self)
        elif task == "preview":
            self.source_sheets = value["source_sheets"]
            self.target_sheets = value["target_sheets"]
            window = tk.Toplevel(self)
            window.title("Sample extraction preview — no files changed")
            window.geometry("920x440")
            ttk.Label(window, text=value["report"], padding=14).pack(anchor="w")
            tree = ttk.Treeview(window, columns=("label", "source", "value", "target"), show="headings")
            for key, title in [("label", "Field"), ("source", "Read from"), ("value", "Extracted value"), ("target", "Write to")]:
                tree.heading(key, text=title)
                tree.column(key, width=210)
            for item in value["values"]:
                tree.insert("", "end", values=tuple(item[key] for key in ("label", "source", "value", "target")))
            tree.pack(fill="both", expand=True, padx=14, pady=(0, 14))
            self.status.set("Preview complete. Check extracted values before generating.")
        else:
            self.last_run = value.run_dir
            failed = sum(result.status == "failed" for result in value.results)
            summary = f"{value.succeeded} succeeded, {failed} failed." + (" Batch stopped early." if value.cancelled else "")
            self.status.set(summary + " Output: " + value.run_dir)
            messagebox.showinfo("Batch finished", summary + "\n\n" + value.run_dir, parent=self)

    def cancel(self):
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status.set("Stop requested. Finishing the current report and closing Excel…")

    def open_output(self):
        target = Path(self.last_run or self.vars["output_dir"].get())
        if target.is_dir() and hasattr(os, "startfile"):
            os.startfile(str(target.resolve()))

    def show_result(self, _event):
        selected = self.run_tree.selection()
        if selected:
            values = self.run_tree.item(selected[0], "values")
            messagebox.showinfo(values[0], values[2], parent=self)

    def help(self):
        messagebox.showinfo("Quick start", "1. Select the approval Excel template, input folder and output folder.\n"
            "2. Select report export or same-name certificate PDFs.\n3. Inspect worksheets, then add your cell mappings.\n"
            "4. Preview a sample and verify the values. Save your profile.\n5. Generate documents.\n\n"
            "Requires Windows and desktop Microsoft Excel. Original files are never saved.\n"
            "Set Excel print areas in advance. Macros and external links are not supported.\n"
            "This app prepares documents; it does not approve or digitally sign them.\n"
            "Merging a digitally signed PDF does not preserve its signature validity. Keep the original certificate.", parent=self)

    def on_close(self):
        if self.busy:
            if messagebox.askyesno("Work in progress", "Stop after the current task/report, then close the app?", parent=self):
                self.close_after_task = True
                self.cancel()
        else:
            self.destroy()


def main():
    Application().mainloop()
