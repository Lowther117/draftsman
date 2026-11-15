"""The Draftsman window: templates on the left, the form a template asks for
on the right, Generate at the bottom."""
from __future__ import annotations

import datetime as dt
import os
import queue
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import APP_NAME, __version__, batch, merge, paths, pdf, starter, theme
from .library import Library, file_name, log_issue, match_keys, read_rows, save_values, unique_path

VALUE_FILES = [("Spreadsheets and value files", "*.csv *.xlsx *.xlsm *.json *.txt"), ("All files", "*.*")]


class App:
    def __init__(self, root, settings=None, quiet=False):
        self.root = root
        self.settings = settings if settings is not None else paths.load_settings()
        self.quiet = quiet                 # the self-test: no dialogs, nothing opened
        self.busy = False
        self.tasks = queue.Queue()
        self.template = None               # path of the template on the form
        self.scan = None
        self.paths = {}                    # tree row id -> template path (those listed)
        self.all_paths = set()             # every template in the library, listed or filtered out
        self.memory = {}                   # field name -> last value typed, carried between templates
        self.ticks = {}                    # (group, clause) -> ticked
        self.field_vars, self.field_texts, self.switch_vars, self.clause_vars, self.control_vars = {}, {}, {}, {}, {}
        self.mode = theme.from_settings(self.settings)
        root.title(APP_NAME)
        root.geometry(self.settings.get("geometry") or "1180x760")
        root.minsize(940, 600)
        self.c = theme.apply(root, self.mode)
        lib = self.settings.get("library")
        self.library = Library(lib if isinstance(lib, str) and lib else os.path.join(paths.documents_dir(), "Draftsman Library"))
        self._build()
        root.bind("<Control-d>", self.cycle_theme)
        root.bind("<Control-Return>", lambda _e: self.generate() and None)
        root.protocol("WM_DELETE_WINDOW", self.quit)
        self._poll()
        self.open_library(first=True)

    # ------------------------------------------------------------ layout
    def _build(self):
        root = self.root
        top = ttk.Frame(root, padding=(12, 10, 12, 6))
        top.pack(fill="x")
        ttk.Label(top, text="Library", style="Field.TLabel").pack(side="left")
        self.lib_var = tk.StringVar()
        ttk.Label(top, textvariable=self.lib_var, style="Dim.TLabel").pack(side="left", padx=8)
        ttk.Button(top, text="Standing values...", command=self.edit_standing).pack(side="right")
        ttk.Button(top, text="Add starter templates", command=self.add_starters).pack(side="right", padx=(0, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_templates).pack(side="right", padx=6)
        ttk.Button(top, text="Open folder", command=lambda: paths.open_in_system(self.library.root)).pack(side="right")
        ttk.Button(top, text="Change...", command=self.change_library).pack(side="right", padx=6)

        self.status_var = tk.StringVar(value="Pick a template on the left.")
        self.control_table = tk.BooleanVar(value=False)
        ttk.Label(root, textvariable=self.status_var, style="Dim.TLabel", padding=(12, 4)).pack(side="bottom", fill="x")

        bottom = ttk.Frame(root, padding=(12, 6, 12, 2))
        bottom.pack(side="bottom", fill="x")
        bottom.columnconfigure(1, weight=1)
        ttk.Label(bottom, text="Save in", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        out = self.settings.get("out_dir")
        self.out_var = tk.StringVar(value=out if isinstance(out, str) and out else os.path.join(paths.documents_dir(), "Draftsman Output"))
        ttk.Entry(bottom, textvariable=self.out_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(bottom, text="Browse...", command=self.pick_out).grid(row=0, column=2)
        ttk.Label(bottom, text="File name", style="Field.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        pattern = self.settings.get("pattern")
        self.pattern_var = tk.StringVar(value=pattern if isinstance(pattern, str) and pattern else "{{TEMPLATE}} - {{Reference}}")
        ttk.Entry(bottom, textvariable=self.pattern_var).grid(row=1, column=1, sticky="ew", padx=6, pady=(6, 0))
        ttk.Label(bottom, text="any field, {{TEMPLATE}}, {{TODAY}}", style="Dim.TLabel").grid(row=1, column=2, sticky="w", pady=(6, 0))
        act = ttk.Frame(bottom)
        act.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(act, text="Load values...", command=self.load_values).pack(side="left")
        ttk.Button(act, text="Save values...", command=self.save_values).pack(side="left", padx=(6, 0))
        ttk.Button(act, text="Clear the form", command=self.clear_form).pack(side="left", padx=6)
        self.open_after = tk.BooleanVar(value=self.settings.get("open_after", True) is not False)
        ttk.Checkbutton(act, text="Open when made", variable=self.open_after).pack(side="left", padx=(10, 0))
        self.pdf_too = tk.BooleanVar(value=self.settings.get("pdf_too") is True)
        ttk.Checkbutton(act, text="PDF copy too", variable=self.pdf_too).pack(side="left", padx=(10, 0))
        ttk.Button(act, text="Generate", style="Accent.TButton", command=self.generate).pack(side="right")
        ttk.Button(act, text="Preview", command=self.preview).pack(side="right", padx=(8, 8))
        ttk.Button(act, text="Batch from a spreadsheet...", command=self.run_batch).pack(side="right")

        pan = ttk.Panedwindow(root, orient="horizontal")
        pan.pack(fill="both", expand=True, padx=12, pady=(4, 0))
        # panes are made as children OF the paned window - a frame made on the
        # root and added afterwards is laid out underneath it and never seen
        left = ttk.Frame(pan)
        right = ttk.Frame(pan)
        pan.add(left, weight=1)
        pan.add(right, weight=3)

        self.filter_var = tk.StringVar()
        find = ttk.Frame(left)
        find.pack(fill="x", pady=(0, 6))
        ttk.Label(find, text="Find", style="Field.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(find, textvariable=self.filter_var).pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.filter_var.trace_add("write", lambda *_a: self._fill_tree())
        ttk.Button(left, text="Edit this template in Word", command=self.edit_template).pack(side="bottom", fill="x", pady=(6, 0))
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.column("#0", width=270)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._picked())

        self.head = ttk.Label(right, text="", style="Title.TLabel", padding=(10, 2))
        self.head.pack(fill="x")
        holder = ttk.Frame(right)
        holder.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(holder, highlightthickness=0, borderwidth=0, background=self.c["bg"])
        vsb = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.form = ttk.Frame(self.canvas, padding=(10, 4, 14, 10))
        self.form_id = self.canvas.create_window((0, 0), window=self.form, anchor="nw")
        self.form.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.form_id, width=e.width))
        self.canvas.bind("<Enter>", lambda _e: self._wheel(True))
        self.canvas.bind("<Leave>", lambda _e: self._wheel(False))

    def _wheel(self, on):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            if on:
                self.root.bind_all(seq, self._scroll)
            else:
                self.root.unbind_all(seq)

    def _scroll(self, event):
        if isinstance(event.widget, tk.Text):       # a several-line box scrolls itself
            return None
        if getattr(event, "num", 0) == 4:
            step = -2
        elif getattr(event, "num", 0) == 5:
            step = 2
        elif sys.platform == "darwin":
            step = -int(event.delta)
        else:
            step = -int(event.delta / 120) * 2
        top, bottom = self.canvas.yview()
        if step and not (top <= 0.0 and bottom >= 1.0):
            self.canvas.yview_scroll(step, "units")
        return "break"

    def _style_text(self, widget):
        c = self.c
        widget.configure(background=c["field"], foreground=c["field_text"], insertbackground=c["accent"],
                         selectbackground=c["accent"], selectforeground=c["accent_text"], relief="flat", borderwidth=0,
                         highlightthickness=1, highlightbackground=c["field_border"], highlightcolor=c["accent"],
                         font=(theme.ui_family(self.root), 10), padx=6, pady=5, wrap="word")

    def cycle_theme(self, _event=None):
        self._remember()
        self.mode = theme.next_theme(self.mode)
        self.c = theme.apply(self.root, self.mode)
        self.canvas.configure(background=self.c["bg"])
        self.settings["theme"] = self.mode
        paths.save_settings(self.settings)
        if self.template:
            self._build_form()
        self.status(f"Theme: {theme.label(self.mode)} (Ctrl+D for the next)")
        return "break"

    # ------------------------------------------------------- small helpers
    def status(self, text):
        self.status_var.set(text)

    def ask(self, text) -> bool:
        return True if self.quiet else messagebox.askyesno(APP_NAME, text, parent=self.root)

    def tell(self, text, bad=False):
        self.status(text.splitlines()[0])
        if not self.quiet:
            (messagebox.showerror if bad else messagebox.showinfo)(APP_NAME, text, parent=self.root)

    def background(self, label, work, done):
        self.busy = True
        self.status(label)

        def runner():
            try:
                result = work(lambda msg: self.tasks.put(("progress", msg)))
                self.tasks.put(("done", (done, result)))
            except Exception as exc:  # noqa: BLE001
                self.tasks.put(("error", exc))
        threading.Thread(target=runner, daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, payload = self.tasks.get_nowait()
                if kind == "progress":
                    self.status(payload)
                elif kind == "done":
                    self.busy = False
                    payload[0](payload[1])
                else:
                    self.busy = False
                    self.tell(f"That did not work.\n\n{type(payload).__name__}: {payload}", bad=True)
        except queue.Empty:
            pass
        self.root.after(120, self._poll)

    def quit(self):
        self._remember()
        self.settings.update(out_dir=self.out_var.get().strip(), pattern=self.pattern_var.get().strip(),
                             open_after=bool(self.open_after.get()), pdf_too=bool(self.pdf_too.get()),
                             library=self.library.root)
        try:
            self.settings["geometry"] = self.root.geometry()
        except tk.TclError:
            pass
        paths.save_settings(self.settings)
        self.root.destroy()

    # ------------------------------------------------------------ library
    def open_library(self, first=False):
        note = ""
        try:
            self.library.ensure()
            fresh = self.library.is_empty() and not os.listdir(self.library.clauses_dir)
            if fresh or starter.is_ours(self.library.root):
                made = starter.build(self.library.root)
                if fresh:
                    note = f"A starter library with {len(made)} files was made in {self.library.root}"
                elif made:
                    note = f"{len(made)} new starter template(s) and clause(s) were added to the library."
        except OSError as exc:
            self.tell(f"The library folder could not be used.\n\n{self.library.root}\n{exc}", bad=True)
        self.lib_var.set(self.library.root)
        self.settings["library"] = self.library.root
        paths.save_settings(self.settings)
        self.refresh_templates()
        if note:
            self.status(note)
        elif first:
            self.status("Pick a template on the left.")

    def change_library(self):
        got = filedialog.askdirectory(parent=self.root, title="Choose the library folder (a shared or synced folder is fine)",
                                      initialdir=self.library.root if os.path.isdir(self.library.root) else paths.documents_dir())
        if got:
            self.library = Library(got)
            self.template, self.scan = None, None
            self.open_library()
            self._build_form()

    def add_starters(self):
        if not self.ask("Add the starter templates and clauses this library has not been given before?\n\n"
                        "Nothing already there is changed, and starters you have deleted do not come back.\n\n"
                        + self.library.root):
            return
        try:
            made = starter.build(self.library.root)
        except OSError as exc:
            self.tell(f"The starters could not be written.\n\n{exc}", bad=True)
            return
        self.refresh_templates()
        self.status(f"{len(made)} starter file(s) added." if made else "This library already has every starter.")

    def edit_template(self):
        if not self.template:
            self.status("Pick a template first.")
            return
        paths.open_in_system(self.template)
        self.status("Opened in Word. Save and close it there, then press Refresh.")

    def refresh_templates(self):
        keep = self.template
        self._fill_tree()
        if keep and keep not in self.all_paths:
            self.template, self.scan = None, None
            self._build_form()
        elif keep:
            self._remember()
            self._load_template(keep)
        self.status(f"{len(self.all_paths)} template(s) in the library.")

    def _fill_tree(self):
        """List the templates, narrowed by the Find box. Does not touch the form."""
        keep = self.template
        needle = self.filter_var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self.paths = {}
        folders = {}
        found = self.library.templates()
        self.all_paths = {path for _f, _n, path in found}
        for n, (folder, name, path) in enumerate(found):
            if needle and not all(word in f"{folder} {name}".lower() for word in needle.split()):
                continue
            parent = ""
            if folder:
                if folder not in folders:
                    folders[folder] = self.tree.insert("", "end", iid=f"f{len(folders)}", text=folder, open=True)
                parent = folders[folder]
            iid = f"t{n}"
            self.tree.insert(parent, "end", iid=iid, text=name)
            self.paths[iid] = path
        for iid, path in self.paths.items():
            if path == keep:
                self.tree.selection_set(iid)
                self.tree.see(iid)
                break

    def edit_standing(self):
        win = tk.Toplevel(self.root)
        win.title("Standing values")
        win.configure(background=self.c["bg"])
        win.geometry("620x420")
        win.transient(self.root)
        ttk.Label(win, style="Dim.TLabel", justify="left", wraplength=580, padding=(12, 10),
                  text="Name = value, one per line. A field with the same name is filled in for you on every template. "
                       "Start a line with a space to carry the value on (an address). Kept in the library folder, "
                       "so a shared library shares these too.").pack(fill="x")
        buttons = ttk.Frame(win, padding=12)
        buttons.pack(side="bottom", fill="x")
        box = tk.Text(win, height=12)
        self._style_text(box)
        box.configure(wrap="none")
        box.pack(fill="both", expand=True, padx=12)
        box.insert("1.0", self.library.standing_text())

        def save():
            try:
                self.library.save_standing(box.get("1.0", "end"))
            except OSError as exc:
                self.tell(f"Could not save the standing values.\n\n{exc}", bad=True)
                return
            win.destroy()
            if self.template:
                self._remember()
                self._build_form()
            self.status("Standing values saved.")
        ttk.Button(buttons, text="Save", style="Accent.TButton", command=save).pack(side="right")
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=8)
        return win

    # --------------------------------------------------------------- form
    def _picked(self):
        sel = self.tree.selection()
        path = self.paths.get(sel[0]) if sel else None
        if path and path != self.template:
            self._remember()
            self._load_template(path)

    def _load_template(self, path):
        try:
            scan = merge.scan_template(path, self.library)
        except PermissionError:
            self.tell("That template is open in Word, which locks it. Close it there and pick it again.", bad=True)
            return False
        except Exception as exc:  # noqa: BLE001
            self.tell(f"That template could not be read.\n\n{type(exc).__name__}: {exc}", bad=True)
            return False
        self.template, self.scan = path, scan
        self._build_form()
        n = len(scan.fields)
        self.status(f"{os.path.basename(path)}: {n} field(s)"
                    + (f", {len(scan.groups)} clause list(s)" if scan.groups else "") + ".")
        return True

    def _build_form(self):
        for child in self.form.winfo_children():
            child.destroy()
        self.field_vars, self.field_texts, self.switch_vars, self.clause_vars, self.control_vars = {}, {}, {}, {}, {}
        scan = self.scan
        if not self.template or scan is None:
            self.head.configure(text="")
            ttk.Label(self.form, style="Dim.TLabel", justify="left", wraplength=640,
                      text="Pick a template on the left.\n\nA template is an ordinary Word document with {{tags}} in it. "
                           "\"Open folder\" shows the library; READ ME.txt in there lists every tag.").pack(anchor="w", pady=20)
            return
        self.head.configure(text=os.path.splitext(os.path.basename(self.template))[0])
        standing = self.library.standing()
        self.form.columnconfigure(0, weight=1)
        row = 0
        if scan.problems:
            ttk.Label(self.form, style="Warn.TLabel", justify="left", wraplength=700,
                      text="Check the template:\n" + "\n".join("  - " + p for p in scan.problems)
                      ).grid(row=row, column=0, sticky="w", pady=(0, 8))
            row += 1
        if scan.fields:
            box = ttk.LabelFrame(self.form, text="Fields  (dimmed names are fine to leave empty)"
                                 if any(f.optional for f in scan.fields) else "Fields", padding=10)
            box.grid(row=row, column=0, sticky="ew", pady=(0, 10))
            box.columnconfigure(1, weight=1)
            row += 1
            for n, f in enumerate(scan.fields):
                value = self.memory.get(f.name) or standing.get(f.name) or merge.resolve_default(f.default)
                ttk.Label(box, text=f.name, style="Dim.TLabel" if f.optional else "Field.TLabel").grid(
                    row=n, column=0, sticky="nw", padx=(0, 10), pady=4)
                if f.choices and "\n" not in (value or ""):
                    var = self.field_vars[f.name] = tk.StringVar(value=value or "")
                    ttk.Combobox(box, textvariable=var, values=f.choices).grid(row=n, column=1, sticky="ew", pady=3)
                elif f.multiline or "\n" in (value or ""):
                    text = tk.Text(box, height=4)
                    self._style_text(text)
                    text.insert("1.0", value or "")
                    text.grid(row=n, column=1, sticky="ew", pady=3)
                    self.field_texts[f.name] = text
                else:
                    var = self.field_vars[f.name] = tk.StringVar(value=value or "")
                    ttk.Entry(box, textvariable=var).grid(row=n, column=1, sticky="ew", pady=3)
        if scan.switches:
            box = ttk.LabelFrame(self.form, text="Optional sections", padding=10)
            box.grid(row=row, column=0, sticky="ew", pady=(0, 10))
            row += 1
            for name in scan.switches:
                var = self.switch_vars[name] = tk.BooleanVar(value=merge.truthy(self.memory.get(name, "")))
                ttk.Checkbutton(box, text=name, variable=var).pack(anchor="w")
        for group in scan.groups:
            box = ttk.LabelFrame(self.form, text=f"Clauses - {group}", padding=10)
            box.grid(row=row, column=0, sticky="ew", pady=(0, 10))
            row += 1
            clauses = self.library.clauses_in(group)
            if not clauses:
                ttk.Label(box, text="No clauses in this folder yet.", style="Dim.TLabel").pack(anchor="w")
            for name, _path in clauses:
                var = self.clause_vars[(group, name)] = tk.BooleanVar(value=self.ticks.get((group, name), False))
                ttk.Checkbutton(box, text=name, variable=var).pack(anchor="w")
            if len(clauses) > 2:
                both = ttk.Frame(box)
                both.pack(anchor="w", pady=(6, 0))
                for label, state in (("All", True), ("None", False)):
                    ttk.Button(both, text=label, command=lambda g=group, s=state: self._tick_group(g, s)).pack(side="left", padx=(0, 6))
        box = ttk.LabelFrame(self.form, text="Document control" + ("" if scan.control else "  (this template has no control tags)"),
                             padding=10)
        box.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        box.columnconfigure(1, weight=1)
        box.columnconfigure(3, weight=1)
        for n, key in enumerate(merge.CONTROL):
            ttk.Label(box, text=merge.CONTROL_LABELS[key], style="Field.TLabel").grid(
                row=n // 2, column=(n % 2) * 2, sticky="w", padx=(0 if n % 2 == 0 else 14, 8), pady=3)
            var = self.control_vars[key] = tk.StringVar(value=self.memory.get(key, ""))
            ttk.Entry(box, textvariable=var).grid(row=n // 2, column=(n % 2) * 2 + 1, sticky="ew", pady=3)
        self.control_table = tk.BooleanVar(value=bool(self.memory.get("__control_table__")))
        ttk.Checkbutton(box, text="Add a document-control table at the end of the document",
                        variable=self.control_table).grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.canvas.yview_moveto(0)

    def _tick_group(self, group, state):
        for (g, _name), var in self.clause_vars.items():
            if g == group:
                var.set(state)

    def _values(self):
        values = {name: var.get() for name, var in self.field_vars.items()}
        values.update({name: text.get("1.0", "end").strip("\n") for name, text in self.field_texts.items()})
        values.update({name: ("yes" if var.get() else "") for name, var in self.switch_vars.items()})
        choices = {}
        for (group, name), var in self.clause_vars.items():
            choices.setdefault(group, [])
            if var.get():
                choices[group].append(name)
        control = {key: var.get().strip() for key, var in self.control_vars.items()}
        return values, choices, control

    def _remember(self):
        """Carry what is on the form over to the next template that asks for it."""
        if not self.template or self.scan is None or not self.control_vars:
            return
        values, _choices, control = self._values()
        self.memory.update(values)
        self.memory.update(control)
        self.memory["__control_table__"] = "1" if self.control_table.get() else ""
        for key, var in self.clause_vars.items():
            self.ticks[key] = bool(var.get())

    def clear_form(self):
        self.memory, self.ticks = {}, {}
        if self.template:
            self._build_form()
        self.status("Form cleared (standing values and defaults are back).")

    def _set_values(self, row) -> int:
        """Put a row of loaded values on the form; returns how many landed."""
        scan = self.scan
        if scan is None:
            return 0
        got = match_keys(row, scan.names() + list(scan.switches) + list(merge.CONTROL))
        self._remember()
        self.memory.update({k: v for k, v in got.items()})
        picks = batch.row_choices(row, scan, self.library, {})
        for group, names in picks.items():
            for name, _p in self.library.clauses_in(group):
                self.ticks[(group, name)] = name in names
        self._build_form()
        return len(got)

    def load_values(self):
        if not self.template:
            self.status("Pick a template first.")
            return
        path = filedialog.askopenfilename(parent=self.root, title="Load values (the first row is used)",
                                          initialdir=self._start_dir(), filetypes=VALUE_FILES)
        if not path:
            return
        try:
            rows = read_rows(path)
        except Exception as exc:  # noqa: BLE001
            self.tell(f"That file could not be read.\n\n{type(exc).__name__}: {exc}", bad=True)
            return
        if not rows:
            self.tell("There are no rows of values in that file.", bad=True)
            return
        self.settings["values_dir"] = os.path.dirname(path)
        n = self._set_values(rows[0])
        self.status(f"{n} value(s) loaded from {os.path.basename(path)}"
                    + (f" (row 1 of {len(rows)} - use Batch for all of them)" if len(rows) > 1 else "") + ".")

    def _start_dir(self):
        d = self.settings.get("values_dir")
        return d if isinstance(d, str) and os.path.isdir(d) else paths.documents_dir()

    def pick_out(self):
        got = filedialog.askdirectory(parent=self.root, title="Where to save documents",
                                      initialdir=self.out_var.get() if os.path.isdir(self.out_var.get()) else paths.documents_dir())
        if got:
            self.out_var.set(got)

    # ----------------------------------------------------------- generate
    def _out_dir(self):
        out = self.out_var.get().strip()
        if not out:
            self.tell("Choose where to save documents first.", bad=True)
            return None
        try:
            os.makedirs(out, exist_ok=True)
        except OSError as exc:
            self.tell(f"That folder cannot be used.\n\n{exc}", bad=True)
            return None
        return out

    def generate(self):
        if self.busy:
            return None
        if not self.template or self.scan is None:
            self.status("Pick a template first.")
            return None
        out = self._out_dir()
        if not out:
            return None
        self._remember()
        values, choices, control = self._values()
        draft = self._draft(values, choices, control)
        if draft is None:
            return None
        tmp, report = draft
        # only the fields that actually ended up in the document, so a blank
        # inside a section that was switched off is not asked about
        empty = self._empties(report)
        if empty and not self.ask(f"{len(empty)} field(s) are empty and will be left blank:\n\n  "
                                  + "\n  ".join(empty[:15]) + ("\n  ..." if len(empty) > 15 else "") + "\n\nGenerate anyway?"):
            shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
            return None
        stem = os.path.splitext(os.path.basename(self.template))[0]
        name = file_name(self.pattern_var.get(), {**control, **values, "TEMPLATE": stem},
                         f"{stem} {dt.date.today().isoformat()}")
        path = unique_path(out, name)
        try:
            shutil.move(tmp, path)
        except OSError as exc:
            self.tell(f"The document could not be saved.\n\n{exc}", bad=True)
            return None
        finally:
            shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
        log_issue(out, stem, os.path.basename(path), values.get("Reference", "") or control.get("DOC_REF", ""),
                  control.get("VERSION", ""))
        problems = []
        if report["missing_clauses"]:
            problems.append("Clauses not found: " + ", ".join(report["missing_clauses"]))
        left = merge.leftover_tags(path)
        if left:
            problems.append("Tags left in the document: " + ", ".join(sorted(set(left))[:10]))
        show = path
        if self.pdf_too.get():
            self.status("Making the PDF copy...")
            self.root.update_idletasks()
            try:
                show = pdf.to_pdf(path)
            except pdf.NoConverter as exc:
                problems.append(str(exc))
        if problems:
            self.tell("The document was made, but check it:\n\n" + "\n".join(problems) + f"\n\n{path}")
        else:
            self.status(f"Made {os.path.basename(path)}" + (" and a PDF copy" if show != path else "") + f" in {out}")
        if self.open_after.get() and not self.quiet:
            paths.open_in_system(show)
        return path

    def _draft(self, values, choices, control):
        """Fill the template into a temporary folder. (path, report), or None
        after saying why not."""
        folder = tempfile.mkdtemp(prefix="draftsman-")
        tmp = os.path.join(folder, "draft.docx")
        try:
            return tmp, merge.render(self.template, tmp, values, self.library, choices, control, self.control_table.get())
        except PermissionError:
            self.tell("The template (or a clause) is open in Word, which locks it. Close it there and try again.", bad=True)
        except Exception as exc:  # noqa: BLE001
            self.tell(f"The document could not be made.\n\n{type(exc).__name__}: {exc}", bad=True)
        shutil.rmtree(folder, ignore_errors=True)
        return None

    def _empties(self, report) -> list:
        optional = {f.name for f in self.scan.fields if f.optional} if self.scan else set()
        return [name for name in report["empty"] if name not in optional]

    def preview(self):
        """The words the document would hold, without making it."""
        if self.busy or not self.template or self.scan is None:
            self.status("Pick a template first.")
            return None
        self._remember()
        draft = self._draft(*self._values())
        if draft is None:
            return None
        tmp, report = draft
        try:
            text = merge.document_text(tmp)
            left = sorted(set(merge.leftover_tags(tmp)))
        finally:
            shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
        notes = []
        empty = self._empties(report)
        if empty:
            notes.append("Empty: " + ", ".join(empty))
        if report["missing_clauses"]:
            notes.append("Clauses not found: " + ", ".join(report["missing_clauses"]))
        if left:
            notes.append("Tags left in: " + ", ".join(left))
        win = tk.Toplevel(self.root)
        win.title("Preview - " + os.path.splitext(os.path.basename(self.template))[0])
        win.configure(background=self.c["bg"])
        win.geometry("760x680")
        win.transient(self.root)
        ttk.Label(win, style="Warn.TLabel" if notes else "Dim.TLabel", justify="left", wraplength=720, padding=(12, 8),
                  text="\n".join(notes) if notes else "Words only - the layout is whatever the template has.").pack(fill="x")
        bar = ttk.Frame(win, padding=(12, 8))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right")
        box = tk.Text(win)
        self._style_text(box)
        sb = ttk.Scrollbar(win, orient="vertical", command=box.yview)
        box.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", pady=(0, 4))
        box.pack(fill="both", expand=True, padx=(12, 0), pady=(0, 4))
        box.insert("1.0", text)
        box.configure(state="disabled")
        return win

    def save_values(self):
        if not self.template or self.scan is None:
            self.status("Pick a template first.")
            return
        self._remember()
        keep = {k: v for k, v in self.memory.items() if not k.startswith("__") and str(v).strip()}
        for (group, name), on in self.ticks.items():
            if on:
                keep[f"clauses:{group}"] = "; ".join(filter(None, [keep.get(f"clauses:{group}", ""), name]))
        if not keep:
            self.status("There is nothing on the form to save yet.")
            return
        suggested = file_name("{{Reference}}", keep, "values") + " - values.json"
        path = filedialog.asksaveasfilename(parent=self.root, title="Save these values to load again later",
                                            initialdir=self._start_dir(), initialfile=suggested,
                                            defaultextension=".json", filetypes=[("Values file", "*.json")])
        if not path:
            return
        try:
            save_values(path, keep)
        except OSError as exc:
            self.tell(f"The values could not be saved.\n\n{exc}", bad=True)
            return
        self.settings["values_dir"] = os.path.dirname(path)
        self.status(f"Values saved to {os.path.basename(path)} - Load values... brings them back.")

    def run_batch(self):
        if self.busy:
            return
        if not self.template or self.scan is None:
            self.status("Pick a template first.")
            return
        path = filedialog.askopenfilename(parent=self.root, title="Spreadsheet with one row per document",
                                          initialdir=self._start_dir(), filetypes=VALUE_FILES)
        if not path:
            return
        try:
            rows = read_rows(path)
        except Exception as exc:  # noqa: BLE001
            self.tell(f"That file could not be read.\n\n{type(exc).__name__}: {exc}", bad=True)
            return
        self.settings["values_dir"] = os.path.dirname(path)
        self.batch_rows(rows, os.path.basename(path))

    def batch_rows(self, rows, source="the spreadsheet"):
        scan = self.scan
        if not rows or scan is None:
            self.tell("There are no rows in that file.", bad=True)
            return
        out = self._out_dir()
        if not out:
            return
        wanted = scan.names() + list(scan.switches) + list(merge.CONTROL)
        matched = sorted(match_keys(rows[0], wanted))
        if not matched:
            self.tell("None of the column headings in that file match this template's fields.\n\n"
                      "The first row must hold the field names, such as:\n  " + ", ".join(scan.names()[:8]), bad=True)
            return
        self._remember()
        values, choices, control = self._values()
        if not self.ask(f"Make {len(rows)} document(s) from {source}?\n\nColumns used: {', '.join(matched)}\n\n"
                        "Anything typed on the form fills the fields the spreadsheet does not have.\n\nSaved in: " + out):
            return
        template, library, table = self.template, self.library, self.control_table.get()
        stem = os.path.splitext(os.path.basename(template))[0]
        pattern = self.pattern_var.get().replace("{{TEMPLATE}}", stem)
        want_pdf = bool(self.pdf_too.get())

        def work(progress):
            reports = batch.run(template, rows, out, pattern, scan, library, values, choices, control, table, progress)
            pdf_ok = want_pdf
            for n, r in enumerate(reports, 1):
                log_issue(out, stem, os.path.basename(r["path"]), str(rows[r["row"] - 1].get("Reference", "")),
                          control.get("VERSION", ""))
                if pdf_ok:
                    progress(f"PDF copy {n} of {len(reports)}")
                    try:
                        pdf.to_pdf(r["path"])
                    except pdf.NoConverter:
                        pdf_ok = False              # no point trying the other ninety-nine
            if want_pdf and not pdf_ok:
                for r in reports[:1]:
                    r["no_pdf"] = True
            return reports

        def done(reports):
            optional = {f.name for f in scan.fields if f.optional}
            for r in reports:
                r["empty"] = [name for name in r["empty"] if name not in optional]
            gaps = [f"row {r['row']}: {', '.join(r['empty'][:6])}" for r in reports if r["empty"]]
            text = f"{len(reports)} document(s) made in {out}"
            if any(r.get("no_pdf") for r in reports):
                text += "\n\nNo PDF copies: that needs Microsoft Word or LibreOffice on this computer."
                gaps = gaps or [""]
            if gaps:
                if any(gaps):
                    text += (f"\n\n{len(gaps)} of them have empty fields:\n  " + "\n  ".join(gaps[:12])
                             + ("\n  ..." if len(gaps) > 12 else ""))
                self.tell(text)
            else:
                self.status(text)
            if not self.quiet:
                paths.open_in_system(out)
        self.background(f"Making {len(rows)} documents...", work, done)


def run(argv=None) -> int:
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


def selftest(library_root, out_dir) -> str:
    """Open the window on the self-test library and use every part of it once."""
    import time
    root = tk.Tk()
    saved, paths.save_settings = paths.save_settings, lambda _s: None       # leave the real settings alone
    try:
        app = App(root, settings={"library": library_root, "out_dir": out_dir}, quiet=True)
        root.update()
        total = len(app.library.templates())
        assert len(app.paths) == total >= 45, f"{len(app.paths)} templates listed, {total} in the library"
        made = 0
        for n, (iid, path) in enumerate(list(app.paths.items())):
            app.tree.selection_set(iid)
            root.update()
            if app.template != path:
                assert app._load_template(path), path
            root.update()
            assert app.scan is not None and not app.scan.problems, (path, app.scan)
            if n % 6 and not path.endswith(("1 Acknowledgement.docx", "3 Response.docx")):
                continue                          # every form is built; every sixth is filled in and made
            for name, var in app.field_vars.items():
                if not var.get():
                    var.set(f"[{name}]")
            for name, text in app.field_texts.items():
                if not text.get("1.0", "end").strip():
                    text.insert("1.0", f"[{name}]\nsecond line")
            for var in list(app.switch_vars.values()) + list(app.clause_vars.values()):
                var.set(True)
            for key, var in app.control_vars.items():
                var.set(f"[{key}]")
            app.control_table.set(True)
            result = app.generate()
            assert result and os.path.isfile(result), f"nothing was made from {path}"
            assert not merge.leftover_tags(result), (path, merge.leftover_tags(result))
            made += 1
        assert app.memory.get("Organisation") == "[Organisation]", "values are not carried between templates"
        assert os.path.isfile(os.path.join(out_dir, "Draftsman log.csv")), "the issue log was not written"
        ack = next(p for p in app.paths.values() if p.endswith("1 Acknowledgement.docx"))
        assert app._load_template(ack)
        assert app._set_values({"requester name": "Alex Morgan", "Reference": "SAR/7"}) == 2
        assert app.field_vars["Requester Name"].get() == "Alex Morgan"
        win = app.preview()
        root.update()
        assert win is not None, "the preview did not open"
        win.destroy()
        app.batch_rows([{"Reference": "B1", "Requester Name": "A"}, {"Reference": "B2", "Requester Name": "B"}])
        done = False
        for _ in range(400):
            root.update()
            if not app.busy and os.path.isfile(os.path.join(out_dir, "1 Acknowledgement - B2.docx")):
                done = True
                break
            time.sleep(0.02)
        assert done, "the batch did not finish"
        app.filter_var.set("breach individuals")
        root.update()
        assert len(app.paths) == 1, f"Find listed {len(app.paths)} templates for 'breach individuals'"
        app.filter_var.set("")
        root.update()
        assert len(app.paths) == total
        refuse = next(p for p in app.paths.values() if p.endswith("7 Refusal - unfounded or excessive.docx"))
        app.memory.pop("Ground", None)
        assert app._load_template(refuse)
        root.update()
        assert app.field_vars["Ground"].get() == "", "a choice field should start empty"
        win = app.edit_standing()
        root.update()
        win.destroy()
        for _ in range(4):
            app.cycle_theme()
            root.update()
        app.clear_form()
        root.update()
    finally:
        paths.save_settings = saved
        try:
            root.destroy()
        except tk.TclError:
            pass
    return (f"{__version__}: {total} forms built, {made} documents made, preview, load values, batch, "
            "find, drop-downs, standing values, themes")
