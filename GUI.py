"""
tb_gui.py

A simple Tkinter desktop app for the TB workflow.
"""

from __future__ import annotations

import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Set

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from final_program import (
    format_outputs,
    read_client_tb,
    read_import_tb,
    reconcile_mvp,
    write_details_workbook,
    write_import_format,
)


@dataclass
class RunConfig:
    client_path: Path
    import_path: Path
    out_dir: Path
    write_mvp: bool


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trial Balance Reconciler")
        self._windowed_geometry = "980x560"
        self._windowed_minsize = (960, 540)
        self.geometry(self._windowed_geometry)
        self.minsize(*self._windowed_minsize)
        self.configure(bg="#eef3f9")

        self.client_var = tk.StringVar()
        self.import_var = tk.StringVar()
        self.out_dir_var = tk.StringVar(value=str(Path.cwd()))
        self.mvp_var = tk.BooleanVar(value=True)
        self.output_hint_var = tk.StringVar()

        self.status_var = tk.StringVar(value="Choose the files and output folder to begin.")
        self.busy_var = tk.BooleanVar(value=False)
        self._interactive_widgets: list[tk.Widget] = []

        self.mvp_var.trace_add("write", self._update_output_hint)
        self._configure_styles()
        self._update_output_hint()
        self._build_ui()
        self.after(0, self._maximize_window)

    def _maximize_window(self):
        try:
            self.state("zoomed")
        except tk.TclError:
            self.geometry(
                f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0"
            )

    def _configure_styles(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("App.TFrame", background="#eef3f9")
        style.configure("Hero.TFrame", background="#173a63")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Inset.TFrame", background="#f6f9fc")
        style.configure("HeroTitle.TLabel", background="#173a63", foreground="#ffffff", font=("Segoe UI", 21, "bold"))
        style.configure("HeroSub.TLabel", background="#173a63", foreground="#d8e7f7", font=("Segoe UI", 10))
        style.configure("SectionTitle.TLabel", background="#ffffff", foreground="#16324f", font=("Segoe UI", 14, "bold"))
        style.configure("Field.TLabel", background="#ffffff", foreground="#34516f", font=("Segoe UI", 10, "bold"))
        style.configure("Body.TLabel", background="#ffffff", foreground="#4d647d", font=("Segoe UI", 10))
        style.configure("StatusValue.TLabel", background="#ffffff", foreground="#1f3650", font=("Segoe UI", 10))
        style.configure("Info.TLabel", background="#f6f9fc", foreground="#61778f", font=("Segoe UI", 9))
        style.configure("Toggle.TCheckbutton", background="#f6f9fc", foreground="#213f5e", font=("Segoe UI", 10, "bold"))
        style.map("Toggle.TCheckbutton", foreground=[("disabled", "#95a6b8")])
        style.configure(
            "Path.TEntry",
            fieldbackground="#f8fbff",
            bordercolor="#c9d6e4",
            lightcolor="#c9d6e4",
            darkcolor="#c9d6e4",
            padding=6,
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(18, 8),
            background="#1f6fd1",
            foreground="#ffffff",
            borderwidth=0,
            focuscolor="",
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#195aa9"), ("disabled", "#a9c2df")],
            foreground=[("disabled", "#eef4fa")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(14, 8),
            background="#eef4fa",
            foreground="#29435f",
            bordercolor="#d0dceb",
            lightcolor="#d0dceb",
            darkcolor="#d0dceb",
            focuscolor="",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#dfeaf5"), ("disabled", "#f4f7fa")],
            foreground=[("disabled", "#95a6b8")],
        )
        style.configure(
            "App.Horizontal.TProgressbar",
            troughcolor="#e6edf5",
            background="#1f6fd1",
            bordercolor="#e6edf5",
            lightcolor="#1f6fd1",
            darkcolor="#1f6fd1",
        )

    def _update_output_hint(self, *_args):
        if self.mvp_var.get():
            self.output_hint_var.set(
                "Always created:\n"
                "  tb_to_import_updated.xlsx\n\n"
                "Optional review file:\n"
                "  tb_mvp_details.xlsx"
            )
        else:
            self.output_hint_var.set(
                "Only created:\n"
                "  tb_to_import_updated.xlsx\n\n"
                "MVP details workbook is skipped."
            )

    def _build_ui(self):
        outer = ttk.Frame(self, style="App.TFrame", padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=3)
        outer.columnconfigure(1, weight=2)
        outer.rowconfigure(1, weight=1)

        hero = ttk.Frame(outer, style="Hero.TFrame", padding=(22, 18))
        hero.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 16))
        hero.columnconfigure(0, weight=1)

        ttk.Label(hero, text="Trial Balance Reconciler", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            hero,
            text="Update the import-ready workbook from the client trial balance and optionally create the MVP review file.",
            style="HeroSub.TLabel",
            wraplength=780,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        form_card = ttk.Frame(outer, style="Card.TFrame", padding=(20, 18))
        form_card.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        form_card.columnconfigure(0, weight=1)
        form_card.columnconfigure(1, weight=1)
        form_card.columnconfigure(2, weight=1)

        ttk.Label(form_card, text="Files", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(
            form_card,
            text="Pick the source workbooks and where the updated files should be saved.",
            style="Body.TLabel",
            wraplength=560,
            justify="left",
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 16))

        ttk.Label(form_card, text="Client TB file", style="Field.TLabel").grid(row=2, column=0, columnspan=4, sticky="w")
        client_entry = ttk.Entry(form_card, textvariable=self.client_var, style="Path.TEntry")
        client_entry.grid(row=3, column=0, columnspan=3, sticky="we", pady=(6, 12))
        client_btn = ttk.Button(form_card, text="Browse", command=self.pick_client, style="Secondary.TButton")
        client_btn.grid(row=3, column=3, sticky="e", padx=(12, 0), pady=(6, 12))

        ttk.Label(form_card, text="TB to import file", style="Field.TLabel").grid(row=4, column=0, columnspan=4, sticky="w")
        import_entry = ttk.Entry(form_card, textvariable=self.import_var, style="Path.TEntry")
        import_entry.grid(row=5, column=0, columnspan=3, sticky="we", pady=(6, 12))
        import_btn = ttk.Button(form_card, text="Browse", command=self.pick_import, style="Secondary.TButton")
        import_btn.grid(row=5, column=3, sticky="e", padx=(12, 0), pady=(6, 12))

        ttk.Label(form_card, text="Output folder", style="Field.TLabel").grid(row=6, column=0, columnspan=4, sticky="w")
        out_dir_entry = ttk.Entry(form_card, textvariable=self.out_dir_var, style="Path.TEntry")
        out_dir_entry.grid(row=7, column=0, columnspan=3, sticky="we", pady=(6, 18))
        out_dir_btn = ttk.Button(form_card, text="Browse", command=self.pick_out_dir, style="Secondary.TButton")
        out_dir_btn.grid(row=7, column=3, sticky="e", padx=(12, 0), pady=(6, 18))

        self.progress = ttk.Progressbar(form_card, mode="indeterminate", style="App.Horizontal.TProgressbar")
        self.progress.grid(row=8, column=0, columnspan=4, sticky="we")

        ttk.Label(form_card, textvariable=self.status_var, style="StatusValue.TLabel", wraplength=600, justify="left").grid(
            row=9, column=0, columnspan=4, sticky="w", pady=(10, 0)
        )

        btns = ttk.Frame(form_card, style="Card.TFrame")
        btns.grid(row=10, column=0, columnspan=4, sticky="e", pady=(18, 0))

        self.run_btn = ttk.Button(btns, text="Run Reconcile", command=self.on_run, style="Primary.TButton")
        self.run_btn.pack(side="right")

        quit_btn = ttk.Button(btns, text="Quit", command=self.destroy, style="Secondary.TButton")
        quit_btn.pack(side="right", padx=(0, 10))

        side_card = ttk.Frame(outer, style="Card.TFrame", padding=(18, 18))
        side_card.grid(row=1, column=1, sticky="nsew")
        side_card.columnconfigure(0, weight=1)

        ttk.Label(side_card, text="Options", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            side_card,
            text="Choose whether to create the optional workbook used for matching review.",
            style="Body.TLabel",
            wraplength=280,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 14))

        toggle_panel = ttk.Frame(side_card, style="Inset.TFrame", padding=(14, 12))
        toggle_panel.grid(row=2, column=0, sticky="nsew")
        toggle_panel.columnconfigure(0, weight=1)

        mvp_toggle = ttk.Checkbutton(
            toggle_panel,
            text="Create MVP spreadsheet",
            variable=self.mvp_var,
            style="Toggle.TCheckbutton",
        )
        mvp_toggle.grid(row=0, column=0, sticky="w")

        ttk.Label(
            toggle_panel,
            text="This creates tb_mvp_details.xlsx so you can review rows that were renamed, updated, added, or removed.",
            style="Info.TLabel",
            wraplength=260,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Label(side_card, text="Files Created", style="SectionTitle.TLabel").grid(row=3, column=0, sticky="w", pady=(18, 0))
        ttk.Label(
            side_card,
            textvariable=self.output_hint_var,
            style="Body.TLabel",
            wraplength=280,
            justify="left",
        ).grid(row=4, column=0, sticky="w", pady=(8, 0))

        ttk.Label(side_card, text="Run Status", style="SectionTitle.TLabel").grid(row=5, column=0, sticky="w", pady=(18, 0))
        ttk.Label(
            side_card,
            textvariable=self.status_var,
            style="Body.TLabel",
            wraplength=280,
            justify="left",
        ).grid(row=6, column=0, sticky="w", pady=(8, 0))

        self._interactive_widgets = [
            client_entry,
            import_entry,
            out_dir_entry,
            client_btn,
            import_btn,
            out_dir_btn,
            mvp_toggle,
            quit_btn,
        ]

    def pick_client(self):
        path = filedialog.askopenfilename(
            title="Select client tb file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.client_var.set(path)

    def pick_import(self):
        path = filedialog.askopenfilename(
            title="Select tb to import file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if path:
            self.import_var.set(path)

    def pick_out_dir(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.out_dir_var.set(path)

    def _set_busy(self, busy: bool):
        self.busy_var.set(busy)
        self.run_btn.configure(state=("disabled" if busy else "normal"))
        self.run_btn.configure(text=("Running..." if busy else "Run Reconcile"))
        for widget in self._interactive_widgets:
            widget.configure(state=("disabled" if busy else "normal"))
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def on_run(self):
        client_path = Path(self.client_var.get().strip())
        import_path = Path(self.import_var.get().strip())
        out_dir = Path(self.out_dir_var.get().strip())

        if not client_path.exists():
            messagebox.showerror("Missing file", "Please select a valid client TB file.")
            return
        if not import_path.exists():
            messagebox.showerror("Missing file", "Please select a valid tb to import file.")
            return
        if not out_dir.exists():
            messagebox.showerror("Missing folder", "Please select a valid output folder.")
            return

        cfg = RunConfig(
            client_path=client_path,
            import_path=import_path,
            out_dir=out_dir,
            write_mvp=bool(self.mvp_var.get()),
        )

        self._set_busy(True)
        self.status_var.set("Running reconcile and preparing the selected output files.")
        threading.Thread(target=self._run_worker, args=(cfg,), daemon=True).start()

    def _run_worker(self, cfg: RunConfig):
        try:
            out_import = cfg.out_dir / "tb_to_import_updated.xlsx"
            out_details = cfg.out_dir / "tb_mvp_details.xlsx"

            client_df = read_client_tb(cfg.client_path)
            import_df = read_import_tb(cfg.import_path)
            result = reconcile_mvp(client_df, import_df)

            write_import_format(result.updated_import, out_import)
            if cfg.write_mvp:
                write_details_workbook(result, out_details)

            new_accts: Set[int] = set()
            if getattr(result, "new_rows_added", None) is not None and len(result.new_rows_added):
                if "acct_no" in result.new_rows_added.columns:
                    new_accts = set(result.new_rows_added["acct_no"].dropna().astype(float).astype(int).tolist())

            if format_outputs is not None:
                format_outputs(
                    out_import,
                    out_details if cfg.write_mvp else None,
                    None,
                    new_accts=new_accts,
                    changed_balance_rows=getattr(result, "changed_existing_rows", None),
                    renamed_rows=getattr(result, "renamed_existing_rows", None),
                )

            msg = "Finished.\n\n" f"Import ready:\n{out_import}\n"
            if cfg.write_mvp:
                msg += f"\nDetails:\n{out_details}\n"
            self._ui_success(msg)

        except PermissionError as e:
            locked_files = ["tb_to_import_updated.xlsx"]
            if cfg.write_mvp:
                locked_files.append("tb_mvp_details.xlsx")
            friendly = (
                "Permission denied while writing an output file.\n\n"
                "Close any of these files if they are open in Excel, then run again:\n"
                + "\n".join(locked_files)
                + "\n\n"
                f"Details:\n{e}"
            )
            self._ui_error("File locked", friendly)

        except Exception as e:
            tb = traceback.format_exc()
            self._ui_error("Error", f"{e}\n\n{tb}")

        finally:
            self._ui_done()

    def _ui_success(self, msg: str):
        def _():
            self.status_var.set("Done. Output files are ready in the selected folder.")
            messagebox.showinfo("Done", msg)

        self.after(0, _)

    def _ui_error(self, title: str, msg: str):
        def _():
            self.status_var.set("Something went wrong. Review the message for details.")
            messagebox.showerror(title, msg)

        self.after(0, _)

    def _ui_done(self):
        def _():
            self._set_busy(False)
            if self.status_var.get() not in (
                "Done. Output files are ready in the selected folder.",
                "Something went wrong. Review the message for details.",
            ):
                self.status_var.set("Ready to run.")

        self.after(0, _)


if __name__ == "__main__":
    App().mainloop()
