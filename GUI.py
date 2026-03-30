"""
Modern desktop UI for the Trial Balance Reconciler.
"""

from __future__ import annotations

import json
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    from PySide6.QtCore import QObject, QStandardPaths, QThread, Qt, QUrl, Signal
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFrame,
        QHBoxLayout,
        QLabel,
        QMainWindow,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:
    raise SystemExit(
        "PySide6 is required to run this interface.\n"
        "Install it with: python -m pip install PySide6"
    ) from exc

from final_program import (
    format_outputs,
    read_client_tb,
    read_import_tb,
    reconcile_mvp,
    write_details_workbook,
    write_import_format,
)


APP_TITLE = "Trial Balance Reconciler"
STATE_FILE_NAME = "tb_reconciler_state.json"
MAX_RECENT_PATHS = 8
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}

APP_STYLESHEET = """
QWidget { background:#f7f8f4; color:#15231b; font-family:"Segoe UI Variable","Segoe UI",sans-serif; font-size:10pt; }
QWidget#Root { background:#f7f8f4; }
QFrame#PathCard, QFrame#SideRail, QFrame#BottomSection { background:transparent; border:none; }
QFrame#PathCard { border-bottom:1px solid #d8ddd5; }
QFrame#PathCard[validationState="ok"] { border-bottom:2px solid #5e816d; }
QFrame#PathCard[validationState="warning"] { border-bottom:2px solid #b08b43; }
QFrame#PathCard[validationState="error"] { border-bottom:2px solid #bf7b72; }
QFrame#SideRail { border-left:1px solid #d8ddd5; }
QFrame#BottomSection { border-top:1px solid #d8ddd5; }
QLabel#PageTitle { font-size:24pt; font-weight:700; color:#183528; }
QLabel#PageSubtitle { color:#5a685d; font-size:10.5pt; }
QLabel#SectionTitle { font-size:12pt; font-weight:700; color:#183528; padding-top:4px; }
QLabel#SectionBody, QLabel#SummaryCaption, QLabel#OutputsLabel, QLabel#FieldBody { color:#5a685d; }
QLabel#SummaryCaption, QLabel#OutputsLabel { line-height:1.45; }
QLabel#FieldTitle { font-size:10.5pt; font-weight:700; color:#1f3c2e; }
QLabel#FieldStatus[tone="idle"] { color:#667466; }
QLabel#FieldStatus[tone="ok"] { color:#35634b; }
QLabel#FieldStatus[tone="warning"] { color:#8f6a1f; }
QLabel#FieldStatus[tone="error"] { color:#a54b3d; }
QComboBox#PathCombo { background:white; border:1px solid #cdd4ca; border-radius:12px; padding:8px 12px; min-height:22px; }
QComboBox#PathCombo:focus { border:1px solid #2f6b4f; background:white; }
QComboBox#PathCombo::drop-down { border:0; width:28px; }
QPushButton#PrimaryButton { background:#2f6b4f; color:white; border:0; border-radius:18px; padding:14px 22px; font-weight:700; }
QPushButton#PrimaryButton:hover { background:#285b43; }
QPushButton#PrimaryButton:disabled { background:#aebbb0; color:#eef2ee; }
QPushButton#SecondaryButton { background:white; color:#274b39; border:1px solid #cdd4ca; border-radius:16px; padding:12px 18px; font-weight:600; }
QPushButton#SecondaryButton:hover { border:1px solid #aebbb0; background:#fbfcfa; }
QPushButton#SecondaryButton:disabled { color:#97a39a; background:#f8f9f6; }
QPlainTextEdit#LogPanel { background:white; color:#33443b; border:1px solid #d5ddd4; border-radius:12px; padding:12px; font-family:"Cascadia Code","Consolas",monospace; }
QProgressBar#RunProgress { background:#e3e8e1; border:0; border-radius:6px; text-align:center; min-height:10px; color:#173024; }
QProgressBar#RunProgress::chunk { border-radius:6px; background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #2f6b4f, stop:1 #4f8b69); }
QLabel#StatusBanner[tone="ready"] { color:#4e5f54; }
QLabel#StatusBanner[tone="good"] { color:#35634b; }
QLabel#StatusBanner[tone="error"] { color:#a54b3d; }
QCheckBox#Toggle { color:#1f3c2e; font-weight:600; spacing:10px; }
QCheckBox#Toggle::indicator { width:20px; height:20px; border-radius:10px; border:1px solid #cdd4ca; background:#ffffff; }
QCheckBox#Toggle::indicator:checked { background:#2f6b4f; border:1px solid #2f6b4f; }
"""


def repolish(widget: QWidget) -> None:
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def timestamped(message: str) -> str:
    return f"[{datetime.now().strftime('%H:%M:%S')}] {message}"


def format_count(value: int) -> str:
    return f"{int(value):,}"


def make_local_link(path: Path) -> str:
    return (
        f'<a href="{QUrl.fromLocalFile(str(path)).toString()}" '
        'style="color:#2f6b4f; text-decoration:none; font-weight:600;">'
        f"{path.name}</a>"
    )


@dataclass
class RunConfig:
    client_path: Path
    import_path: Path
    out_dir: Path
    write_mvp: bool


class SettingsStore:
    def __init__(self) -> None:
        self.path = self._state_file_path()
        self.fallback_path = Path.cwd() / STATE_FILE_NAME
        self.data = self._load()

    def _state_file_path(self) -> Path:
        base_dir = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
        return Path(base_dir) / STATE_FILE_NAME if base_dir else Path.cwd() / STATE_FILE_NAME

    def _default_data(self) -> dict:
        return {
            "recent_client_files": [],
            "recent_import_files": [],
            "recent_output_dirs": [],
            "write_mvp": True,
        }

    def _load(self) -> dict:
        load_path = self.path if self.path.exists() else self.fallback_path
        if not load_path.exists():
            return self._default_data()
        try:
            raw = json.loads(load_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return self._default_data()
        data = self._default_data()
        for key, default in data.items():
            value = raw.get(key, default)
            if isinstance(default, list) and isinstance(value, list):
                data[key] = [str(item) for item in value if str(item).strip()]
            elif isinstance(default, bool):
                data[key] = bool(value)
        return data

    def save(self) -> None:
        payload = json.dumps(self.data, indent=2)
        for candidate in (self.path, self.fallback_path):
            try:
                candidate.parent.mkdir(parents=True, exist_ok=True)
                candidate.write_text(payload, encoding="utf-8")
                self.path = candidate
                return
            except OSError:
                continue

    def recent(self, key: str) -> list[str]:
        values = self.data.get(key, [])
        return [str(item) for item in values if str(item).strip()] if isinstance(values, list) else []

    def remember_path(self, key: str, value: str) -> None:
        text = value.strip()
        if not text:
            return
        existing = self.recent(key)
        fresh = [text] + [item for item in existing if item.lower() != text.lower()]
        self.data[key] = fresh[:MAX_RECENT_PATHS]
        self.save()

    def remember_run_config(self, cfg: RunConfig) -> None:
        self.remember_path("recent_client_files", str(cfg.client_path))
        self.remember_path("recent_import_files", str(cfg.import_path))
        self.remember_path("recent_output_dirs", str(cfg.out_dir))
        self.data["write_mvp"] = bool(cfg.write_mvp)
        self.save()

    def write_mvp_default(self) -> bool:
        return bool(self.data.get("write_mvp", True))


class DropPathCard(QFrame):
    pathChanged = Signal(str)

    def __init__(
        self,
        title: str,
        body: str,
        placeholder: str,
        button_text: str,
        browse_kind: str,
        file_filter: str = "All files (*)",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.browse_kind = browse_kind
        self.file_filter = file_filter
        self.setObjectName("PathCard")
        self.setProperty("validationState", "idle")
        self.setAcceptDrops(True)
        self.setMinimumHeight(116)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 10, 0, 12)
        layout.setSpacing(8)

        title_label = QLabel(title, self)
        title_label.setObjectName("FieldTitle")
        layout.addWidget(title_label)

        body_label = QLabel(body, self)
        body_label.setObjectName("FieldBody")
        body_label.setWordWrap(True)
        body_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        layout.addWidget(body_label)

        input_row = QHBoxLayout()
        input_row.setSpacing(12)
        layout.addLayout(input_row)

        self.combo = QComboBox(self)
        self.combo.setObjectName("PathCombo")
        self.combo.setEditable(True)
        self.combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.combo.lineEdit().setPlaceholderText(placeholder)
        self.combo.currentTextChanged.connect(self.pathChanged.emit)
        input_row.addWidget(self.combo, 1)

        self.browse_button = QPushButton(button_text, self)
        self.browse_button.setObjectName("SecondaryButton")
        self.browse_button.clicked.connect(self._browse)
        input_row.addWidget(self.browse_button)

        self.status_label = QLabel("Drop a path here or choose one from the list.", self)
        self.status_label.setObjectName("FieldStatus")
        self.status_label.setProperty("tone", "idle")
        self.status_label.setWordWrap(True)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        layout.addWidget(self.status_label)

    def set_recent_paths(self, paths: list[str]) -> None:
        current = self.path_text()
        self.combo.blockSignals(True)
        self.combo.clear()
        self.combo.addItems(paths)
        self.combo.setCurrentText(current)
        self.combo.blockSignals(False)

    def set_path(self, value: str) -> None:
        self.combo.setCurrentText(value)

    def path_text(self) -> str:
        return self.combo.currentText().strip()

    def set_status(self, tone: str, message: str) -> None:
        self.setProperty("validationState", tone)
        self.status_label.setProperty("tone", tone)
        self.status_label.setText(message)
        repolish(self)
        repolish(self.status_label)

    def set_interactive(self, enabled: bool) -> None:
        self.combo.setEnabled(enabled)
        self.browse_button.setEnabled(enabled)

    def _suggest_dialog_start(self) -> str:
        text = self.path_text()
        if not text:
            return str(Path.cwd())
        path = Path(text).expanduser()
        if path.exists():
            return str(path if path.is_dir() else path.parent)
        return str(path.parent) if path.parent.exists() else str(Path.cwd())

    def _browse(self) -> None:
        start_dir = self._suggest_dialog_start()
        if self.browse_kind == "directory":
            chosen = QFileDialog.getExistingDirectory(self, "Select output folder", start_dir)
        else:
            chosen, _ = QFileDialog.getOpenFileName(self, "Select workbook", start_dir, self.file_filter)
        if chosen:
            self.set_path(chosen)

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            dropped = Path(url.toLocalFile())
            target = dropped if self.browse_kind == "file" else (dropped if dropped.is_dir() else dropped.parent)
            self.set_path(str(target))
            event.acceptProposedAction()
            return
        event.ignore()


class ReconcileWorker(QObject):
    progress = Signal(int, str)
    log = Signal(str)
    success = Signal(dict)
    error = Signal(str)
    finished = Signal()

    def __init__(self, cfg: RunConfig) -> None:
        super().__init__()
        self.cfg = cfg

    def run(self) -> None:
        total_steps = 7 if self.cfg.write_mvp else 6
        current_step = 0

        def advance(message: str) -> None:
            nonlocal current_step
            current_step += 1
            percent = int(round((current_step / total_steps) * 100))
            self.progress.emit(percent, message)
            self.log.emit(message)

        try:
            advance("Preparing output folder...")
            self.cfg.out_dir.mkdir(parents=True, exist_ok=True)
            out_import = self.cfg.out_dir / "tb_to_import_updated.xlsx"
            out_details = self.cfg.out_dir / "tb_mvp_details.xlsx"

            advance("Reading client TB workbook...")
            client_df = read_client_tb(self.cfg.client_path)
            advance("Reading import workbook...")
            import_df = read_import_tb(self.cfg.import_path)
            advance("Reconciling accounts and balances...")
            result = reconcile_mvp(client_df, import_df)
            advance("Writing import-ready workbook...")
            write_import_format(result.updated_import, out_import)

            if self.cfg.write_mvp:
                advance("Writing detailed review workbook...")
                write_details_workbook(result, out_details)

            advance("Applying final workbook formatting...")
            new_accts: set[int] = set()
            if len(result.new_rows_added) and "acct_no" in result.new_rows_added.columns:
                new_accts = set(result.new_rows_added["acct_no"].dropna().astype(float).astype(int).tolist())

            format_outputs(
                out_import,
                out_details if self.cfg.write_mvp else None,
                None,
                new_accts=new_accts,
                changed_balance_rows=result.changed_existing_rows,
                renamed_rows=result.renamed_existing_rows,
            )

            summary_map = {str(row["metric"]): row["value"] for _, row in result.summary.iterrows()}
            counts = {
                "renamed": len(result.renamed_existing_rows),
                "changed": len(result.changed_existing_rows),
                "added": len(result.new_rows_added),
                "removed": len(result.removed_import_rows),
                "output_rows": int(summary_map.get("output row count", len(result.updated_import))),
                "zero_balance_removed": int(summary_map.get("deleted zero-balance import rows", 0)),
            }
            self.progress.emit(100, "Completed.")
            self.log.emit("Completed successfully.")
            self.success.emit(
                {
                    "headline": (
                        f"Run complete. {counts['output_rows']} rows are ready for import, with "
                        f"{counts['renamed']} renamed, {counts['changed']} changed, "
                        f"{counts['added']} added, and {counts['removed']} removed."
                    ),
                    "counts": counts,
                    "outputs": {
                        "import": str(out_import),
                        "details": str(out_details) if self.cfg.write_mvp else "",
                        "folder": str(self.cfg.out_dir),
                    },
                }
            )
        except PermissionError as exc:
            self.error.emit(
                "Permission denied while writing an output workbook.\n"
                "Close any open Excel copies of tb_to_import_updated.xlsx or tb_mvp_details.xlsx, then run again.\n\n"
                f"Details: {exc}"
            )
        except Exception as exc:  # noqa: BLE001
            self.error.emit(f"{exc}\n\n{traceback.format_exc()}")
        finally:
            self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1380, 920)
        self.settings_store = SettingsStore()
        self.worker_thread: QThread | None = None
        self.worker: ReconcileWorker | None = None
        self.latest_output_dir: Path | None = None
        self.busy = False

        self._build_ui()
        self._load_settings()
        self._update_output_hint()
        self._validate_form()

    def _build_ui(self) -> None:
        root = QWidget(self)
        root.setObjectName("Root")
        self.setCentralWidget(root)

        shell = QVBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        scroll = QScrollArea(root)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        shell.addWidget(scroll)

        content = QWidget(scroll)
        content.setMaximumWidth(1120)
        scroll.setWidget(content)

        outer = QVBoxLayout(content)
        outer.setContentsMargins(42, 28, 42, 36)
        outer.setSpacing(24)

        title = QLabel(APP_TITLE, content)
        title.setObjectName("PageTitle")
        outer.addWidget(title)

        subtitle = QLabel(
            "Choose the files, pick an output folder, then run reconcile.",
            content,
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        subtitle.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        outer.addWidget(subtitle)

        workspace = QHBoxLayout()
        workspace.setSpacing(28)
        outer.addLayout(workspace)

        main_column = QVBoxLayout()
        main_column.setSpacing(16)
        workspace.addLayout(main_column, 3)

        side_frame = QFrame(content)
        side_frame.setObjectName("SideRail")
        side_frame.setMaximumWidth(320)
        side_layout = QVBoxLayout(side_frame)
        side_layout.setContentsMargins(24, 0, 0, 0)
        side_layout.setSpacing(18)
        workspace.addWidget(side_frame, 1)

        files_title = QLabel("Files", content)
        files_title.setObjectName("SectionTitle")
        main_column.addWidget(files_title)

        self.client_card = DropPathCard(
            "Client TB workbook",
            "Drop the client trial balance here or pick it from your recent files list.",
            "C:\\path\\to\\client tb.xlsx",
            "Browse",
            "file",
            "Excel files (*.xlsx *.xlsm *.xls)",
            root,
        )
        self.client_card.pathChanged.connect(self._validate_form)
        main_column.addWidget(self.client_card)

        self.import_card = DropPathCard(
            "TB to import workbook",
            "Choose the import template workbook that should be updated with the reconciled results.",
            "C:\\path\\to\\tb to import.xlsx",
            "Browse",
            "file",
            "Excel files (*.xlsx *.xlsm *.xls)",
            root,
        )
        self.import_card.pathChanged.connect(self._validate_form)
        main_column.addWidget(self.import_card)

        self.out_dir_card = DropPathCard(
            "Output folder",
            "Drop a folder here, reuse a recent destination, or type a new one to create it on the next run.",
            "C:\\path\\to\\output folder",
            "Browse",
            "directory",
            parent=root,
        )
        self.out_dir_card.pathChanged.connect(self._validate_form)
        main_column.addWidget(self.out_dir_card)

        options_title = QLabel("Options", side_frame)
        options_title.setObjectName("SectionTitle")
        side_layout.addWidget(options_title)

        self.mvp_checkbox = QCheckBox("Create detailed review workbook", side_frame)
        self.mvp_checkbox.setObjectName("Toggle")
        self.mvp_checkbox.toggled.connect(self._update_output_hint)
        side_layout.addWidget(self.mvp_checkbox)

        self.output_hint_label = QLabel(side_frame)
        self.output_hint_label.setObjectName("SectionBody")
        self.output_hint_label.setWordWrap(True)
        self.output_hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.output_hint_label.setMinimumHeight(58)
        self.output_hint_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        side_layout.addWidget(self.output_hint_label)

        run_title = QLabel("Run", side_frame)
        run_title.setObjectName("SectionTitle")
        side_layout.addWidget(run_title)

        self.run_button = QPushButton("Run Reconcile", side_frame)
        self.run_button.setObjectName("PrimaryButton")
        self.run_button.setMinimumHeight(56)
        self.run_button.setMinimumWidth(220)
        self.run_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.run_button.clicked.connect(self._start_run)
        side_layout.addWidget(self.run_button)

        self.open_output_button = QPushButton("Open Output Folder", side_frame)
        self.open_output_button.setObjectName("SecondaryButton")
        self.open_output_button.setEnabled(False)
        self.open_output_button.setMinimumHeight(46)
        self.open_output_button.clicked.connect(self._open_output_folder)
        side_layout.addWidget(self.open_output_button)

        self.quit_button = QPushButton("Quit", side_frame)
        self.quit_button.setObjectName("SecondaryButton")
        self.quit_button.setMinimumHeight(46)
        self.quit_button.clicked.connect(self.close)
        side_layout.addWidget(self.quit_button)

        summary_title = QLabel("Last Run", side_frame)
        summary_title.setObjectName("SectionTitle")
        side_layout.addWidget(summary_title)

        self.summary_caption = QLabel(
            "Run the workflow to see a simple summary of renamed, changed, added, and removed rows.",
            side_frame,
        )
        self.summary_caption.setObjectName("SummaryCaption")
        self.summary_caption.setWordWrap(True)
        self.summary_caption.setMinimumHeight(38)
        self.summary_caption.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        side_layout.addWidget(self.summary_caption)

        self.outputs_label = QLabel("Output files will appear here after a successful run.", side_frame)
        self.outputs_label.setObjectName("OutputsLabel")
        self.outputs_label.setWordWrap(True)
        self.outputs_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        self.outputs_label.linkActivated.connect(self._open_link)
        side_layout.addWidget(self.outputs_label)
        side_layout.addStretch(1)

        bottom_frame = QFrame(content)
        bottom_frame.setObjectName("BottomSection")
        bottom_layout = QVBoxLayout(bottom_frame)
        bottom_layout.setContentsMargins(0, 18, 0, 0)
        bottom_layout.setSpacing(14)
        outer.addWidget(bottom_frame)

        progress_title = QLabel("Progress", bottom_frame)
        progress_title.setObjectName("SectionTitle")
        bottom_layout.addWidget(progress_title)

        self.status_banner = QLabel("Complete the three inputs to enable reconciliation.", bottom_frame)
        self.status_banner.setObjectName("StatusBanner")
        self.status_banner.setProperty("tone", "ready")
        self.status_banner.setWordWrap(True)
        self.status_banner.setMinimumHeight(28)
        self.status_banner.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        bottom_layout.addWidget(self.status_banner)

        self.progress_bar = QProgressBar(bottom_frame)
        self.progress_bar.setObjectName("RunProgress")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        bottom_layout.addWidget(self.progress_bar)

        log_title = QLabel("Progress Log", bottom_frame)
        log_title.setObjectName("SectionTitle")
        bottom_layout.addWidget(log_title)

        self.log_panel = QPlainTextEdit(bottom_frame)
        self.log_panel.setObjectName("LogPanel")
        self.log_panel.setReadOnly(True)
        self.log_panel.setMinimumHeight(220)
        bottom_layout.addWidget(self.log_panel)

        self._append_log("Application ready. Waiting for input files.")

    def _load_settings(self) -> None:
        client_recent = self.settings_store.recent("recent_client_files")
        import_recent = self.settings_store.recent("recent_import_files")
        out_recent = self.settings_store.recent("recent_output_dirs")

        self.client_card.set_recent_paths(client_recent)
        self.import_card.set_recent_paths(import_recent)
        self.out_dir_card.set_recent_paths(out_recent)

        self.client_card.set_path("")
        self.import_card.set_path("")
        self.out_dir_card.set_path(out_recent[0] if out_recent else str(Path.cwd()))
        self.mvp_checkbox.setChecked(self.settings_store.write_mvp_default())

    def _update_output_hint(self, *_args) -> None:
        if self.mvp_checkbox.isChecked():
            self.output_hint_label.setText(
                "Creates:\n"
                "tb_to_import_updated.xlsx\n"
                "tb_mvp_details.xlsx"
            )
        else:
            self.output_hint_label.setText(
                "Creates:\n"
                "tb_to_import_updated.xlsx\n"
                "Review workbook skipped"
            )

    def _validate_excel_path(self, card: DropPathCard, role_label: str) -> tuple[bool, Path | None]:
        text = card.path_text()
        if not text:
            card.set_status("idle", f"Drop or browse the {role_label.lower()} workbook.")
            return False, None

        path = Path(text).expanduser()
        if not path.exists():
            card.set_status("error", f"{role_label} workbook was not found.")
            return False, None
        if not path.is_file():
            card.set_status("error", f"{role_label} path must point to an Excel workbook.")
            return False, None
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            card.set_status("error", "Use an Excel file with .xlsx, .xlsm, or .xls extension.")
            return False, None

        card.set_status("ok", f"{role_label} workbook is ready.")
        return True, path

    def _validate_output_dir(self) -> tuple[bool, Path | None]:
        text = self.out_dir_card.path_text()
        if not text:
            self.out_dir_card.set_status("idle", "Choose or type the folder where outputs should be written.")
            return False, None

        path = Path(text).expanduser()
        if path.exists():
            if not path.is_dir():
                self.out_dir_card.set_status("error", "Output path must be a folder, not a file.")
                return False, None
            self.out_dir_card.set_status("ok", "Output files will be written into this folder.")
            return True, path

        parent = path.parent if str(path.parent) else Path.cwd()
        if parent.exists() and parent.is_dir():
            self.out_dir_card.set_status("warning", "This folder does not exist yet. It will be created on run.")
            return True, path

        self.out_dir_card.set_status("error", "Output folder cannot be created because its parent path is missing.")
        return False, None

    def _validate_form(self, *_args, update_banner: bool = True) -> bool:
        client_ok, client_path = self._validate_excel_path(self.client_card, "Client TB")
        import_ok, import_path = self._validate_excel_path(self.import_card, "Import")
        output_ok, _ = self._validate_output_dir()

        if client_ok and import_ok and client_path and import_path:
            try:
                if client_path.resolve() == import_path.resolve():
                    self.import_card.set_status(
                        "error",
                        "Choose a different workbook here so the client TB and import template are not the same file.",
                    )
                    import_ok = False
            except OSError:
                pass

        valid = client_ok and import_ok and output_ok
        self.run_button.setEnabled(valid and not self.busy)

        if update_banner:
            if self.busy:
                tone = "ready"
                message = "Reconciliation is running. Follow the progress log below."
            elif valid:
                tone = "good"
                message = "Everything looks ready. Start the reconcile run whenever you are ready."
            else:
                tone = "ready"
                message = "Complete the three inputs to enable reconciliation."

            self.status_banner.setProperty("tone", tone)
            self.status_banner.setText(message)
            repolish(self.status_banner)
        return valid

    def _build_run_config(self) -> RunConfig:
        return RunConfig(
            client_path=Path(self.client_card.path_text()).expanduser(),
            import_path=Path(self.import_card.path_text()).expanduser(),
            out_dir=Path(self.out_dir_card.path_text()).expanduser(),
            write_mvp=self.mvp_checkbox.isChecked(),
        )

    def _set_busy(self, busy: bool, preserve_banner: bool = False) -> None:
        self.busy = busy
        self.client_card.set_interactive(not busy)
        self.import_card.set_interactive(not busy)
        self.out_dir_card.set_interactive(not busy)
        self.mvp_checkbox.setEnabled(not busy)
        if busy:
            self.run_button.setEnabled(False)
        else:
            self._validate_form(update_banner=not preserve_banner)

    def _start_run(self, *_args) -> None:
        if not self._validate_form():
            self._append_log("Run blocked because at least one input is not valid.")
            return

        cfg = self._build_run_config()
        self.settings_store.remember_run_config(cfg)
        self._refresh_recent_lists()

        self._append_log("Starting reconcile run.")
        self._set_busy(True)
        self.progress_bar.setValue(0)
        self.open_output_button.setEnabled(False)
        self.latest_output_dir = None

        self.worker_thread = QThread(self)
        self.worker = ReconcileWorker(cfg)
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.log.connect(self._append_log)
        self.worker.success.connect(self._on_worker_success)
        self.worker.error.connect(self._on_worker_error)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self._on_worker_finished)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def _on_worker_progress(self, percent: int, message: str) -> None:
        self.progress_bar.setValue(percent)
        self.status_banner.setProperty("tone", "ready")
        self.status_banner.setText(message)
        repolish(self.status_banner)

    def _on_worker_success(self, payload: dict) -> None:
        counts = payload["counts"]
        outputs = payload["outputs"]

        self.summary_caption.setText(
            "Import-ready rows: "
            f"{format_count(counts['output_rows'])}\n"
            "Renamed: "
            f"{format_count(counts['renamed'])}   "
            "Changed: "
            f"{format_count(counts['changed'])}   "
            "Added: "
            f"{format_count(counts['added'])}   "
            "Removed: "
            f"{format_count(counts['removed'])}\n"
            "Zero-balance removals: "
            f"{format_count(counts['zero_balance_removed'])}"
        )

        import_path = Path(outputs["import"])
        details_path = Path(outputs["details"]) if outputs["details"] else None
        folder_path = Path(outputs["folder"])
        self.latest_output_dir = folder_path
        self.open_output_button.setEnabled(True)

        lines = [
            f"Import-ready workbook: {make_local_link(import_path)}",
            f"Output folder: {make_local_link(folder_path)}",
        ]
        if details_path is not None:
            lines.insert(1, f"Details workbook: {make_local_link(details_path)}")
        else:
            lines.insert(1, "Details workbook: skipped for this run.")
        self.outputs_label.setText("<br>".join(lines))

        self.progress_bar.setValue(100)
        self.status_banner.setProperty("tone", "good")
        self.status_banner.setText("Reconciliation finished successfully.")
        repolish(self.status_banner)

    def _on_worker_error(self, message: str) -> None:
        self.status_banner.setProperty("tone", "error")
        self.status_banner.setText("The run stopped before completion. Review the log for details.")
        repolish(self.status_banner)
        self._append_log("Run failed.")
        for line in message.splitlines():
            self._append_log(line)

    def _on_worker_finished(self) -> None:
        self.worker_thread = None
        self.worker = None
        self._set_busy(False, preserve_banner=True)

    def _append_log(self, message: str) -> None:
        self.log_panel.appendPlainText(timestamped(message))
        scrollbar = self.log_panel.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _open_output_folder(self, *_args) -> None:
        if self.latest_output_dir is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.latest_output_dir)))

    def _open_link(self, url: str) -> None:
        QDesktopServices.openUrl(QUrl(url))

    def _refresh_recent_lists(self) -> None:
        self.client_card.set_recent_paths(self.settings_store.recent("recent_client_files"))
        self.import_card.set_recent_paths(self.settings_store.recent("recent_import_files"))
        self.out_dir_card.set_recent_paths(self.settings_store.recent("recent_output_dirs"))

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.client_card.path_text():
            self.settings_store.remember_path("recent_client_files", self.client_card.path_text())
        if self.import_card.path_text():
            self.settings_store.remember_path("recent_import_files", self.import_card.path_text())
        if self.out_dir_card.path_text():
            self.settings_store.remember_path("recent_output_dirs", self.out_dir_card.path_text())
        self.settings_store.data["write_mvp"] = self.mvp_checkbox.isChecked()
        self.settings_store.save()
        super().closeEvent(event)


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationDisplayName(APP_TITLE)
    app.setOrganizationName("DSA")
    app.setApplicationName("AccountingProject")
    app.setStyleSheet(APP_STYLESHEET)

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
