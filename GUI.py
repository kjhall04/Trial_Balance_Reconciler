"""
Modern desktop UI for the Trial Balance Builder.
"""

from __future__ import annotations

import json
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    from PySide6.QtCore import QObject, QSize, QStandardPaths, QThread, Qt, QUrl, Signal
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

from trial_balance_pipeline import (
    WorkbookSpec,
    build_from_workbooks,
    format_outputs,
    split_path_text,
    write_details_workbook,
    write_import_workbook,
    write_review_workbook,
)


APP_TITLE = "Trial Balance Builder"
STATE_FILE_NAME = "tb_reconciler_state.json"
MAX_RECENT_PATHS = 8
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
APP_STYLESHEET = """
QWidget { background:#f7f8f4; color:#15231b; font-family:"Segoe UI Variable","Segoe UI",sans-serif; font-size:10pt; }
QWidget#Root { background:#f7f8f4; }
QFrame#PathCard, QFrame#SideRail, QFrame#BottomSection { background:transparent; border:none; }
QFrame#AdvancedPanel { background:#fbfcfa; border:1px solid #d8ddd5; border-radius:18px; }
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
QPlainTextEdit#PreviewPanel { background:white; color:#33443b; border:1px solid #d5ddd4; border-radius:12px; padding:12px; }
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
    client_specs: list[WorkbookSpec]
    prior_specs: list[WorkbookSpec]
    out_dir: Path
    write_mvp: bool
    delete_zero_balance_rows: bool


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
            "recent_prior_files": [],
            "recent_output_dirs": [],
            "write_mvp": True,
            "delete_zero_balance_rows": True,
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
            elif isinstance(default, str):
                data[key] = str(value)
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
        joined_client_paths = " | ".join(str(spec.path) for spec in cfg.client_specs)
        joined_prior_paths = " | ".join(str(spec.path) for spec in cfg.prior_specs)
        self.remember_path("recent_client_files", joined_client_paths)
        self.remember_path("recent_prior_files", joined_prior_paths)
        self.remember_path("recent_output_dirs", str(cfg.out_dir))
        self.data["write_mvp"] = bool(cfg.write_mvp)
        self.data["delete_zero_balance_rows"] = bool(cfg.delete_zero_balance_rows)
        self.save()

    def write_mvp_default(self) -> bool:
        return bool(self.data.get("write_mvp", True))

    def delete_zero_balance_rows_default(self) -> bool:
        return bool(self.data.get("delete_zero_balance_rows", True))


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
        allow_multiple: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.browse_kind = browse_kind
        self.file_filter = file_filter
        self.allow_multiple = allow_multiple
        self.setObjectName("PathCard")
        self.setProperty("validationState", "idle")
        self.setAcceptDrops(True)
        self.setMinimumHeight(116)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

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
        browse_size = self.browse_button.sizeHint()
        self.browse_button.setFixedWidth(browse_size.width())
        self.browse_button.setFixedHeight(browse_size.height() + 2)
        self.browse_button.clicked.connect(self._browse)
        input_row.addWidget(self.browse_button)

        self.status_label = QLabel("Drop a path here or choose one from the list.", self)
        self.status_label.setObjectName("FieldStatus")
        self.status_label.setProperty("tone", "idle")
        self.status_label.setWordWrap(True)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        layout.addWidget(self.status_label)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        layout = self.layout()
        if layout is None:
            return super().heightForWidth(width)
        return max(self.minimumHeight(), layout.totalHeightForWidth(max(0, width)))

    def sizeHint(self) -> QSize:
        layout = self.layout()
        if layout is None:
            return super().sizeHint()
        hint = layout.sizeHint()
        return QSize(hint.width(), max(self.minimumHeight(), hint.height()))

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

    def path_texts(self) -> list[str]:
        if self.allow_multiple:
            return split_path_text(self.path_text())
        text = self.path_text()
        return [text] if text else []

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
        elif self.allow_multiple:
            chosen_paths, _ = QFileDialog.getOpenFileNames(self, "Select workbooks", start_dir, self.file_filter)
            chosen = " | ".join(chosen_paths)
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
        dropped_paths: list[str] = []
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            dropped = Path(url.toLocalFile())
            target = dropped if self.browse_kind == "file" else (dropped if dropped.is_dir() else dropped.parent)
            dropped_paths.append(str(target))

        if dropped_paths:
            value = " | ".join(dropped_paths) if self.allow_multiple else dropped_paths[0]
            self.set_path(value)
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
        total_steps = 6 if self.cfg.write_mvp else 5
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
            out_details = self.cfg.out_dir / "tb_build_details.xlsx"
            out_review = self.cfg.out_dir / "tb_review_style.xlsx"

            advance("Reading workbook data...")
            result = build_from_workbooks(
                current_specs=self.cfg.client_specs,
                prior_specs=self.cfg.prior_specs,
                keep_zero_rows=not self.cfg.delete_zero_balance_rows,
            )
            advance("Writing import-ready workbook...")
            write_import_workbook(result.updated_import, out_import)

            if self.cfg.write_mvp:
                advance("Writing audit details workbook...")
                write_details_workbook(result, out_details)
                advance("Writing review-style workbook...")
                write_review_workbook(result.updated_import, out_review, "Review Trial Balance")

            advance("Applying final workbook formatting...")
            format_outputs(
                out_import,
                out_details if self.cfg.write_mvp else None,
                out_review if self.cfg.write_mvp else None,
            )

            summary_map = {str(row["metric"]): row["value"] for _, row in result.summary.iterrows()}
            counts = {
                "prior_rows": int(summary_map.get("prior-year rows parsed", len(result.prior_year_rows))),
                "matched": int(summary_map.get("matched to prior year", len(result.matched_rows))),
                "new_rows": int(summary_map.get("new current-year rows", len(result.new_rows))),
                "carryforward": int(summary_map.get("carryforward prior-year rows", len(result.carryforward_rows))),
                "renumbered": int(summary_map.get("renumbered rows", len(result.renumbered_rows))),
                "review_queue": int(summary_map.get("review queue rows", len(result.review_queue))),
                "high_confidence": int(summary_map.get("high-confidence rows", 0)),
                "medium_confidence": int(summary_map.get("medium-confidence rows", 0)),
                "low_confidence": int(summary_map.get("low-confidence rows", 0)),
                "output_rows": int(summary_map.get("output row count", len(result.updated_import))),
                "current_total": float(summary_map.get("output current-year total", float(result.updated_import["cy_balance"].sum()))),
            }
            self.progress.emit(100, "Completed.")
            self.log.emit("Completed successfully.")
            self.success.emit(
                {
                    "headline": (
                        f"Run complete. {counts['output_rows']} rows are ready for import, with "
                        f"{counts['matched']} matched to prior year, {counts['new_rows']} new, "
                        f"{counts['carryforward']} carried forward, {counts['renumbered']} renumbered, "
                        f"and {counts['review_queue']} waiting for review. Confidence levels: "
                        f"{counts['high_confidence']} high, {counts['medium_confidence']} medium, "
                        f"{counts['low_confidence']} low."
                    ),
                    "ready_for_import": bool(result.ready_for_import),
                    "used_prior_comparison": bool(counts["prior_rows"]),
                    "counts": counts,
                    "outputs": {
                        "import": str(out_import),
                        "details": str(out_details) if self.cfg.write_mvp else "",
                        "review": str(out_review) if self.cfg.write_mvp else "",
                        "folder": str(self.cfg.out_dir),
                    },
                }
            )
        except PermissionError as exc:
            self.error.emit(
                "Permission denied while writing an output workbook.\n"
                "Close any open Excel copies of tb_to_import_updated.xlsx, tb_build_details.xlsx, or tb_review_style.xlsx, then run again.\n\n"
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
            "Choose the current-year files, optional prior-year official TB files, and an output folder. The program builds a fresh trial balance, compares it to last year when those official TB workbooks are provided, and adds confidence highlighting plus an audit trail workbook.",
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
            "Client trial balance file(s)",
            "Choose the current-year trial balance files you received from the client. If the job should be combined into one final TB, add every current-year file here.",
            "C:\\path\\to\\client tb.xlsx | C:\\path\\to\\ORE TB.xlsx",
            "Browse",
            "file",
            "Excel files (*.xlsx *.xlsm *.xls)",
            allow_multiple=True,
            parent=root,
        )
        self.client_card.pathChanged.connect(self._validate_form)
        main_column.addWidget(self.client_card)

        self.import_card = DropPathCard(
            "Prior-year official TB file(s) (Optional)",
            "Choose the previous year's official trial balance workbook(s) that the new trial balance should line up against. Do not use last year's TB import workbook here.",
            "C:\\path\\to\\2022 Trial Balance.xlsx | C:\\path\\to\\HR 2022 Trial Balance.xlsx",
            "Browse",
            "file",
            "Excel files (*.xlsx *.xlsm *.xls)",
            allow_multiple=True,
            parent=root,
        )
        self.import_card.pathChanged.connect(self._validate_form)
        main_column.addWidget(self.import_card)

        prior_note = QLabel(
            "Leave the prior-year field blank if you only want to build from the current-year files. The run will still finish, but unmatched rows will stay flagged for review.",
            content,
        )
        prior_note.setObjectName("SectionBody")
        prior_note.setWordWrap(True)
        prior_note.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        main_column.addWidget(prior_note)

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

        self.mvp_checkbox = QCheckBox("Create audit details workbook", side_frame)
        self.mvp_checkbox.setObjectName("Toggle")
        self.mvp_checkbox.toggled.connect(self._update_output_hint)
        side_layout.addWidget(self.mvp_checkbox)

        self.zero_rows_checkbox = QCheckBox("Remove zero-balance rows from final file", side_frame)
        self.zero_rows_checkbox.setObjectName("Toggle")
        self.zero_rows_checkbox.toggled.connect(self._update_output_hint)
        side_layout.addWidget(self.zero_rows_checkbox)

        self.output_hint_label = QLabel(side_frame)
        self.output_hint_label.setObjectName("SectionBody")
        self.output_hint_label.setWordWrap(True)
        self.output_hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.output_hint_label.setMinimumHeight(76)
        self.output_hint_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        side_layout.addWidget(self.output_hint_label)

        run_title = QLabel("Run", side_frame)
        run_title.setObjectName("SectionTitle")
        side_layout.addWidget(run_title)

        self.run_button = QPushButton("Build Trial Balance", side_frame)
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
            "Run the workflow to see a simple summary of matched, new, carryforward, and renumbered rows.",
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

        self.status_banner = QLabel("Complete the current-year workbook inputs and output folder to enable reconciliation.", bottom_frame)
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
        import_recent = self.settings_store.recent("recent_prior_files")
        out_recent = self.settings_store.recent("recent_output_dirs")

        self.client_card.set_recent_paths(client_recent)
        self.import_card.set_recent_paths(import_recent)
        self.out_dir_card.set_recent_paths(out_recent)

        self.client_card.set_path("")
        self.import_card.set_path("")
        self.out_dir_card.set_path(out_recent[0] if out_recent else str(Path.cwd()))
        self.mvp_checkbox.setChecked(self.settings_store.write_mvp_default())
        self.zero_rows_checkbox.setChecked(self.settings_store.delete_zero_balance_rows_default())

    def _update_output_hint(self, *_args) -> None:
        zero_row_policy = (
            "Zero-balance rows: deleted from the final import workbook"
            if self.zero_rows_checkbox.isChecked()
            else "Zero-balance rows: kept in the final import workbook"
        )
        if self.mvp_checkbox.isChecked():
            self.output_hint_label.setText(
                "Creates:\n"
                "tb_to_import_updated.xlsx\n"
                "tb_build_details.xlsx\n"
                "tb_review_style.xlsx\n"
                "Confidence colors: green/high, yellow/medium, red/low\n"
                "Single combined TB output hides the entity column automatically\n"
                f"{zero_row_policy}"
            )
        else:
            self.output_hint_label.setText(
                "Creates:\n"
                "tb_to_import_updated.xlsx\n"
                "Audit details and review workbooks skipped\n"
                "Single combined TB output hides the entity column automatically\n"
                f"{zero_row_policy}"
            )

    def _validate_prior_inputs(self) -> tuple[bool, list[Path]]:
        prior_paths = [Path(text).expanduser() for text in self.import_card.path_texts()]
        if not prior_paths:
            self.import_card.set_status(
                "idle",
                "Optional. Add prior-year official TB workbook files here when you want a last-year comparison.",
            )
            return True, []

        valid_paths: list[Path] = []
        for path in prior_paths:
            if not path.exists():
                self.import_card.set_status("error", f"Prior-year trial balance file was not found: {path}")
                return False, []
            if not path.is_file():
                self.import_card.set_status("error", "Each prior-year path must point to an Excel workbook.")
                return False, []
            if path.suffix.lower() not in EXCEL_SUFFIXES:
                self.import_card.set_status("error", "Use Excel files with .xlsx, .xlsm, or .xls extensions.")
                return False, []
            valid_paths.append(path)

        workbook_count = len(valid_paths)
        noun = "file" if workbook_count == 1 else "files"
        self.import_card.set_status("ok", f"{workbook_count} prior-year official TB {noun} ready.")
        return True, valid_paths

    def _validate_client_inputs(self) -> tuple[bool, list[Path]]:
        raw_paths = [Path(text).expanduser() for text in self.client_card.path_texts()]
        if not raw_paths:
            self.client_card.set_status("idle", "Drop or browse one or more client trial balance files.")
            return False, []

        valid_paths: list[Path] = []
        for path in raw_paths:
            if not path.exists():
                self.client_card.set_status("error", f"Client trial balance file was not found: {path}")
                return False, []
            if not path.is_file():
                self.client_card.set_status("error", "Each client trial balance path must point to an Excel workbook.")
                return False, []
            if path.suffix.lower() not in EXCEL_SUFFIXES:
                self.client_card.set_status("error", "Use Excel files with .xlsx, .xlsm, or .xls extensions.")
                return False, []
            valid_paths.append(path)

        workbook_count = len(valid_paths)
        noun = "file" if workbook_count == 1 else "files"
        self.client_card.set_status("ok", f"{workbook_count} client trial balance {noun} ready.")
        return True, valid_paths

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
        client_ok, client_paths = self._validate_client_inputs()
        import_ok, import_paths = self._validate_prior_inputs()
        output_ok, _ = self._validate_output_dir()

        if client_ok and import_ok and client_paths and import_paths:
            for client_path in client_paths:
                for prior_path in import_paths:
                    try:
                        if client_path.resolve() == prior_path.resolve():
                            self.import_card.set_status(
                                "error",
                                "Choose different files here so the current-year workbooks and prior-year workbooks are not the same file.",
                            )
                            import_ok = False
                            break
                    except OSError:
                        continue

        valid = client_ok and import_ok and output_ok
        self.run_button.setEnabled(valid and not self.busy)

        if update_banner:
            if self.busy:
                tone = "ready"
                message = "Reconciliation is running. Follow the progress log below."
            elif valid:
                tone = "good"
                message = "Everything looks ready. Start the build whenever you are ready."
            else:
                tone = "ready"
                message = "Complete the current-year workbook inputs and output folder to enable reconciliation."

            self.status_banner.setProperty("tone", tone)
            self.status_banner.setText(message)
            repolish(self.status_banner)
        return valid

    def _build_run_config(self) -> RunConfig:
        client_specs = [
            WorkbookSpec(path=Path(text).expanduser())
            for text in self.client_card.path_texts()
        ]
        prior_specs = [
            WorkbookSpec(path=Path(text).expanduser())
            for text in self.import_card.path_texts()
        ]
        return RunConfig(
            client_specs=client_specs,
            prior_specs=prior_specs,
            out_dir=Path(self.out_dir_card.path_text()).expanduser(),
            write_mvp=self.mvp_checkbox.isChecked(),
            delete_zero_balance_rows=self.zero_rows_checkbox.isChecked(),
        )

    def _set_busy(self, busy: bool, preserve_banner: bool = False) -> None:
        self.busy = busy
        self.client_card.set_interactive(not busy)
        self.import_card.set_interactive(not busy)
        self.out_dir_card.set_interactive(not busy)
        self.mvp_checkbox.setEnabled(not busy)
        self.zero_rows_checkbox.setEnabled(not busy)
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

        self._append_log("Starting trial-balance build.")
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
        ready_for_import = bool(payload.get("ready_for_import", True))
        used_prior_comparison = bool(payload.get("used_prior_comparison", False))

        self.summary_caption.setText(
            "Import-ready rows: "
            f"{format_count(counts['output_rows'])}\n"
            "Prior-year rows parsed: "
            f"{format_count(counts['prior_rows'])}\n"
            "Matched: "
            f"{format_count(counts['matched'])}   "
            "New: "
            f"{format_count(counts['new_rows'])}   "
            "Carryforward: "
            f"{format_count(counts['carryforward'])}   "
            "Renumbered: "
            f"{format_count(counts['renumbered'])}   "
            "Review Queue: "
            f"{format_count(counts['review_queue'])}\n"
            "Confidence: "
            f"{format_count(counts['high_confidence'])} high   "
            f"{format_count(counts['medium_confidence'])} medium   "
            f"{format_count(counts['low_confidence'])} low\n"
            "Current-year total: "
            f"{counts['current_total']:,.2f}"
        )

        import_path = Path(outputs["import"])
        details_path = Path(outputs["details"]) if outputs["details"] else None
        review_path = Path(outputs["review"]) if outputs.get("review") else None
        folder_path = Path(outputs["folder"])
        self.latest_output_dir = folder_path
        self.open_output_button.setEnabled(True)

        lines = [
            f"Import-ready workbook: {make_local_link(import_path)}",
            f"Output folder: {make_local_link(folder_path)}",
        ]
        if details_path is not None:
            lines.insert(1, f"Audit details workbook: {make_local_link(details_path)}")
        else:
            lines.insert(1, "Audit details workbook: skipped for this run.")
        if review_path is not None:
            lines.insert(2, f"Review workbook: {make_local_link(review_path)}")
        else:
            lines.insert(2, "Review workbook: skipped for this run.")
        self.outputs_label.setText("<br>".join(lines))

        self.progress_bar.setValue(100)
        self.status_banner.setProperty("tone", "good" if ready_for_import else "ready")
        if ready_for_import and used_prior_comparison:
            message = "Reconciliation finished successfully."
        elif ready_for_import:
            message = "Build finished successfully from the current-year files."
        else:
            message = "Run completed, but the review queue still needs manual attention before import."
        self.status_banner.setText(message)
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
        self.import_card.set_recent_paths(self.settings_store.recent("recent_prior_files"))
        self.out_dir_card.set_recent_paths(self.settings_store.recent("recent_output_dirs"))

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.client_card.path_text():
            self.settings_store.remember_path("recent_client_files", self.client_card.path_text())
        if self.import_card.path_text():
            self.settings_store.remember_path("recent_prior_files", self.import_card.path_text())
        if self.out_dir_card.path_text():
            self.settings_store.remember_path("recent_output_dirs", self.out_dir_card.path_text())
        self.settings_store.data["write_mvp"] = self.mvp_checkbox.isChecked()
        self.settings_store.data["delete_zero_balance_rows"] = self.zero_rows_checkbox.isChecked()
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
