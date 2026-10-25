"""FrMat local desktop data workbench. Run: python data_workbench.py"""
import io
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import threading
import zlib
from datetime import datetime
from pathlib import Path

import pandas as pd
from defusedxml import ElementTree as SafeET
from docx import Document
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QObject, QThread, Signal, Slot
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFrame,
    QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox,
    QProgressBar, QPushButton, QSplitter, QTableView, QTextEdit, QVBoxLayout, QWidget,
)

APP_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "DataWorkbench"
APP_DIR.mkdir(parents=True, exist_ok=True)
WORKSPACE = APP_DIR / "workspace.sqlite3"
PREVIEW_ROWS = 500
QUALITY_SAMPLE_ROWS = 2000


def data_quality(frame):
    """Heuristic quality estimate on a reproducible sample, never a data mutation."""
    total = len(frame)
    if total == 0 or len(frame.columns) == 0:
        return None
    sample = frame.sample(min(total, QUALITY_SAMPLE_ROWS), random_state=42) if total > QUALITY_SAMPLE_ROWS else frame
    n, cols = sample.shape

    def is_blank(value):
        if value is None or value is pd.NA:
            return True
        if isinstance(value, (str, bytes)):
            return not value.strip()
        if isinstance(value, (list, dict, tuple, set)):
            return False
        try:
            return bool(pd.isna(value))
        except (TypeError, ValueError):
            return False

    missing = sample.apply(lambda column: column.map(is_blank))
    completeness = 100 * (1 - missing.to_numpy().sum() / (n * cols))
    duplicates = sample.astype(str).duplicated().sum()
    uniqueness = 100 * (1 - duplicates / n)

    numeric_columns = []
    for col in sample.columns:
        nonblank = sample[col].where(~missing[col])
        if nonblank.notna().sum() < 4:
            continue
        numeric = pd.to_numeric(nonblank, errors="coerce")
        if numeric.notna().sum() / nonblank.notna().sum() >= .8:
            numeric_columns.append(numeric)

    outlier_rows = pd.Series(False, index=sample.index)
    evaluated = False
    for numeric in numeric_columns:
        q1, q3 = numeric.quantile([.25, .75])
        iqr = q3 - q1
        if pd.notna(iqr) and iqr > 0:
            evaluated = True
            outlier_rows |= numeric.lt(q1 - 1.5 * iqr) | numeric.gt(q3 + 1.5 * iqr)

    outliers = 100 * (1 - outlier_rows.sum() / n) if evaluated else None

    text_cells = 0
    hard_to_read = 0
    for col in sample.columns:
        values = sample[col]
        if not (pd.api.types.is_object_dtype(values) or pd.api.types.is_string_dtype(values)):
            continue
        for value in values:
            if not isinstance(value, str) or not value.strip():
                continue
            text_cells += 1
            hard_to_read += bool("\ufffd" in value or len(value) > 300 or any(ord(ch) < 32 and ch not in "\n\r\t" for ch in value))

    readability = 100 * (1 - hard_to_read / text_cells) if text_cells else None

    metrics = [
        ("Complete cells", completeness, 35),
        ("Unique rows", uniqueness, 20),
        ("No IQR outliers", outliers, 25),
        ("Readable text", readability, 20),
    ]
    applicable = [(value, weight) for _, value, weight in metrics if value is not None]
    score = sum(value * weight for value, weight in applicable) / sum(weight for _, weight in applicable)

    return {
        "score": score,
        "metrics": metrics,
        "rows": n,
        "total": total,
        "duplicates": int(duplicates),
        "outlier_rows": int(outlier_rows.sum()),
        "hard_to_read": hard_to_read,
    }


def unique_columns(frame):
    frame = frame.copy()
    used = set()
    names = []
    for i, col in enumerate(frame.columns):
        base = str(col).strip() or f"column_{i + 1}"
        name, counter = base, 2
        while name.casefold() in used:
            name = f"{base}_{counter}"
            counter += 1
        used.add(name.casefold())
        names.append(name)
    frame.columns = names
    return frame.reset_index(drop=True)


def text_frame(text):
    return pd.DataFrame({"line": text.splitlines() or [""]})


def parse_xml(path):
    root = SafeET.parse(path).getroot()
    elements = list(root)
    if not elements:
        elements = [root]
    rows = []
    for item in elements:
        row = {"tag": item.tag}
        row.update({"@" + k: v for k, v in item.attrib.items()})
        for child in item:
            key = child.tag
            value = " ".join(child.itertext()).strip()
            row[key] = (row[key] + " | " + value) if key in row else value
        if not list(item):
            row["text"] = (item.text or "").strip()
        rows.append(row)
    return pd.DataFrame(rows)


def sqlite_tables(connection):
    rows = connection.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall()
    return [row[0] for row in rows]


def read_file(path, parent, pdf_mode="Detected tables", doc_mode="Detected tables"):
    suffix = path.suffix.lower()
    if suffix == ".csv":
        try:
            return pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python")
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="cp1252", sep=None, engine="python")
    if suffix in (".xls", ".xlsx"):
        book = pd.ExcelFile(path)
        sheet, ok = QInputDialog.getItem(parent, "Choose sheet", "Worksheet:", book.sheet_names, 0, False)
        if not ok:
            return None
        return pd.read_excel(book, sheet_name=sheet)
    if suffix == ".txt":
        try:
            text = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            text = path.read_text(encoding="cp1252")
        first = "\n".join(text.splitlines()[:10])
        if first and any(c in first for c in ("\t", ",", ";", "|")):
            try:
                return pd.read_csv(io.StringIO(text), sep=None, engine="python")
            except (ValueError, pd.errors.ParserError):
                pass
        return text_frame(text)
    if suffix == ".json":
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(data, list):
            return pd.json_normalize(data) if data else pd.DataFrame()
        if isinstance(data, dict):
            for value in data.values():
                if isinstance(value, list) and value and isinstance(value[0], dict):
                    return pd.json_normalize(value)
            return pd.json_normalize(data)
        return pd.DataFrame({"value": [data]})
    if suffix == ".xml":
        return parse_xml(path)
    if suffix == ".pdf":
        import pdfplumber
        tables, lines = [], []
        with pdfplumber.open(path) as pdf:
            for page_no, page in enumerate(pdf.pages, 1):
                for table_no, table in enumerate(page.extract_tables() or [], 1):
                    if not table:
                        continue
                    header = [str(c).strip() if c is not None else "" for c in table[0]]
                    used = set()
                    for i, value in enumerate(header):
                        base = value or f"column_{i + 1}"
                        name, count = base, 2
                        while name.casefold() in used or name in ("page", "table"):
                            name, count = f"{base}_{count}", count + 1
                        used.add(name.casefold())
                        header[i] = name
                    if not any(c is not None and str(c).strip() for c in table[0]):
                        header = [f"column_{i + 1}" for i in range(len(table[0]))]
                        records = table
                    else:
                        records = table[1:]
                    for cells in records:
                        if any(c is not None and str(c).strip() for c in cells):
                            row = {name or f"column_{i + 1}": value for i, (name, value) in enumerate(zip(header, cells))}
                            tables.append({"page": page_no, "table": table_no, **row})
                for line in (page.extract_text() or "").splitlines():
                    lines.append({"page": page_no, "line": line})
        if pdf_mode == "Text lines" or not tables:
            return pd.DataFrame(lines, columns=["page", "line"])
        return pd.DataFrame(tables)
    if suffix in (".doc", ".docx"):
        if suffix == ".doc":
            with tempfile.TemporaryDirectory() as temp:
                result = subprocess.run(
                    ["soffice", "-env:UserInstallation=file:///" + temp.replace("\\", "/") + "/lo_profile", "--headless", "--convert-to", "docx", "--outdir", temp, str(path)],
                    capture_output=True, text=True, timeout=120,
                )
                converted = Path(temp) / (path.stem + ".docx")
                if result.returncode != 0 or not converted.exists():
                    raise ValueError(".doc requires LibreOffice (soffice) installed and available on PATH")
                return read_docx(converted, doc_mode)
        return read_docx(path, doc_mode)
    if suffix in (".db", ".sqlite", ".sqlite3", ".sql"):
        if suffix == ".sql":
            answer = QMessageBox.question(parent, "Run SQL dump?", "Only import SQL dumps you trust. The script runs in a temporary in-memory SQLite database. Continue?")
            if answer != QMessageBox.StandardButton.Yes:
                return None
            connection = sqlite3.connect(":memory:")
            forbidden = {sqlite3.SQLITE_ATTACH, sqlite3.SQLITE_DETACH, sqlite3.SQLITE_PRAGMA}
            if hasattr(sqlite3, "SQLITE_VACUUM"):
                forbidden.add(sqlite3.SQLITE_VACUUM)
            connection.set_authorizer(lambda action, *_: sqlite3.SQLITE_DENY if action in forbidden else sqlite3.SQLITE_OK)
            try:
                connection.executescript(path.read_text(encoding="utf-8-sig"))
                return choose_table(connection, parent)
            finally:
                connection.close()
        connection = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)
        try:
            return choose_table(connection, parent)
        finally:
            connection.close()
    raise ValueError(f"Unsupported extension: {suffix}")


def read_docx(path, mode="Detected tables"):
    document = Document(path)
    tables = document.tables
    if tables and mode == "Detected tables":
        rows = []
        for number, table in enumerate(tables, 1):
            if not table.rows:
                continue
            headers = [cell.text.strip() or f"column_{i + 1}" for i, cell in enumerate(table.rows[0].cells)]
            for source_row in table.rows[1:]:
                rows.append({"table": number, **{h: cell.text for h, cell in zip(headers, source_row.cells)}})
        if rows:
            return pd.DataFrame(rows)
    rows = [{"type": "paragraph", "text": p.text} for p in document.paragraphs if p.text.strip()]
    for number, table in enumerate(tables, 1):
        rows.extend({"type": f"table {number}", "text": " | ".join(cell.text for cell in row.cells)} for row in table.rows)
    return pd.DataFrame(rows, columns=["type", "text"])


def choose_table(connection, parent):
    tables = sqlite_tables(connection)
    if not tables:
        raise ValueError("No data tables found in the SQLite database")
    table, ok = QInputDialog.getItem(parent, "Choose table", "Table:", tables, 0, False)
    if not ok:
        return None
    return pd.read_sql_query('SELECT * FROM "' + table.replace('"', '""') + '"', connection)


def serialize(frame):
    return zlib.compress(frame.to_json(orient="split", date_format="iso", default_handler=str).encode("utf-8"))


def deserialize(blob):
    return pd.read_json(io.StringIO(zlib.decompress(blob).decode("utf-8")), orient="split")


def apply_conditions(frame, expression):
    if not expression.strip():
        raise ValueError("Enter a condition, such as [age] >= 18 AND [region] = West")
    mask = pd.Series(True, index=frame.index)
    for condition in re.split(r"\s+AND\s+", expression.strip(), flags=re.IGNORECASE):
        match = re.fullmatch(r"\s*\[([^]]+)\]\s*(>=|<=|!=|=|>|<|~)\s*(.*?)\s*", condition)
        if not match:
            raise ValueError("Use [column] operator value, joined by AND. Operators: = != > >= < <= ~")
        column, operator, value = match.groups()
        if column not in frame.columns:
            raise ValueError(f"Unknown column: {column}")
        series = frame[column]
        if operator in ("=", "!=") and value.lower() == "null":
            result = series.isna() if operator == "=" else series.notna()
        elif operator == "~":
            result = series.astype("string").str.contains(value, case=False, regex=False, na=False)
        elif operator in ("=", "!="):
            result = series.astype("string").str.casefold() == value.casefold()
            result = result.fillna(False)
            if operator == "!=":
                result = (~result) & series.notna()
        else:
            number = pd.to_numeric(value, errors="coerce")
            if pd.isna(number):
                raise ValueError("Comparison values for >, >=, <, <= must be numbers")
            numeric = pd.to_numeric(series, errors="coerce")
            result = {">": numeric.gt, ">=": numeric.ge, "<": numeric.lt, "<=": numeric.le}[operator](number).fillna(False)
        mask &= result.fillna(False)
    return frame.loc[mask].reset_index(drop=True)


def chart_eligibility(frame, kind, x, y):
    """Return row-level eligibility for a chart and why each excluded row fails."""
    x_values = frame[x]
    if kind == 1:
        return pd.Series(True, index=frame.index), "Any category, including missing, is usable"
    if kind == 0:
        x_ok = pd.to_numeric(x_values, errors="coerce").notna()
        return x_ok, "X must be numeric; missing or nonnumeric X is excluded"
    y_ok = pd.to_numeric(frame[y], errors="coerce").notna()
    if kind == 2:
        x_ok = x_values.notna() & x_values.astype("string").str.strip().ne("").fillna(False)
        rule = "X must be nonempty; Y must be numeric"
    else:
        x_ok = pd.to_numeric(x_values, errors="coerce").notna()
        rule = "Both X and Y must be numeric"
    return x_ok & y_ok, rule


def preset_result(frame, choice, columns):
    if not columns:
        raise ValueError("Select at least one column for this preset")
    subset = frame[columns]
    if choice == 1:
        keep = ~subset.duplicated()
    elif choice == 2:
        keep = ~subset.isna().any(axis=1)
    elif choice == 3:
        numeric = subset.apply(pd.to_numeric, errors="coerce")
        valid = [col for col in columns if numeric[col].notna().sum() >= 4]
        if not valid:
            raise ValueError("Select numeric columns with at least four values for IQR cleaning")
        numeric = numeric[valid]
        q1, q3 = numeric.quantile(.25), numeric.quantile(.75)
        iqr = q3 - q1
        keep = ~((numeric.lt(q1 - 1.5 * iqr) | numeric.gt(q3 + 1.5 * iqr)) & iqr.gt(0)).any(axis=1)
    else:
        raise ValueError("Choose a cleaning preset")
    return frame.loc[keep].reset_index(drop=True), frame.loc[~keep].reset_index(drop=True)


class TaskWorker(QObject):
    finished = Signal(object, object)

    def __init__(self, operation, cancel_event, discard_result=True):
        super().__init__()
        self.operation = operation
        self.cancel_event = cancel_event
        self.discard_result = discard_result

    @Slot()
    def run(self):
        try:
            result = self.operation()
            if self.cancel_event.is_set() and self.discard_result:
                self.finished.emit(None, "Operation cancelled; no changes saved")
            else:
                self.finished.emit(result, None)
        except Exception as exc:
            self.finished.emit(None, str(exc))


class FrameModel(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.frame = pd.DataFrame()
        self.eligibility = None
        self.x_column = None
        self.y_column = None

    def set_frame(self, frame):
        self.beginResetModel()
        self.frame = frame.head(PREVIEW_ROWS)
        self.eligibility = None
        self.endResetModel()

    def set_eligibility(self, eligibility, x, y):
        self.eligibility = eligibility.head(PREVIEW_ROWS) if eligibility is not None else None
        self.x_column, self.y_column = x, y
        if not self.frame.empty and len(self.frame.columns):
            first = self.index(0, 0)
            last = self.index(len(self.frame) - 1, len(self.frame.columns) - 1)
            self.dataChanged.emit(first, last, [Qt.ItemDataRole.BackgroundRole, Qt.ItemDataRole.ToolTipRole])

    def rowCount(self, parent=QModelIndex()):
        return len(self.frame)

    def columnCount(self, parent=QModelIndex()):
        return len(self.frame.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            value = self.frame.iat[index.row(), index.column()]
            try:
                missing = pd.isna(value)
                return "" if isinstance(missing, (bool, int)) and missing else str(value)
            except (TypeError, ValueError):
                return str(value)
        if self.eligibility is not None:
            if role == Qt.ItemDataRole.BackgroundRole:
                return QColor("#19483f" if bool(self.eligibility.iloc[index.row()]) else "#633039")
            if role == Qt.ItemDataRole.ToolTipRole:
                return "Usable for selected chart" if bool(self.eligibility.iloc[index.row()]) else "Excluded from selected chart"
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return str(self.frame.columns[section]) if orientation == Qt.Orientation.Horizontal else str(section + 1)
        return None


class Workbench(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FrMat")
        self.resize(1250, 800)
        self.connection = sqlite3.connect(WORKSPACE)
        self.connection.execute("CREATE TABLE IF NOT EXISTS datasets (id INTEGER PRIMARY KEY, name TEXT NOT NULL, source TEXT NOT NULL, imported_at TEXT NOT NULL, original BLOB NOT NULL, current BLOB NOT NULL)")
        self.connection.execute("CREATE TABLE IF NOT EXISTS history (id INTEGER PRIMARY KEY, dataset_id INTEGER NOT NULL, created_at TEXT NOT NULL, action TEXT NOT NULL, rows_before INTEGER NOT NULL, rows_after INTEGER NOT NULL, snapshot BLOB NOT NULL)")
        self.connection.commit()
        self.dataset_id = None
        self.original = pd.DataFrame()
        self.current = pd.DataFrame()
        self.model = FrameModel()
        self.busy = False
        self.task_thread = None
        self.task_worker = None
        self.build_ui()
        self.reload_datasets()

    def build_ui(self):
        main = QWidget()
        layout = QVBoxLayout(main)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        header = QHBoxLayout()
        heading = QVBoxLayout()
        title = QLabel("FrMat")
        title.setObjectName("appTitle")
        heading.addWidget(title)
        subtitle = QLabel("Import · clean · explore · export — locally on your PC")
        subtitle.setObjectName("subtitle")
        heading.addWidget(subtitle)
        header.addLayout(heading, 1)

        for label, handler, style in (("Import file", self.import_file, "primaryButton"), ("Export data", self.export_file, "successButton"), ("Reset to import", self.reset_data, "")):
            button = QPushButton(label)
            button.setObjectName(style)
            button.clicked.connect(handler)
            header.addWidget(button)

        layout.addLayout(header)

        def panel(title_text):
            card = QFrame()
            card.setObjectName("panel")
            body = QVBoxLayout(card)
            body.setContentsMargins(16, 12, 16, 14)
            body.setSpacing(9)
            label = QLabel(title_text)
            label.setObjectName("sectionTitle")
            body.addWidget(label)
            return card, body

        workspace, body = panel("WORKSPACE")
        top = QHBoxLayout()
        top.addWidget(QLabel("Dataset"))
        self.datasets = QComboBox()
        self.datasets.currentIndexChanged.connect(self.load_dataset)
        top.addWidget(self.datasets, 1)
        body.addLayout(top)

        quality_row = QHBoxLayout()
        quality_row.addWidget(QLabel("Data quality estimate"))
        self.quality_score = QLabel("N/A")
        self.quality_score.setToolTip("Heuristic estimate for the current dataset, not a certification of correctness")
        quality_row.addWidget(self.quality_score)
        quality_row.addStretch(1)
        body.addLayout(quality_row)

        self.quality_bar = QProgressBar()
        self.quality_bar.setRange(0, 100)
        self.quality_bar.setValue(0)
        self.quality_bar.setTextVisible(False)
        body.addWidget(self.quality_bar)

        self.quality_details = QLabel("Import data to see completeness, duplicates, outliers, and readability.")
        self.quality_details.setObjectName("guidance")
        self.quality_details.setWordWrap(True)
        body.addWidget(self.quality_details)

        layout.addWidget(workspace)

        processing, body = panel("FILTER & CLEAN")
        filters = QHBoxLayout()
        self.prompt = QLineEdit()
        self.prompt.setPlaceholderText("Filter example: [age] >= 18 AND [city] ~ York")
        filters.addWidget(self.prompt, 1)

        button = QPushButton("Apply filter")
        button.setObjectName("primaryButton")
        button.clicked.connect(self.filter_data)
        filters.addWidget(button)
        body.addLayout(filters)

        presets = QHBoxLayout()
        self.presets = QComboBox()
        self.presets.addItems(["Choose preset", "Remove duplicates", "Drop rows with missing values", "Remove IQR outliers"])
        presets.addWidget(self.presets, 1)

        button = QPushButton("Apply preset")
        button.clicked.connect(self.apply_preset)
        presets.addWidget(button)
        body.addLayout(presets)

        actions = QHBoxLayout()
        undo = QPushButton("Undo last step")
        undo.clicked.connect(self.undo_data)
        actions.addWidget(undo)

        history = QPushButton("Processing history")
        history.clicked.connect(self.show_history)
        actions.addWidget(history)
        actions.addStretch(1)
        body.addLayout(actions)

        layout.addWidget(processing)

        visualization, body = panel("VISUALIZE")
        chart = QHBoxLayout()
        self.chart_type = QComboBox()
        self.chart_type.addItems(["Histogram", "Bar (count by category)", "Line", "Scatter"])
        self.x_axis = QComboBox()
        self.y_axis = QComboBox()

        for label, widget in (("Chart", self.chart_type), ("X / value", self.x_axis), ("Y", self.y_axis)):
            chart.addWidget(QLabel(label))
            chart.addWidget(widget, 1)

        self.selected_only = QCheckBox("Selected preview rows only")
        self.selected_only.toggled.connect(self.clear_chart_check)
        chart.addWidget(self.selected_only)
        body.addLayout(chart)

        for picker in (self.chart_type, self.x_axis, self.y_axis):
            picker.currentIndexChanged.connect(self.clear_chart_check)

        actions = QHBoxLayout()
        for label, handler, style in (("Check X/Y suitability", self.check_chart_data, ""), ("Draw chart", self.draw_chart, "primaryButton"), ("Save chart PNG", self.save_chart, "")):
            button = QPushButton(label)
            button.setObjectName(style)
            button.clicked.connect(handler)
            actions.addWidget(button)

        actions.addStretch(1)
        body.addLayout(actions)

        self.chart_guidance = QLabel("Choose a chart and axes, then check which rows can be plotted.")
        self.chart_guidance.setObjectName("guidance")
        self.chart_guidance.setWordWrap(True)
        body.addWidget(self.chart_guidance)

        layout.addWidget(visualization)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        self.table.selectionModel().selectionChanged.connect(self.clear_chart_check)
        splitter.addWidget(self.table)

        self.figure = Figure(figsize=(9, 4), facecolor="#101a2a")
        self.canvas = FigureCanvasQTAgg(self.figure)
        splitter.addWidget(self.canvas)
        splitter.setSizes([460, 230])
        layout.addWidget(splitter, 1)

        self.status = QLabel("Import a file to begin. Data stays on this PC.")
        self.status.setObjectName("status")
        layout.addWidget(self.status)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.cancel_button = QPushButton("Cancel after current step")
        self.cancel_button.clicked.connect(self.cancel_task)
        self.cancel_button.hide()
        layout.addWidget(self.cancel_button)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(60)
        layout.addWidget(self.log)

        self.setCentralWidget(main)

    def error(self, exc):
        QMessageBox.critical(self, "FrMat", str(exc))

    def reload_datasets(self, select_id=None):
        rows = self.connection.execute("SELECT id, name, imported_at FROM datasets ORDER BY id DESC").fetchall()
        self.datasets.blockSignals(True)
        self.datasets.clear()
        for dataset_id, name, date in rows:
            self.datasets.addItem(f"{name} — {date}", dataset_id)
        if select_id is not None:
            index = self.datasets.findData(select_id)
            if index >= 0:
                self.datasets.setCurrentIndex(index)
        self.datasets.blockSignals(False)
        self.load_dataset()

    def load_dataset(self, *_):
        dataset_id = self.datasets.currentData()
        if dataset_id is None:
            return
        row = self.connection.execute("SELECT original, current FROM datasets WHERE id = ?", (dataset_id,)).fetchone()
        if row:
            self.dataset_id = dataset_id
            self.original, self.current = deserialize(row[0]), deserialize(row[1])
            self.refresh("Loaded saved dataset")

    def update_quality(self):
        try:
            result = data_quality(self.current)
        except Exception as exc:
            self.quality_score.setText("Unavailable")
            self.quality_bar.setValue(0)
            self.quality_details.setText(f"Quality estimate could not be calculated: {exc}")
            return

        if result is None:
            self.quality_score.setText("N/A")
            self.quality_bar.setValue(0)
            self.quality_details.setText("No rows to assess; an empty dataset has no quality percentage.")
            return

        score = result["score"]
        self.quality_score.setText(f"{score:.0f}%")
        self.quality_bar.setValue(round(score))
        labels = [f"{name}: {value:.0f}%" if value is not None else f"{name}: N/A" for name, value, _ in result["metrics"]]
        scope = f"sample of {result['rows']:,} / {result['total']:,} rows" if result["rows"] < result["total"] else f"all {result['total']:,} rows"
        self.quality_details.setText(f"Based on {scope} · " + "  |  ".join(labels) + ". Higher is better; outliers can be legitimate. This is an estimate, not a validation.")

    def refresh(self, message):
        self.model.set_frame(self.current)
        self.update_quality()
        self.chart_guidance.setText("Choose a chart and axes, then check which rows can be plotted.")
        previous_x, previous_y = self.x_axis.currentText(), self.y_axis.currentText()
        self.x_axis.blockSignals(True)
        self.y_axis.blockSignals(True)
        self.x_axis.clear()
        self.y_axis.clear()
        columns = list(self.current.columns)
        self.x_axis.addItems(columns)
        self.y_axis.addItems(columns)
        if previous_x in columns:
            self.x_axis.setCurrentText(previous_x)
        if previous_y in columns:
            self.y_axis.setCurrentText(previous_y)
        self.x_axis.blockSignals(False)
        self.y_axis.blockSignals(False)
        self.status.setText(f"{message} | {len(self.current):,} rows × {len(columns)} columns | showing first {PREVIEW_ROWS} rows")
        self.log.append(f"{datetime.now():%H:%M:%S}  {message}")

    def persist(self, result, message):
        before, after = len(self.current), len(result)
        with self.connection:
            self.connection.execute("INSERT INTO history (dataset_id, created_at, action, rows_before, rows_after, snapshot) VALUES (?, ?, ?, ?, ?, ?)", (self.dataset_id, datetime.now().isoformat(timespec="seconds"), message, before, after, serialize(self.current)))
            self.connection.execute("UPDATE datasets SET current = ? WHERE id = ?", (serialize(result), self.dataset_id))
        self.current = result
        self.refresh(message)

    def undo_data(self):
        if self.dataset_id is None or self.busy:
            return
        row = self.connection.execute("SELECT id, action, snapshot FROM history WHERE dataset_id = ? ORDER BY id DESC LIMIT 1", (self.dataset_id,)).fetchone()
        if not row:
            self.status.setText("Nothing to undo for this dataset")
            return
        restored = deserialize(row[2])
        with self.connection:
            self.connection.execute("UPDATE datasets SET current = ? WHERE id = ?", (row[2], self.dataset_id))
            self.connection.execute("DELETE FROM history WHERE id = ?", (row[0],))
        self.current = restored
        self.refresh(f"Undid: {row[1]}")

    def show_history(self):
        if self.dataset_id is None:
            return
        rows = self.connection.execute("SELECT created_at, action, rows_before, rows_after FROM history WHERE dataset_id = ? ORDER BY id DESC LIMIT 100", (self.dataset_id,)).fetchall()
        QMessageBox.information(self, "Processing history (latest 100)", "\n".join(f"{date}: {action} ({before:,} → {after:,} rows)" for date, action, before, after in rows) or "No processing steps yet.")

    @Slot(object, object)
    def receive_task_result(self, result, error):
        self.task_result = (result, error)

    def cancel_task(self):
        self.cancel_event.set()
        self.cancel_button.setEnabled(False)
        self.status.setText("Cancellation requested; waiting for current step…")

    def start_task(self, operation, completed, label, discard_result=True):
        if self.busy:
            return
        self.busy = True
        self.datasets.setEnabled(False)
        self.progress.show()
        self.cancel_button.setEnabled(True)
        self.cancel_button.show()
        self.cancel_event = threading.Event()
        self.task_thread = QThread(self)
        self.task_worker = TaskWorker(operation, self.cancel_event, discard_result)
        self.task_worker.moveToThread(self.task_thread)
        self.task_thread.started.connect(self.task_worker.run)
        self.task_worker.finished.connect(self.receive_task_result)
        self.task_worker.finished.connect(self.task_thread.quit)
        self.task_worker.finished.connect(self.task_worker.deleteLater)
        self.task_result = (None, "Task ended without result")

        self.task_thread.finished.connect(lambda: self._task_finished(completed))
        self.task_thread.start()
        self.status.setText(label)

    def _task_finished(self, completed):
        try:
            self.busy = False
            self.datasets.setEnabled(True)
            self.progress.hide()
            self.cancel_button.hide()
            result, error = self.task_result
            self.task_thread.deleteLater()
            self.task_thread = None
            self.task_worker = None
            if error:
                if "cancel" not in str(error).lower():
                    self.error(error)
                self.status.setText(str(error) if "cancel" in str(error).lower() else f"Failed: {completed.__name__}")
            else:
                try:
                    completed(result)
                except Exception as exc:
                    self.error(exc)
        except Exception as exc:
            self.error(exc)

    def import_file(self):
        if self.busy:
            return
        filename, _ = QFileDialog.getOpenFileName(self, "Import data", "", "Data (*.csv *.xls *.xlsx *.txt *.pdf *.doc *.docx *.json *.xml *.sql *.db *.sqlite *.sqlite3)")
        if not filename:
            return
        path = Path(filename)

        def completed(frame):
            if frame is None:
                return
            frame = unique_columns(frame)
            if not len(frame.columns):
                raise ValueError("No columns could be extracted from this file")
            if QMessageBox.question(self, "Import preview", f"{path.name}: {len(frame):,} rows × {len(frame.columns)} columns\n\nSave this import?") != QMessageBox.StandardButton.Yes:
                return
            payload = serialize(frame)
            cursor = self.connection.execute("INSERT INTO datasets (name, source, imported_at, original, current) VALUES (?, ?, ?, ?, ?)", (path.name, filename, datetime.now().isoformat(timespec="seconds"), payload, payload))
            self.connection.commit()
            self.reload_datasets(cursor.lastrowid)
            self.refresh(f"Imported {path.name}")

        if path.suffix.lower() in (".xls", ".xlsx", ".db", ".sqlite", ".sqlite3", ".sql"):
            try:
                completed(read_file(path, self))
            except Exception as exc:
                self.error(exc)
        else:
            pdf_mode, doc_mode = "Detected tables", "Detected tables"
            if path.suffix.lower() == ".pdf":
                pdf_mode, ok = QInputDialog.getItem(self, "PDF extraction", "Import:", ["Detected tables", "Text lines"], 0, False)
                if not ok:
                    return
            if path.suffix.lower() in (".doc", ".docx"):
                doc_mode, ok = QInputDialog.getItem(self, "Word extraction", "Import:", ["Detected tables", "Paragraphs and table text"], 0, False)
                if not ok:
                    return
            self.start_task(lambda: read_file(path, None, pdf_mode, doc_mode), completed, "Importing file…")

    def reset_data(self):
        if self.dataset_id is not None and not self.busy:
            self.current = self.original.copy()
            self.persist(self.current, "Restored original imported data")

    def filter_data(self):
        if self.dataset_id is None or self.busy:
            return
        frame = self.current

        def completed(result):
            self.persist(result, f"Filter {self.prompt.text()}; removed {len(frame) - len(result):,} rows")

        self.start_task(lambda: apply_conditions(frame, self.prompt.text()), completed, "Applying filter…")

    def apply_preset(self):
        if self.dataset_id is None or self.busy or self.presets.currentIndex() == 0:
            return
        choice, frame = self.presets.currentIndex(), self.current
        picker = QDialog(self)
        picker.setWindowTitle("Select columns to clean")
        body = QVBoxLayout(picker)
        body.addWidget(QLabel("Choose columns. Duplicate detection compares the selected columns only."))
        columns_list = QListWidget()
        for col in frame.columns:
            item = QListWidgetItem(col)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if choice != 3 or pd.to_numeric(frame[col], errors="coerce").notna().sum() >= 4 else Qt.CheckState.Unchecked)
            columns_list.addItem(item)
        body.addWidget(columns_list)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(picker.accept)
        buttons.rejected.connect(picker.reject)
        body.addWidget(buttons)

        if picker.exec() != QDialog.DialogCode.Accepted:
            return

        columns = [columns_list.item(i).text() for i in range(columns_list.count()) if columns_list.item(i).checkState() == Qt.CheckState.Checked]
        name = self.presets.currentText()

        def preview(pair):
            result, removed = pair
            sample = removed.head(8).to_string(index=False, max_colwidth=24) if not removed.empty else "None"
            text = f"{name} on {', '.join(columns)}\n\nRemove {len(removed):,} of {len(frame):,} rows?\n\nFirst affected rows:\n{sample}"
            if QMessageBox.question(self, "Confirm cleaning", text) == QMessageBox.StandardButton.Yes:
                self.persist(result, f"{name} [{', '.join(columns)}]; removed {len(removed):,} rows")

        self.start_task(lambda: preset_result(frame, choice, columns), preview, "Previewing cleaning…")

    def clear_chart_check(self, *_):
        self.chart_guidance.setText("Chart scope or axes changed; click Check X/Y suitability to update row colors.")
        self.model.set_eligibility(None, None, None)

    def chart_source(self):
        if not self.selected_only.isChecked():
            return self.current, "all current rows"
        indexes = sorted(index.row() for index in self.table.selectionModel().selectedRows())
        if not indexes:
            raise ValueError("Select rows in the first 500-row preview, or uncheck Selected preview rows only")
        return self.model.frame.iloc[indexes], "selected preview rows (first 500 only)"

    def check_chart_data(self):
        if self.dataset_id is None or self.busy:
            return
        try:
            kind = self.chart_type.currentIndex()
            x, y = self.x_axis.currentText(), self.y_axis.currentText()
            source, scope = self.chart_source()
            eligible, rule = chart_eligibility(source, kind, x, y)
            preview_colors = pd.Series(False, index=self.model.frame.index)
            if self.selected_only.isChecked():
                preview_colors.loc[source.index] = eligible
            else:
                preview_colors = eligible.head(PREVIEW_ROWS)
            self.model.set_eligibility(preview_colors, x, y)
            note = " (Y not used)" if kind in (0, 1) else ""
            self.chart_guidance.setText(f"Scope: {scope}. X: {x}; Y: {y}{note}. {int(eligible.sum()):,} of {len(eligible):,} scoped rows usable. Green = usable; red = excluded or outside selection. {rule}.")
        except Exception as exc:
            self.error(exc)

    def draw_chart(self):
        if self.dataset_id is None or self.current.empty or self.busy:
            return
        try:
            x, y = self.x_axis.currentText(), self.y_axis.currentText()
            kind = self.chart_type.currentIndex()
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            if kind == 0:
                values = pd.to_numeric(self.current[x], errors="coerce").dropna()
                if values.empty:
                    raise ValueError("Histogram requires numeric X values")
                ax.hist(values, bins=30)
                ax.set_xlabel(x)
            elif kind == 1:
                counts = self.current[x].astype("string").fillna("(missing)").value_counts().head(30)
                counts.plot.bar(ax=ax)
                ax.set_ylabel("Count (top 30 categories)")
                ax.set_xlabel(x)
            else:
                values = pd.to_numeric(self.current[y], errors="coerce")
                if not values.notna().any():
                    raise ValueError("Line and scatter charts require numeric Y values")
                if kind == 2:
                    ax.plot(self.current[x].astype(str), values)
                    if len(self.current) > 30:
                        ax.xaxis.set_major_locator(__import__("matplotlib").ticker.MaxNLocator(12))
                else:
                    xvalues = pd.to_numeric(self.current[x], errors="coerce")
                    if not xvalues.notna().any():
                        raise ValueError("Scatter requires numeric X values")
                    ax.scatter(xvalues, values, s=12)
                ax.set_xlabel(x)
                ax.set_ylabel(y)
            self.figure.tight_layout()
            self.canvas.draw()
            self.status.setText("Chart drawn from all current rows. Use Save chart PNG to export it.")
        except Exception as exc:
            self.error(exc)

    def save_chart(self):
        if not self.figure.axes:
            self.error(ValueError("Draw a chart first"))
            return
        filename, _ = QFileDialog.getSaveFileName(self, "Save chart", "chart.png", "PNG image (*.png)")
        if filename:
            self.figure.savefig(filename, dpi=150, bbox_inches="tight")

    def export_file(self):
        if self.dataset_id is None or self.busy:
            return
        filename, _ = QFileDialog.getSaveFileName(self, "Export processed data", "processed.csv", "CSV (*.csv);;Excel (*.xlsx);;JSON (*.json);;Text (*.txt);;XML (*.xml);;PDF (*.pdf);;Word (*.docx);;SQLite (*.db);;SQL dump (*.sql)")
        if not filename:
            return
        path = Path(filename)
        if path.exists() and QMessageBox.question(self, "Replace file?", f"Replace existing file?\n{path}") != QMessageBox.StandardButton.Yes:
            return

        def write_export():
            suffix = path.suffix.lower()
            if suffix not in (".csv", ".xlsx", ".json", ".txt", ".xml", ".pdf", ".docx", ".db", ".sql"):
                raise ValueError("Choose a supported output extension")
            fd, temporary = tempfile.mkstemp(prefix=".dataworkbench_", suffix=suffix, dir=path.parent)
            os.close(fd)
            target = Path(temporary)
            try:
                self.write_format(self.current, target, suffix)
                os.replace(target, path)
            finally:
                if target.exists():
                    target.unlink(missing_ok=True)
            return len(self.current)

        def completed(count):
            self.log.append(f"Exported {count:,} rows to {path}")
            self.status.setText(f"Exported {count:,} rows to {path}")

        self.start_task(write_export, completed, "Exporting data…", discard_result=False)

    @staticmethod
    def write_format(frame, path, suffix):
        if suffix == ".csv":
            frame.to_csv(path, index=False, encoding="utf-8-sig")
        elif suffix == ".xlsx":
            frame.to_excel(path, index=False)
        elif suffix == ".json":
            frame.to_json(path, orient="records", force_ascii=False, indent=2, date_format="iso")
        elif suffix == ".txt":
            frame.to_csv(path, index=False, sep="\t", encoding="utf-8")
        elif suffix == ".xml":
            frame.to_xml(path, index=False, root_name="data", row_name="row")
        elif suffix == ".db":
            with sqlite3.connect(path) as database:
                frame.to_sql("data", database, index=False, if_exists="replace")
        elif suffix == ".sql":
            with tempfile.TemporaryDirectory() as temp:
                with sqlite3.connect(Path(temp) / "export.db") as database:
                    frame.to_sql("data", database, index=False, if_exists="replace")
                    path.write_text("\n".join(database.iterdump()) + "\n", encoding="utf-8")
        elif suffix == ".docx":
            doc = Document()
            doc.add_heading("Processed data", 0)
            table = doc.add_table(rows=1, cols=len(frame.columns))
            for cell, col in zip(table.rows[0].cells, frame.columns):
                cell.text = str(col)
            for row in frame.itertuples(index=False, name=None):
                for cell, value in zip(table.add_row().cells, row):
                    cell.text = "" if pd.isna(value) else str(value)
            doc.save(path)
        elif suffix == ".pdf":
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import landscape, letter
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
            from reportlab.lib.utils import simpleSplit
            if len(frame.columns) > 12:
                raise ValueError("PDF export supports up to 12 columns; use CSV/XLSX for wide data")
            doc = SimpleDocTemplate(str(path), pagesize=landscape(letter))
            width = (landscape(letter)[0] - 72) / len(frame.columns)
            data = [list(map(str, frame.columns))]
            for row in frame.itertuples(index=False, name=None):
                data.append(["\n".join(simpleSplit("" if pd.isna(v) else str(v), "Helvetica", 7, width - 8)[:3]) for v in row])
            table = Table(data, colWidths=[width] * len(frame.columns), repeatRows=1)
            table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), .25, colors.grey), ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("FONTSIZE", (0, 0), (-1, -1), 7)]))
            doc.build([table])
        else:
            raise ValueError("Choose a supported output extension")

    def closeEvent(self, event):
        if self.busy:
            QMessageBox.information(self, "Operation in progress", "Please wait for the current operation to finish before closing.")
            event.ignore()
            return
        self.connection.close()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Workbench()
    window.show()
    sys.exit(app.exec())
