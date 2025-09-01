#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, os, json, re
from pathlib import Path
import pandas as pd

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox, QWidget, QTabWidget,
    QToolBar, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QCheckBox, QTableView, QStatusBar, QMenu
)

APP_NAME = "Cinteza CSV Exporter"
APP_VER  = "1.0"

# Schema actualizată: fără "Details 4"
COLUMNS = [
    "PO_No.", "Amount", "CCY/RON", "Payer Account IBAN",
    "Payee Name", "Payee address 1", "Payee address 2",
    "Payee CUI", "Payee Account IBAN",
    "Details 1", "Details 2", "Details 3",
    "Processing date", "Processing Method"
]

IBAN_REGEX = re.compile(r"^RO[A-Z0-9]{2,}$", re.IGNORECASE)

def money_to_csv(s: str) -> str:
    s = str(s).strip()
    if not s:
        return ""
    s = s.replace(" ", "").replace("\u00A0", "")
    s = s.replace(".", "#").replace(",", ".").replace("#", "")
    return s

class PandasModel(QAbstractTableModel):
    dataChangedSignal = Signal()
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        val = self._df.iat[index.row(), index.column()]
        if role in (Qt.DisplayRole, Qt.EditRole):
            return "" if pd.isna(val) else str(val)
        if role == Qt.ForegroundRole and self._df.columns[index.column()] in ("Payer Account IBAN", "Payee Account IBAN"):
            text = "" if pd.isna(val) else str(val).strip()
            if text and not IBAN_REGEX.match(text):
                from PySide6.QtGui import QBrush, QColor
                return QBrush(QColor("#b00020"))
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        return str(self._df.columns[section]) if orientation == Qt.Horizontal else str(section + 1)

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable if index.isValid() else Qt.ItemIsEnabled

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole and index.isValid():
            self._df.iat[index.row(), index.column()] = value
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
            self.dataChangedSignal.emit()
            return True
        return False

    def insertRows(self, position, rows=1, parent=QModelIndex()):
        self.beginInsertRows(QModelIndex(), position, position + rows - 1)
        empty = pd.DataFrame([[""] * self.columnCount() for _ in range(rows)], columns=self._df.columns)
        top = self._df.iloc[:position]
        bottom = self._df.iloc[position:]
        self._df = pd.concat([top, empty, bottom], ignore_index=True)
        self.endInsertRows()
        self.dataChangedSignal.emit()
        return True

    def removeRows(self, position, rows=1, parent=QModelIndex()):
        if position < 0 or position >= self.rowCount():
            return False
        self.beginRemoveRows(QModelIndex(), position, min(position + rows - 1, self.rowCount() - 1))
        self._df = self._df.drop(self._df.index[position:position + rows]).reset_index(drop=True)
        self.endRemoveRows()
        self.dataChangedSignal.emit()
        return True

class CompanyTab(QWidget):
    titleChanged = Signal(QWidget, str)

    def __init__(self, name="Firma"):
        super().__init__()
        self.company_name = name
        self.df = pd.DataFrame(columns=COLUMNS)
        self.model = PandasModel(self.df)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Nume companie (redenumire tab via semnal)
        title_bar = QHBoxLayout()
        self.name_edit = QLineEdit(self.company_name)
        self.name_edit.setPlaceholderText("Nume Firma")
        self.name_edit.textChanged.connect(self._emit_title_change)
        title_bar.addWidget(QLabel("Firma:"))
        title_bar.addWidget(self.name_edit, 1)
        layout.addLayout(title_bar)

        # Calea CSV
        row = QHBoxLayout()
        row.addWidget(QLabel("Locatie Export CSV :"))
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Alege unde sa salvezi export.csv")
        row.addWidget(self.path_edit, 1)
        b = QPushButton("Browse…")
        b.clicked.connect(self.choose_path)
        row.addWidget(b)
        layout.addLayout(row)

        # Tabel
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._ctx_menu)
        layout.addWidget(self.table, 1)

        # Opțiuni export
        opt = QHBoxLayout()
        self.no_header_chk = QCheckBox("No header in CSV"); self.no_header_chk.setChecked(True)
        self.crlf_chk = QCheckBox("CRLF line endings");     self.crlf_chk.setChecked(True)
        self.bom_chk  = QCheckBox("UTF-8 with BOM");        self.bom_chk.setChecked(True)
        opt.addWidget(self.no_header_chk); opt.addWidget(self.crlf_chk); opt.addWidget(self.bom_chk); opt.addStretch(1)
        layout.addLayout(opt)

        # Butoane
        btns = QHBoxLayout()
        add_btn = QPushButton("+ Add");     add_btn.clicked.connect(lambda: self.add_row())
        del_btn = QPushButton("− Delete");  del_btn.clicked.connect(self.delete_selected)
        import_btn = QPushButton("Import Excel/CSV…"); import_btn.clicked.connect(self.load_data)
        export_btn = QPushButton("Export CSV");        export_btn.clicked.connect(self.export_csv)
        btns.addWidget(add_btn); btns.addWidget(del_btn); btns.addStretch(1); btns.addWidget(import_btn); btns.addWidget(export_btn)
        layout.addLayout(btns)

        # Total live
        info = QHBoxLayout()
        self.total_label = QLabel("Total : 0.00")
        info.addWidget(self.total_label); info.addStretch(1)
        layout.addLayout(info)

        self.model.dataChangedSignal.connect(self.update_total)
        self.update_total()

    def _emit_title_change(self):
        self.company_name = self.name_edit.text().strip()
        self.titleChanged.emit(self, self.company_name or "Firma")

    def choose_path(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV As", self.path_edit.text() or "export.csv", "CSV (*.csv)")
        if path:
            if not path.lower().endswith(".csv"):
                path += ".csv"
            self.path_edit.setText(path)

    def _ctx_menu(self, pos):
        m = QMenu(self)
        m.addAction("Adauga rand", lambda: self.add_row(self.table.currentIndex().row() + 1))
        m.addAction("Sterge rand", self.delete_selected)
        m.exec(self.table.viewport().mapToGlobal(pos))

    def add_row(self, position=None):
        if position is None:
            position = self.model.rowCount()
        self.model.insertRows(position, 1)
        # PO_No. auto-increment
        try:
            if position > 0:
                prev = self.model._df.at[position - 1, "PO_No."]
                n = int(str(prev)) if str(prev).isdigit() else position + 1
            else:
                n = 1
            self.model._df.at[position, "PO_No."] = n
        except Exception:
            pass

    def delete_selected(self):
        idx = self.table.currentIndex()
        if idx.isValid():
            self.model.removeRows(idx.row(), 1)

    def load_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel/CSV", "", "Excel/CSV (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=object, keep_default_na=False)
            else:
                df = pd.read_excel(path, dtype=object)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df = df[COLUMNS].fillna("")
            self.df = df
            self.model = PandasModel(self.df)
            self.table.setModel(self.model)
            self.model.dataChangedSignal.connect(self.update_total)
            self.update_total()
        except Exception as e:
            QMessageBox.critical(self, "Import error", str(e))

    def export_csv(self):
        path = self.path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "No path", "Choose a CSV path first.")
            return
        df = self.model._df.copy()

        # normalize Amount
        if "Amount" in df.columns:
            df["Amount"] = df["Amount"].map(money_to_csv)

        # strip CR/LF din text
        for col in df.columns:
            df[col] = df[col].map(lambda x: str(x).replace("\r", " ").replace("\n", " ").strip())

        # avertizare simplă IBAN
        bad = []
        for col in ("Payer Account IBAN", "Payee Account IBAN"):
            for i, val in enumerate(df[col].tolist()):
                if str(val).strip() and not IBAN_REGEX.match(str(val).strip()):
                    bad.append((i + 1, col))
        if bad:
            QMessageBox.warning(self, "Validation", f"Some IBANs look invalid (should start with RO). Sample rows: {bad[:5]}")

        header = not self.no_header_chk.isChecked()
        encoding = "utf-8-sig" if self.bom_chk.isChecked() else "utf-8"
        line_ending = "\r\n" if self.crlf_chk.isChecked() else "\n"

        try:
            tmp = path + ".tmp"
            df.to_csv(tmp, index=False, header=header, encoding=encoding, lineterminator=line_ending)
            if os.path.exists(path):
                os.remove(path)
            os.replace(tmp, path)
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))
            return

        QMessageBox.information(self, "Exported", f"CSV saved:\n{path}")

    def update_total(self):
        try:
            s = 0.0
            for x in self.model._df["Amount"].tolist():
                if str(x).strip():
                    try:
                        s += float(money_to_csv(x))
                    except:
                        pass
            self.total_label.setText(f"Total: {s:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        except Exception:
            self.total_label.setText("Total : 0.00")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VER}")
        self.resize(1200, 720)

        # temă dark dacă există style.qss lângă fișier
        qss = Path(__file__).with_name("style.qss")
        if qss.exists():
            self.setStyleSheet(qss.read_text(encoding="utf-8"))

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.remove_tab)
        self.setCentralWidget(self.tabs)

        tb = QToolBar("Main")
        self.addToolBar(tb)
        add_act  = QAction("Adauga Firma", self, triggered=self.add_company_tab)
        rm_act   = QAction("Sterge Firma", self, triggered=lambda: self.remove_tab(self.tabs.currentIndex()))
        save_act = QAction("Salveaza Profil…", self, triggered=self.save_profile)
        load_act = QAction("Incarca Profil…", self, triggered=self.load_profile)
        tb.addAction(add_act); tb.addAction(rm_act); tb.addSeparator(); tb.addAction(save_act); tb.addAction(load_act)

        self.setStatusBar(QStatusBar())
        self.add_company_tab()

    def add_company_tab(self):
        tab = CompanyTab(f"Company {self.tabs.count()+1}")
        idx = self.tabs.addTab(tab, tab.company_name)
        self.tabs.setCurrentIndex(idx)
        tab.titleChanged.connect(self._rename_tab)

    def _rename_tab(self, widget, new_title: str):
        idx = self.tabs.indexOf(widget)
        if idx != -1:
            self.tabs.setTabText(idx, new_title)

    def remove_tab(self, idx):
        if idx < 0:
            return
        name = self.tabs.tabText(idx)
        if QMessageBox.question(self, "Sterge Firma", f"Sterge '{name}' tab?") == QMessageBox.Yes:
            w = self.tabs.widget(idx)
            self.tabs.removeTab(idx)
            w.deleteLater()
            if self.tabs.count() == 0:
                self.add_company_tab()

    def save_profile(self):
        path, _ = QFileDialog.getSaveFileName(self, "Salveaza Profil", "", "Profil (*.json)")
        if not path:
            return
        data = []
        for i in range(self.tabs.count()):
            w: CompanyTab = self.tabs.widget(i)
            data.append({
                "name": w.company_name,
                "path": w.path_edit.text(),
                "options": {
                    "no_header": w.no_header_chk.isChecked(),
                    "crlf": w.crlf_chk.isChecked(),
                    "bom": w.bom_chk.isChecked()
                },
                "rows": w.model._df.fillna("").values.tolist()
            })
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.statusBar().showMessage(f"Saved profile: {path}", 3000)

    def load_profile(self):
        path, _ = QFileDialog.getOpenFileName(self, "Incarca Profil", "", "Profil (*.json)")
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.tabs.clear()
        for comp in data:
            tab = CompanyTab(comp.get("name", "Firma"))
            tab.path_edit.setText(comp.get("path", ""))
            opts = comp.get("options", {})
            tab.no_header_chk.setChecked(opts.get("no_header", True))
            tab.crlf_chk.setChecked(opts.get("crlf", True))
            tab.bom_chk.setChecked(opts.get("bom", True))
            rows = comp.get("rows", [])
            df = pd.DataFrame(rows, columns=COLUMNS)
            tab.df = df
            tab.model = PandasModel(df)
            tab.table.setModel(tab.model)
            tab.model.dataChangedSignal.connect(tab.update_total)
            tab.update_total()
            self.tabs.addTab(tab, tab.company_name)
            tab.titleChanged.connect(self._rename_tab)
        if self.tabs.count() == 0:
            self.add_company_tab()

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()


