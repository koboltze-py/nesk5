"""
Sanitätsmaterial – Buchungsverlauf
Transaktionshistorie mit Filter, Pagination und CSV-Export.
"""

import csv
import re
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QLineEdit, QDateEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QMessageBox, QFileDialog,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor

PAGE_SIZE = 150

_TYP_FARBEN = {
    "einlagerung": QColor("#E8F5E9"),
    "entnahme":    QColor("#FFF3E0"),
    "verbrauch":   QColor("#FFF3E0"),
    "korrektur":   QColor("#E3F2FD"),
}


def _format_datum(iso: str) -> str:
    try:
        return datetime.strptime(iso[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return iso


class VerlaufView(QWidget):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._offset = 0
        self._total = 0
        self._setup_ui()
        self._load_filter_combos()
        self._load_data()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        lbl_title = QLabel("Buchungsverlauf")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Alle Einlagerungen, Entnahmen, Verbräuche und Korrekturen")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        # Filter-Zeile 1
        f1 = QHBoxLayout(); f1.setSpacing(10)
        f1.addWidget(QLabel("Artikel:"))
        self.cb_art = QComboBox()
        self.cb_art.setMinimumWidth(200)
        f1.addWidget(self.cb_art)
        f1.addWidget(QLabel("Typ:"))
        self.cb_typ = QComboBox()
        self.cb_typ.addItems(["Alle", "einlagerung", "entnahme", "verbrauch", "korrektur"])
        self.cb_typ.setMinimumWidth(120)
        f1.addWidget(self.cb_typ)
        f1.addWidget(QLabel("Von:"))
        self.de_von = QDateEdit()
        self.de_von.setDisplayFormat("dd.MM.yyyy")
        self.de_von.setCalendarPopup(True)
        self.de_von.setDate(QDate.currentDate().addDays(-7))
        self.de_von.setMaximumWidth(130)
        f1.addWidget(self.de_von)
        f1.addWidget(QLabel("Bis:"))
        self.de_bis = QDateEdit(QDate.currentDate())
        self.de_bis.setDisplayFormat("dd.MM.yyyy")
        self.de_bis.setCalendarPopup(True)
        self.de_bis.setMaximumWidth(130)
        f1.addWidget(self.de_bis)
        f1.addStretch()
        layout.addLayout(f1)

        # Schnellfilter
        qf = QHBoxLayout(); qf.setSpacing(6)
        qf.addWidget(QLabel("Zeitraum:"))
        for label, fn in [
            ("Heute",       lambda: (QDate.currentDate(), QDate.currentDate())),
            ("7 Tage",      lambda: (QDate.currentDate().addDays(-7), QDate.currentDate())),
            ("30 Tage",     lambda: (QDate.currentDate().addDays(-30), QDate.currentDate())),
            ("Dieses Jahr", lambda: (QDate(QDate.currentDate().year(), 1, 1), QDate.currentDate())),
            ("Alles",       lambda: (QDate(2000, 1, 1), QDate.currentDate())),
        ]:
            btn_q = QPushButton(label)
            btn_q.setObjectName("btn_secondary")
            btn_q.setMaximumHeight(30)
            btn_q.setStyleSheet("QPushButton{padding:3px 10px;font-size:12px;}")
            def _make_cb(f):
                def _cb():
                    von, bis = f()
                    self.de_von.setDate(von)
                    self.de_bis.setDate(bis)
                    self._search()
                return _cb
            btn_q.clicked.connect(_make_cb(fn))
            qf.addWidget(btn_q)
        qf.addStretch()
        layout.addLayout(qf)

        # Filter-Zeile 2
        f2 = QHBoxLayout(); f2.setSpacing(10)
        f2.addWidget(QLabel("Suche:"))
        self.le_suche = QLineEdit()
        self.le_suche.setPlaceholderText("Artikelname, Von, Bemerkung ...")
        self.le_suche.setMinimumWidth(250)
        f2.addWidget(self.le_suche)
        btn_suchen = QPushButton("🔍  Suchen")
        btn_suchen.setObjectName("btn_primary")
        btn_suchen.clicked.connect(self._search)
        f2.addWidget(btn_suchen)
        btn_reset = QPushButton("Zurücksetzen")
        btn_reset.setObjectName("btn_secondary")
        btn_reset.clicked.connect(self._reset_filter)
        f2.addWidget(btn_reset)
        f2.addStretch()
        btn_export = QPushButton("📄  CSV-Export")
        btn_export.setObjectName("btn_secondary")
        btn_export.clicked.connect(self._export)
        f2.addWidget(btn_export)
        layout.addLayout(f2)

        # Tabelle
        self._tbl = QTableWidget(0, 7)
        self._tbl.setHorizontalHeaderLabels(
            ["Datum", "Typ", "Artikel", "Menge", "Von", "Bemerkung", "Erstellt"]
        )
        self._tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._tbl.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        for c in [0, 1, 3, 4, 6]:
            self._tbl.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.verticalHeader().setVisible(False)
        self._tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl.setAlternatingRowColors(True)
        self._tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self._tbl)

        # Paginierung
        pag_row = QHBoxLayout()
        self._btn_prev = QPushButton("◀  Zurück")
        self._btn_prev.setObjectName("btn_secondary")
        self._btn_prev.clicked.connect(self._prev_page)
        pag_row.addWidget(self._btn_prev)
        self._lbl_page = QLabel("")
        self._lbl_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pag_row.addWidget(self._lbl_page, stretch=1)
        self._btn_next = QPushButton("Weiter  ▶")
        self._btn_next.setObjectName("btn_secondary")
        self._btn_next.clicked.connect(self._next_page)
        pag_row.addWidget(self._btn_next)
        layout.addLayout(pag_row)

    def _load_filter_combos(self):
        self.cb_art.clear()
        self.cb_art.addItem("Alle Artikel", None)
        for a in self.db.get_artikel():
            self.cb_art.addItem(a["bezeichnung"], a["id"])

    def _get_filter(self) -> dict:
        return {
            "artikel_id": self.cb_art.currentData(),
            "typ": self.cb_typ.currentText(),
            "datum_von": self.de_von.date().toString("yyyy-MM-dd"),
            "datum_bis": self.de_bis.date().toString("yyyy-MM-dd"),
            "suche": self.le_suche.text().strip() or None,
        }

    def _load_data(self):
        f = self._get_filter()
        self._total = self.db.count_buchungen(**f)
        buchungen = self.db.get_buchungen(limit=PAGE_SIZE, offset=self._offset, **f)

        self._tbl.setRowCount(len(buchungen))

        # Einsatz-Gruppen ermitteln: verbrauch-Zeilen mit gleicher Einsatz-ID
        _GROUP_COLORS = [QColor("#FFD180"), QColor("#FFF176")]
        eid_list = []
        eid_count = {}
        for b in buchungen:
            if b.get("typ") == "verbrauch":
                m = re.search(r"\(ID\s*(\d+)\)", b.get("bemerkung", "") or "")
                eid = int(m.group(1)) if m else None
            else:
                eid = None
            eid_list.append(eid)
            if eid is not None:
                eid_count[eid] = eid_count.get(eid, 0) + 1
        eid_color_map = {}
        gi = 0
        for eid, cnt in eid_count.items():
            if cnt > 1:
                eid_color_map[eid] = _GROUP_COLORS[gi % len(_GROUP_COLORS)]
                gi += 1

        for r, (b, eid) in enumerate(zip(buchungen, eid_list)):
            typ = b.get("typ", "")
            menge = int(b.get("menge", 0))
            if typ == "verbrauch":
                menge_str = str(abs(menge))
            else:
                menge_str = f"+{menge}" if menge > 0 else str(menge)
            bemerkung = b.get("bemerkung", "") or ""
            if typ == "verbrauch":
                bemerkung = re.sub(r"\s*\(G?ID\s*\d+\)", "", bemerkung).strip()
            vals = [
                _format_datum(b.get("datum", "")), typ,
                b.get("artikel_name", ""), menge_str,
                b.get("von", ""), bemerkung,
                b.get("erstellt_am", "")[:16] if b.get("erstellt_am") else "",
            ]
            if eid is not None and eid in eid_color_map:
                color = eid_color_map[eid]
                tooltip = f"Mehrere Artikel in diesem Einsatz (ID {eid})"
            else:
                color = _TYP_FARBEN.get(typ, QColor("white"))
                tooltip = ""
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setBackground(color)
                if tooltip:
                    it.setToolTip(tooltip)
                self._tbl.setItem(r, c, it)

        pages = max(1, (self._total + PAGE_SIZE - 1) // PAGE_SIZE)
        cur_page = self._offset // PAGE_SIZE + 1
        self._lbl_page.setText(f"Seite {cur_page} / {pages}  ({self._total} Einträge)")
        self._btn_prev.setEnabled(self._offset > 0)
        self._btn_next.setEnabled(self._offset + PAGE_SIZE < self._total)

    def _search(self):
        self._offset = 0
        self._load_data()

    def _reset_filter(self):
        self.cb_art.setCurrentIndex(0)
        self.cb_typ.setCurrentIndex(0)
        self.de_von.setDate(QDate.currentDate().addDays(-7))
        self.de_bis.setDate(QDate.currentDate())
        self.le_suche.clear()
        self._offset = 0
        self._load_data()

    def _prev_page(self):
        self._offset = max(0, self._offset - PAGE_SIZE)
        self._load_data()

    def _next_page(self):
        if self._offset + PAGE_SIZE < self._total:
            self._offset += PAGE_SIZE
            self._load_data()

    def _export(self):
        pfad, _ = QFileDialog.getSaveFileName(
            self, "Buchungsverlauf exportieren",
            f"Buchungsverlauf_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV-Dateien (*.csv)"
        )
        if not pfad:
            return
        f = self._get_filter()
        buchungen = self.db.get_buchungen(limit=1_000_000, offset=0, **f)
        try:
            with open(pfad, "w", newline="", encoding="utf-8-sig") as fp:
                w = csv.writer(fp, delimiter=";")
                w.writerow(["Datum", "Typ", "Artikel", "Menge", "Von", "Bemerkung"])
                for b in buchungen:
                    menge = b.get("menge", 0)
                    w.writerow([
                        _format_datum(b.get("datum", "")), b.get("typ", ""),
                        b.get("artikel_name", ""),
                        f"+{menge}" if menge > 0 else str(menge),
                        b.get("von", ""), b.get("bemerkung", ""),
                    ])
            QMessageBox.information(self, "Exportiert", f"Gespeichert: {pfad}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", str(e))

    def showEvent(self, event):
        super().showEvent(event)
        self._load_filter_combos()
        self._load_data()
