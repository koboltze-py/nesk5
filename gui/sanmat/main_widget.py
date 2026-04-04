"""
SanmatWidget – Hauptwidget für die Sanitätsmaterial-Verwaltung in Nesk3.
Enthält eine eigene Tab-Navigation und einen gestapelten Inhaltsbereich.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel,
    QPushButton, QStackedWidget, QFrame, QButtonGroup,
)
from PySide6.QtCore import Qt

from database.sanmat_db import SanmatDB
from gui.sanmat.artikel      import ArtikelView
from gui.sanmat.bestand      import BestandView
from gui.sanmat.entnahme     import EntnahmeView
from gui.sanmat.verlauf      import VerlaufView
from gui.sanmat.verbrauch    import VerbrauchView
from gui.sanmat.einstellungen import EinstellungenView

_NAV = [
    ("📋", "Artikel"),
    ("📦", "Bestand"),
    ("➡", "Entnahme"),
    ("📜", "Verlauf"),
    ("🧰", "Verbrauch"),
    ("ℹ", "Info"),
]


class _TabButton(QPushButton):
    def __init__(self, icon: str, text: str, parent=None):
        super().__init__(f" {icon}  {text}", parent)
        self.setCheckable(True)
        self.setMinimumHeight(40)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._apply_style(False)

    def _apply_style(self, active: bool):
        if active:
            self.setStyleSheet("""
                QPushButton {
                    background-color: rgba(255,255,255,38);
                    color: white;
                    border: none;
                    border-left: 3px solid #5bc0de;
                    border-radius: 0px;
                    padding: 8px 14px;
                    text-align: left;
                    font-size: 12px;
                    font-weight: bold;
                }
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: #c0cfe0;
                    border: none;
                    border-left: 3px solid transparent;
                    border-radius: 0px;
                    padding: 8px 14px;
                    text-align: left;
                    font-size: 12px;
                }
                QPushButton:hover {
                    background-color: rgba(255,255,255,13);
                    color: white;
                }
            """)

    def setActive(self, active: bool):
        self._apply_style(active)
        self.setChecked(active)


class SanmatWidget(QWidget):
    """Sanitätsmaterial-Modul eingebettet in Nesk3."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = SanmatDB()
        self.db.initialize()
        self._setup_ui()
        self._navigate(0)

    def _setup_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Linke Sub-Sidebar
        sidebar = QWidget()
        sidebar.setFixedWidth(170)
        sidebar.setStyleSheet("background-color: #2d4155;")
        sb_layout = QVBoxLayout(sidebar)
        sb_layout.setContentsMargins(0, 12, 0, 12)
        sb_layout.setSpacing(2)

        # Titel
        lbl_header = QLabel("Sanmat")
        lbl_header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_header.setStyleSheet(
            "color: #5bc0de; font-size: 14px; font-weight: bold; padding: 8px 0px;"
        )
        sb_layout.addWidget(lbl_header)

        lbl_sub = QLabel("Sanitätsmaterial")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_sub.setStyleSheet("color:#7a9bb5; font-size:10px; padding-bottom:10px;")
        sb_layout.addWidget(lbl_sub)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#3d5570;")
        sb_layout.addWidget(sep)

        # Navigations-Tabs
        self._tab_buttons: list[_TabButton] = []
        self._btn_group = QButtonGroup(self)
        self._btn_group.setExclusive(True)
        for i, (icon, label) in enumerate(_NAV):
            btn = _TabButton(icon, label)
            btn.clicked.connect(lambda _, idx=i: self._navigate(idx))
            self._btn_group.addButton(btn)
            self._tab_buttons.append(btn)
            sb_layout.addWidget(btn)

        sb_layout.addStretch()
        root.addWidget(sidebar)

        # Trennlinie
        line = QFrame()
        line.setFrameShape(QFrame.Shape.VLine)
        line.setStyleSheet("color:#d9d9d9;")
        root.addWidget(line)

        # Inhaltsbereich
        self._stack = QStackedWidget()
        self._stack.addWidget(ArtikelView(self.db))
        self._stack.addWidget(BestandView(self.db))
        self._stack.addWidget(EntnahmeView(self.db))
        self._stack.addWidget(VerlaufView(self.db))
        self._stack.addWidget(VerbrauchView(self.db))
        self._stack.addWidget(EinstellungenView(self.db))
        root.addWidget(self._stack, 1)

    def _navigate(self, idx: int):
        self._stack.setCurrentIndex(idx)
        for i, btn in enumerate(self._tab_buttons):
            btn.setActive(i == idx)

    def showEvent(self, event):
        super().showEvent(event)
