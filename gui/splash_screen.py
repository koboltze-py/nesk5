"""
Ladebildschirm (Splash Screen)
Wird beim App-Start angezeigt während Backup, Migration und Sync laufen.
"""
import os
import sys

from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap, QIcon

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Farben passend zum nesk3.ico Logo
_BG      = "#253545"   # dunkler Hintergrund (tiefer als Logo-Basis)
_CARD    = "#354A5D"   # Logo-Hauptfarbe (dominant im Icon)
_ACCENT  = "#5B8AAA"   # hellblau-teal Akzent aus Logo-Palette
_GOLD    = "#C0944A"   # gold/amber Akzent aus Logo
_WHITE   = "#FFFFFF"
_GRAY    = "#9BB5C8"   # helles blaugrau


class SplashScreen(QWidget):
    """
    Frameless Ladebildschirm – zentriert, immer im Vordergrund.
    Aufruf:
        splash = SplashScreen()
        splash.show()
        QApplication.processEvents()
        splash.set_status("Schritt 1 ...")
        ...
        splash.finish(main_window)  # schließt Splash, übergibt Fokus
    """

    def __init__(self, version: str = ""):
        super().__init__()
        self._version = version
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.SplashScreen
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setFixedSize(480, 300)
        self._center()
        self._build_ui()

    def _center(self):
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            self.move(
                geo.center().x() - self.width() // 2,
                geo.center().y() - self.height() // 2,
            )

    def _build_ui(self):
        self.setStyleSheet(f"background-color: {_BG};")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Gold-Akzentstreifen oben ──────────────────────────────────────
        top_band = QWidget()
        top_band.setFixedHeight(4)
        top_band.setStyleSheet(f"background-color: {_GOLD};")
        outer.addWidget(top_band)

        # ── Hauptbereich ──────────────────────────────────────────────────
        content = QWidget()
        content.setStyleSheet(f"background-color: {_BG};")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(44, 28, 44, 22)
        layout.setSpacing(0)
        outer.addWidget(content, 1)

        # ── Logo + Titel nebeneinander ────────────────────────────────────
        header_row = QHBoxLayout()
        header_row.setSpacing(24)

        logo_lbl = QLabel()
        logo_lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        logo_px = self._load_logo(80)
        logo_lbl.setPixmap(logo_px)
        logo_lbl.setFixedSize(84, 84)
        header_row.addWidget(logo_lbl)

        title_col = QVBoxLayout()
        title_col.setSpacing(4)
        title_col.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        title = QLabel("NESK 3")
        title.setStyleSheet(
            f"color: {_WHITE}; font-size: 30pt; font-weight: bold; "
            f"letter-spacing: 3px; background: transparent;"
        )
        title_col.addWidget(title)

        sub = QLabel("DRK Flughafen Köln/Bonn")
        sub.setStyleSheet(f"color: {_ACCENT}; font-size: 11pt; background: transparent;")
        title_col.addWidget(sub)

        if self._version:
            ver = QLabel(f"Version {self._version}")
            ver.setStyleSheet(f"color: {_GRAY}; font-size: 9pt; background: transparent;")
            title_col.addWidget(ver)

        header_row.addLayout(title_col)
        header_row.addStretch()
        layout.addLayout(header_row)

        layout.addStretch()

        # ── Trennlinie ────────────────────────────────────────────────────
        line = QWidget()
        line.setFixedHeight(1)
        line.setStyleSheet(f"background-color: {_CARD};")
        layout.addWidget(line)

        layout.addSpacing(10)

        # ── Status-Zeile ──────────────────────────────────────────────────
        self._status = QLabel("Wird gestartet …")
        self._status.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self._status.setWordWrap(True)
        self._status.setStyleSheet(
            f"color: {_GRAY}; font-size: 9pt; background: transparent;"
        )
        layout.addWidget(self._status)

        # ── Blauer Akzentstreifen unten ───────────────────────────────────
        bot_band = QWidget()
        bot_band.setFixedHeight(3)
        bot_band.setStyleSheet(f"background-color: {_ACCENT};")
        outer.addWidget(bot_band)

    def _load_logo(self, size: int) -> QPixmap:
        """Lädt das nesk3.ico als Pixmap; Fallback: stilisiertes 'N'."""
        icon_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "Daten", "Logo", "nesk3.ico"
        )
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            px = icon.pixmap(size, size)
            if not px.isNull():
                return px

        # Fallback: stilisiertes "N" in Akzentfarbe
        px = QPixmap(size, size)
        px.fill(QColor(_CARD))
        painter = QPainter(px)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(_ACCENT), max(3, size // 10)))
        m = int(size * 0.2)
        painter.drawLine(m, m, m, size - m)
        painter.drawLine(m, m, size - m, size - m)
        painter.drawLine(size - m, m, size - m, size - m)
        painter.end()
        return px

    def set_status(self, message: str):
        """Aktualisiert die Status-Zeile und verarbeitet Events sofort."""
        self._status.setText(message)
        QApplication.processEvents()

    def finish(self, main_window=None):
        """Schließt den Splash Screen. Optional: Fokus an main_window übergeben."""
        self.close()
