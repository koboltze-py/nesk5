"""
Ladebildschirm (Splash Screen)
Wird beim App-Start angezeigt waehrend Backup, Migration und Sync laufen.
Enthaelt einen animierten Lichtstrahl-Effekt (Shimmer) ueber den gesamten Screen.
"""
import os
import sys

from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QApplication
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap, QIcon, QLinearGradient, QBrush

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Farben passend zum nesk3.ico Logo
_BG     = "#1C2B38"   # tiefer dunkler Hintergrund
_CARD   = "#354A5D"   # Logo-Hauptfarbe
_ACCENT = "#5B8AAA"   # hellblau-teal
_GOLD   = "#C0944A"   # gold/amber
_WHITE  = "#ECEFF4"   # leicht warm-weiss
_GRAY   = "#7A99B0"   # mittleres blaugrau


class _ShimmerOverlay(QWidget):
    """
    Transparentes Widget, das einen wandernden Lichtstrahl ueber den Splash Screen zeichnet.
    Wird immer ueber allen anderen Kindwidgets gerendert.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self._pos = -0.5   # normalisierte X-Position (0.0 = links, 1.0 = rechts)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(16)     # ~60 fps

    def _tick(self):
        self._pos += 0.007
        if self._pos > 1.5:
            self._pos = -0.5
        self.update()

    def paintEvent(self, event):
        w = self.width()
        h = self.height()
        cx = self._pos * w
        band = w * 0.28           # Breite des Lichtstrahls

        # Diagonaler Lichtstrahl (leicht schraeg)
        grad = QLinearGradient(cx - band, 0.0, cx + band * 0.6, float(h))
        grad.setColorAt(0.0,  QColor(255, 255, 255, 0))
        grad.setColorAt(0.42, QColor(255, 255, 255, 0))
        grad.setColorAt(0.50, QColor(255, 255, 255, 22))
        grad.setColorAt(0.58, QColor(255, 255, 255, 0))
        grad.setColorAt(1.0,  QColor(255, 255, 255, 0))

        painter = QPainter(self)
        painter.fillRect(0, 0, w, h, QBrush(grad))
        painter.end()


class SplashScreen(QWidget):
    """
    Frameless moderner Ladebildschirm mit Shimmer-Animation.
    Aufruf:
        splash = SplashScreen()
        splash.show()
        QApplication.processEvents()
        splash.set_status("Schritt ...")
        splash.finish(main_window)
    """

    def __init__(self, version=""):
        super().__init__()
        self._version = version
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.SplashScreen
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setFixedSize(520, 290)
        self._center()
        self._build_ui()
        # Shimmer-Overlay zuletzt hinzufuegen -> liegt ueber allem
        self._shimmer = _ShimmerOverlay(self)
        self._shimmer.setGeometry(0, 0, self.width(), self.height())
        self._shimmer.raise_()

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

        # Goldener Akzentstreifen oben
        top_band = QWidget()
        top_band.setFixedHeight(3)
        top_band.setStyleSheet(f"background-color: {_GOLD};")
        outer.addWidget(top_band)

        # Hauptbereich
        content = QWidget()
        content.setStyleSheet("background-color: transparent;")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(52, 36, 52, 0)
        layout.setSpacing(0)
        outer.addWidget(content, 1)

        # Logo + Titel nebeneinander
        header = QHBoxLayout()
        header.setSpacing(26)

        logo_lbl = QLabel()
        logo_lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        logo_lbl.setPixmap(self._load_logo(76))
        logo_lbl.setFixedSize(80, 80)
        logo_lbl.setStyleSheet("background: transparent;")
        header.addWidget(logo_lbl)

        text_col = QVBoxLayout()
        text_col.setSpacing(3)
        text_col.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        # Titel "NeSk" -- duenn, elegant, viel Buchstabenabstand
        name_lbl = QLabel("NeSk")
        name_lbl.setStyleSheet(
            f"color: {_WHITE}; font-size: 42pt; font-weight: 300; "
            f"letter-spacing: 8px; background: transparent;"
        )
        text_col.addWidget(name_lbl)

        sub_lbl = QLabel("DRK  \u00b7  Flughafen K\u00f6ln / Bonn")
        sub_lbl.setStyleSheet(
            f"color: {_ACCENT}; font-size: 9.5pt; "
            f"letter-spacing: 2px; background: transparent;"
        )
        text_col.addWidget(sub_lbl)

        if self._version:
            ver_lbl = QLabel(f"v{self._version}")
            ver_lbl.setStyleSheet(
                f"color: {_GRAY}; font-size: 8pt; background: transparent;"
            )
            text_col.addWidget(ver_lbl)

        header.addLayout(text_col)
        header.addStretch()
        layout.addLayout(header)
        layout.addStretch()

        # Trennlinie
        divider = QWidget()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background-color: {_CARD};")
        layout.addWidget(divider)

        layout.addSpacing(12)

        # Status-Zeile
        self._status = QLabel("Wird gestartet ...")
        self._status.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self._status.setWordWrap(True)
        self._status.setStyleSheet(
            f"color: {_GRAY}; font-size: 8.5pt; background: transparent;"
        )
        layout.addWidget(self._status)
        layout.addSpacing(18)

        # Teal-Akzentstreifen unten
        bot_band = QWidget()
        bot_band.setFixedHeight(2)
        bot_band.setStyleSheet(f"background-color: {_ACCENT};")
        outer.addWidget(bot_band)

    def _load_logo(self, size):
        """Laedt nesk3.ico als Pixmap; Fallback: stilisiertes N."""
        icon_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "Daten", "Logo", "nesk3.ico",
        )
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            px = icon.pixmap(size, size)
            if not px.isNull():
                return px

        # Fallback: stilisiertes "N"
        px = QPixmap(size, size)
        px.fill(QColor(_CARD))
        p = QPainter(px)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QPen(QColor(_ACCENT), max(3, size // 10)))
        m = int(size * 0.2)
        p.drawLine(m, m, m, size - m)
        p.drawLine(m, m, size - m, size - m)
        p.drawLine(size - m, m, size - m, size - m)
        p.end()
        return px

    def set_status(self, message):
        """Aktualisiert die Status-Zeile und verarbeitet Events sofort."""
        self._status.setText(message)
        QApplication.processEvents()

    def finish(self, main_window=None):
        """Stoppt Animation und schliesst den Splash Screen."""
        self._shimmer._timer.stop()
        self.close()