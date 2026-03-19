"""
Ladebildschirm (Splash Screen) - Moderne Variante
- grosses Logo zentriert
- rotierender Teal-Ring um das Logo
- pulsierender Glow-Ring (Opacity-Animation)
- "NeSk" Schriftzug mit wanderndem Lichteffekt
- Statuszeile am unteren Rand
"""
import os, sys, math

from PySide6.QtWidgets import QWidget, QApplication
from PySide6.QtCore    import Qt, QTimer, QRectF, QPointF
from PySide6.QtGui     import (
    QColor, QPainter, QPen, QPixmap, QIcon,
    QLinearGradient, QRadialGradient, QBrush, QFont, QFontMetricsF,
    QConicalGradient,
)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# ---------------------------------------------------------------------------
# Farbpalette (Logo-Farben)
# ---------------------------------------------------------------------------
_BG        = QColor("#131E28")   # sehr tiefer Hintergrund
_RING1     = QColor("#5B8AAA")   # Teal-Blau (rotierender Ring)
_RING2     = QColor("#C0944A")   # Gold-Amber (zweiter Ring, gegenlaeutig)
_GLOW      = QColor(91, 138, 170, 60)   # Teal mit Alpha fuer Glow
_WHITE     = QColor("#ECEFF4")
_SUBTEXT   = QColor("#5B8AAA")
_GRAY      = QColor("#4A6880")
_LOGO_BG   = QColor("#1C2B38")   # Kreis hinter dem Logo


class SplashScreen(QWidget):
    """
    Frameless moderner Splash Screen.
    Alles wird in paintEvent() gezeichnet - kein Layout, pure QPainter.

    Aufruf:
        splash = SplashScreen()
        splash.show();  QApplication.processEvents()
        splash.set_status("...")
        splash.finish()
    """

    W, H = 540, 360

    def __init__(self, version=""):
        super().__init__()
        self._version   = version
        self._angle1    = 0.0      # Winkel rotierender Ring 1 (Grad)
        self._angle2    = 180.0    # Ring 2 startet versetzt
        self._pulse     = 0.0      # 0..2pi fuer sin-Glow
        self._shimmer   = -0.4     # 0..1 Position des Lichtstreifens auf "NeSk"
        self._status    = "Wird gestartet \u2026"

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.SplashScreen
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setFixedSize(self.W, self.H)

        # Logo laden
        icon_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "Daten", "Logo", "nesk3.ico",
        )
        self._logo_px = None
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            px   = icon.pixmap(120, 120)
            if not px.isNull():
                self._logo_px = px

        # Zentrieren
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            self.move(
                geo.center().x() - self.W // 2,
                geo.center().y() - self.H // 2,
            )

        # Timer ~60 fps
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(16)

    # ------------------------------------------------------------------
    def _tick(self):
        self._angle1  = (self._angle1 + 1.4) % 360
        self._angle2  = (self._angle2 - 0.9) % 360
        self._pulse  += 0.045
        self._shimmer += 0.009
        if self._shimmer > 1.4:
            self._shimmer = -0.4
        self.update()

    # ------------------------------------------------------------------
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        W, H = float(self.W), float(self.H)

        # ── Hintergrund ──────────────────────────────────────────────────
        # radialer Gradient von Mitte nach aussen: etwas heller in der Mitte
        bg_grad = QRadialGradient(W / 2, H * 0.42, W * 0.55)
        bg_grad.setColorAt(0.0, QColor("#1E3040"))
        bg_grad.setColorAt(1.0, QColor("#0E1820"))
        p.fillRect(0, 0, int(W), int(H), QBrush(bg_grad))

        # ── Goldener Streifen oben ────────────────────────────────────────
        p.fillRect(0, 0, int(W), 3, QBrush(QColor("#C0944A")))

        # ── Logo-Bereich (Mitte oben) ─────────────────────────────────────
        cx, cy = W / 2, H * 0.40
        logo_r = 68.0        # Radius des Logo-Kreises

        # Glow-Ring (pulsierend)
        glow_alpha = int(30 + 25 * math.sin(self._pulse))
        glow_r = logo_r + 22
        glow_grad = QRadialGradient(cx, cy, glow_r)
        glow_grad.setColorAt(0.70, QColor(91, 138, 170, glow_alpha))
        glow_grad.setColorAt(1.00, QColor(91, 138, 170, 0))
        p.setBrush(QBrush(glow_grad))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - glow_r, cy - glow_r, glow_r * 2, glow_r * 2))

        # Hintergrundkreis hinter dem Logo
        p.setBrush(QBrush(QColor("#1C2B38")))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawEllipse(QRectF(cx - logo_r, cy - logo_r, logo_r * 2, logo_r * 2))

        # Rotierender Teal-Ring (Ring 1)
        arc_pen1 = QPen(_RING1, 3.0, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap)
        p.setPen(arc_pen1)
        p.setBrush(Qt.BrushStyle.NoBrush)
        ring1_r = logo_r + 10
        ring1_rect = QRectF(cx - ring1_r, cy - ring1_r, ring1_r * 2, ring1_r * 2)
        # 240 Grad Bogen, der sich dreht
        start1 = int(-(self._angle1) * 16)
        p.drawArc(ring1_rect, start1, 240 * 16)

        # Gegenlaeuftiger Gold-Ring (Ring 2, duenner)
        arc_pen2 = QPen(_RING2, 1.5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap)
        p.setPen(arc_pen2)
        ring2_r = logo_r + 16
        ring2_rect = QRectF(cx - ring2_r, cy - ring2_r, ring2_r * 2, ring2_r * 2)
        start2 = int(-(self._angle2) * 16)
        p.drawArc(ring2_rect, start2, 120 * 16)

        # Logo Pixmap (oder Fallback-"N")
        if self._logo_px:
            lw, lh = self._logo_px.width(), self._logo_px.height()
            p.drawPixmap(
                int(cx - lw / 2), int(cy - lh / 2),
                self._logo_px,
            )
        else:
            # Fallback: "N" in der Mitte
            font = QFont("Segoe UI", 40, QFont.Weight.Light)
            p.setFont(font)
            p.setPen(QPen(_RING1))
            fm = QFontMetricsF(font)
            tw = fm.horizontalAdvance("N")
            p.drawText(QPointF(cx - tw / 2, cy + fm.ascent() / 2 - 4), "N")

        # ── "NeSk" Schriftzug ─────────────────────────────────────────────
        font_title = QFont("Segoe UI", 36, QFont.Weight.Light)
        p.setFont(font_title)
        fm_t = QFontMetricsF(font_title)
        text = "NeSk"
        tw   = fm_t.horizontalAdvance(text)
        tx   = cx - tw / 2
        ty   = cy + logo_r + 38    # unterhalb des Logo-Bereichs

        # Basis-Text (grau)
        p.setPen(QPen(QColor(180, 200, 215, 120)))
        p.drawText(QPointF(tx, ty), text)

        # Shimmer-Overlay: heller Streifen wandert durch den Text
        sh_cx = tx + self._shimmer * (tw + 60) - 20
        sh_width = tw * 0.35
        sh_grad = QLinearGradient(sh_cx - sh_width, float(ty) - 40,
                                  sh_cx + sh_width, float(ty))
        sh_grad.setColorAt(0.0, QColor(255, 255, 255, 0))
        sh_grad.setColorAt(0.5, QColor(255, 255, 255, 200))
        sh_grad.setColorAt(1.0, QColor(255, 255, 255, 0))
        p.setPen(QPen(QBrush(sh_grad), 0))
        p.drawText(QPointF(tx, ty), text)

        # Letter-spacing Simulation: Buchstaben-Abstand durch Extra-Zeichnung
        # (Qt unterstuetzt kein letter-spacing in drawText direkt;
        #  wir zeichnen einfach mit font_title und geniessen den Font-Standard)

        # ── Untertitel ────────────────────────────────────────────────────
        font_sub = QFont("Segoe UI", 9)
        p.setFont(font_sub)
        fm_s = QFontMetricsF(font_sub)
        sub  = "DRK  \u00b7  Flughafen K\u00f6ln / Bonn"
        sw   = fm_s.horizontalAdvance(sub)
        p.setPen(QPen(_SUBTEXT))
        p.drawText(QPointF(cx - sw / 2, ty + 26), sub)

        # Version
        if self._version:
            font_ver = QFont("Segoe UI", 8)
            p.setFont(font_ver)
            fm_v = QFontMetricsF(font_ver)
            vt   = f"v{self._version}"
            vw   = fm_v.horizontalAdvance(vt)
            p.setPen(QPen(_GRAY))
            p.drawText(QPointF(cx - vw / 2, ty + 44), vt)

        # ── Status-Zeile ──────────────────────────────────────────────────
        vert_sep_y = H - 42
        p.setPen(QPen(QColor("#253545")))
        p.drawLine(QPointF(32, vert_sep_y), QPointF(W - 32, vert_sep_y))

        font_st = QFont("Segoe UI", 8)
        p.setFont(font_st)
        p.setPen(QPen(QColor(90, 120, 145)))
        p.drawText(QPointF(32, H - 22), self._status)

        # Teal-Streifen ganz unten
        p.fillRect(0, int(H) - 2, int(W), 2, QBrush(_RING1))

        p.end()

    # ------------------------------------------------------------------
    def set_status(self, message: str):
        self._status = message
        QApplication.processEvents()

    def finish(self, main_window=None):
        self._timer.stop()
        self.close()