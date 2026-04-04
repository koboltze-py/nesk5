"""
gui/slot_symbols.py  –  Handgezeichnete Alice-im-Wunderland Symbole
Alle Figuren originalgetreu mit QPainter — keine Fremdgrafiken notwendig.

Koordinaten: draw_symbol(p, idx, cx, cy, size)
  cx/cy = Mittelpunkt der Zelle  |  size ≈ halbe Symbolhöhe (~30 px)
"""
from __future__ import annotations

import math
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import (
    QPainter, QPen, QBrush, QColor, QPainterPath,
    QFont, QLinearGradient, QRadialGradient,
)


# ── Hilfsfunktionen ────────────────────────────────────────────────────────────
def _pth(*pts: QPointF) -> QPainterPath:
    path = QPainterPath()
    path.moveTo(pts[0])
    for pt in pts[1:]:
        path.lineTo(pt)
    path.closeSubpath()
    return path


def _star(cx: float, cy: float, ro: float, ri: float, n: int = 5) -> QPainterPath:
    path = QPainterPath()
    for i in range(2 * n):
        angle = math.pi * i / n - math.pi / 2
        r = ro if i % 2 == 0 else ri
        pt = QPointF(cx + r * math.cos(angle), cy + r * math.sin(angle))
        if i == 0:
            path.moveTo(pt)
        else:
            path.lineTo(pt)
    path.closeSubpath()
    return path


def _solid(p: QPainter, color: str) -> None:
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QBrush(QColor(color)))


def _heart(cx: float, cy: float, s: float) -> QPainterPath:
    path = QPainterPath()
    path.moveTo(cx, cy + s * 0.8)
    path.cubicTo(cx - s * 1.5, cy + s * 0.1,
                 cx - s * 1.5, cy - s * 0.7, cx, cy - s * 0.1)
    path.cubicTo(cx + s * 1.5, cy - s * 0.7,
                 cx + s * 1.5, cy + s * 0.1, cx, cy + s * 0.8)
    return path


# ══════════════════════════════════════════════════════════════════════════════
#  0 – Wild: Magischer Zauberstab  (glühend, goldener Stern)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_wild(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Stab (dunkelbraun → gold Verlauf)
    grad = QLinearGradient(cx - s * 0.55, cy + s * 0.75,
                           cx + s * 0.1,  cy - s * 0.3)
    grad.setColorAt(0.0, QColor("#6d4c41"))
    grad.setColorAt(1.0, QColor("#c9a227"))
    p.setPen(QPen(QBrush(grad), max(3.0, s * 0.16),
                  Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawLine(QPointF(cx - s * 0.55, cy + s * 0.75),
               QPointF(cx + s * 0.1,  cy - s * 0.3))

    # Glow-Ring
    grd = QRadialGradient(cx + s * 0.16, cy - s * 0.44, s * 0.7)
    grd.setColorAt(0.0, QColor(255, 220, 80, 120))
    grd.setColorAt(1.0, QColor(255, 220, 80, 0))
    p.setBrush(QBrush(grd))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawEllipse(QPointF(cx + s * 0.16, cy - s * 0.44), s * 0.7, s * 0.7)

    # Gold-Stern
    _solid(p, "#f9ca24")
    p.drawPath(_star(cx + s * 0.16, cy - s * 0.44, s * 0.52, s * 0.22, 5))

    # Funken
    for dx, dy, r in [(-0.7, -0.62, 0.09), (0.7, -0.75, 0.07), (0.6, 0.05, 0.07)]:
        grd2 = QRadialGradient(cx + dx * s, cy + dy * s, s * r * 2.5)
        grd2.setColorAt(0, QColor(255, 255, 200, 230))
        grd2.setColorAt(1, QColor(255, 255, 200, 0))
        p.setBrush(QBrush(grd2))
        p.drawEllipse(QPointF(cx + dx * s, cy + dy * s), s * r * 2.5, s * r * 2.5)


# ══════════════════════════════════════════════════════════════════════════════
#  1 – Alice  (Scatter)  blauweißes Kleid, goldene Haare, schwarzes Haarband
# ══════════════════════════════════════════════════════════════════════════════
def _draw_alice(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Blaues Kleid
    _solid(p, "#1565c0")
    p.drawPath(_pth(
        QPointF(cx - s * 0.58, cy + s * 0.85),
        QPointF(cx + s * 0.58, cy + s * 0.85),
        QPointF(cx + s * 0.32, cy + s * 0.05),
        QPointF(cx - s * 0.32, cy + s * 0.05),
    ))
    # Weißes Schürze
    _solid(p, "#eceff1")
    p.drawPath(_pth(
        QPointF(cx - s * 0.2,  cy + s * 0.85),
        QPointF(cx + s * 0.2,  cy + s * 0.85),
        QPointF(cx + s * 0.13, cy + s * 0.15),
        QPointF(cx - s * 0.13, cy + s * 0.15),
    ))
    # Kragen
    _solid(p, "#eceff1")
    p.drawPath(_pth(
        QPointF(cx - s * 0.32, cy + s * 0.05),
        QPointF(cx + s * 0.32, cy + s * 0.05),
        QPointF(cx + s * 0.2,  cy - s * 0.12),
        QPointF(cx - s * 0.2,  cy - s * 0.12),
    ))
    # Goldene Haare (hinter Kopf)
    _solid(p, "#f9a825")
    p.drawEllipse(QPointF(cx,            cy - s * 0.52), s * 0.32, s * 0.3)
    p.drawEllipse(QPointF(cx - s * 0.22, cy - s * 0.35), s * 0.2,  s * 0.28)
    p.drawEllipse(QPointF(cx + s * 0.22, cy - s * 0.35), s * 0.2,  s * 0.28)
    # Gesicht
    _solid(p, "#fce4d6")
    p.drawEllipse(QPointF(cx, cy - s * 0.26), s * 0.28, s * 0.3)
    # Schwarzes Haarband
    _solid(p, "#1a1a2e")
    p.drawPath(_pth(
        QPointF(cx - s * 0.35, cy - s * 0.50),
        QPointF(cx + s * 0.35, cy - s * 0.50),
        QPointF(cx + s * 0.35, cy - s * 0.40),
        QPointF(cx - s * 0.35, cy - s * 0.40),
    ))
    # Augen
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx - s * 0.1,  cy - s * 0.27), s * 0.045, s * 0.045)
    p.drawEllipse(QPointF(cx + s * 0.1,  cy - s * 0.27), s * 0.045, s * 0.045)
    # Lächeln
    p.setPen(QPen(QColor("#e57373"), max(1.0, s * 0.075)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawArc(QRectF(cx - s * 0.12, cy - s * 0.22, s * 0.24, s * 0.15),
              210 * 16, 120 * 16)
    # Scatter-Krone (goldener Punkt oben)
    _solid(p, "#f9ca24")
    p.drawEllipse(QPointF(cx, cy - s * 0.85), s * 0.1, s * 0.1)


# ══════════════════════════════════════════════════════════════════════════════
#  2 – Hutmacher  (sehr hoher Hut, verrückter Ausdruck, Fliege)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_hatter(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Hutturm (sehr hoch, leicht schief)
    _solid(p, "#1a1a2e")
    p.drawPath(_pth(
        QPointF(cx - s * 0.33, cy - s * 0.08),
        QPointF(cx + s * 0.38, cy - s * 0.08),
        QPointF(cx + s * 0.32, cy - s * 0.95),
        QPointF(cx - s * 0.28, cy - s * 0.95),
    ))
    # Hutrand
    p.drawPath(_pth(
        QPointF(cx - s * 0.58, cy - s * 0.08),
        QPointF(cx + s * 0.62, cy - s * 0.08),
        QPointF(cx + s * 0.55, cy + s * 0.06),
        QPointF(cx - s * 0.5,  cy + s * 0.06),
    ))
    # Grünes/türkisenes Band
    _solid(p, "#00897b")
    p.drawPath(_pth(
        QPointF(cx - s * 0.33, cy - s * 0.16),
        QPointF(cx + s * 0.38, cy - s * 0.16),
        QPointF(cx + s * 0.37, cy - s * 0.33),
        QPointF(cx - s * 0.32, cy - s * 0.33),
    ))
    # Preisschild "10/6" auf Hut
    _solid(p, "#fffde7")
    p.drawRect(QRectF(cx + s * 0.01, cy - s * 0.76, s * 0.28, s * 0.2))
    p.setPen(QPen(QColor("#1a1a2e"), max(1.0, s * 0.06)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    f = QFont("Segoe UI", max(4, int(s * 0.17)), QFont.Weight.Bold)
    p.setFont(f)
    p.drawText(QRectF(cx + s * 0.01, cy - s * 0.76, s * 0.28, s * 0.2),
               Qt.AlignmentFlag.AlignCenter, "10/6")
    # Gesicht
    _solid(p, "#fce4d6")
    p.drawEllipse(QPointF(cx, cy + s * 0.3), s * 0.35, s * 0.3)
    # Große Augen (verrückt, unterschiedlich groß)
    _solid(p, "#ffffff")
    p.drawEllipse(QPointF(cx - s * 0.14, cy + s * 0.27), s * 0.1, s * 0.1)
    p.drawEllipse(QPointF(cx + s * 0.14, cy + s * 0.27), s * 0.08, s * 0.08)
    _solid(p, "#5c6bc0")
    p.drawEllipse(QPointF(cx - s * 0.14, cy + s * 0.27), s * 0.06, s * 0.06)
    _solid(p, "#7e57c2")
    p.drawEllipse(QPointF(cx + s * 0.14, cy + s * 0.27), s * 0.05, s * 0.05)
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx - s * 0.14, cy + s * 0.27), s * 0.03, s * 0.03)
    p.drawEllipse(QPointF(cx + s * 0.14, cy + s * 0.27), s * 0.025, s * 0.025)
    # Breites Grinsen
    p.setPen(QPen(QColor("#c62828"), max(1.5, s * 0.09)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawArc(QRectF(cx - s * 0.22, cy + s * 0.28, s * 0.44, s * 0.22),
              205 * 16, 130 * 16)
    # Fliege (rot)
    _solid(p, "#c62828")
    p.drawPath(_pth(
        QPointF(cx - s * 0.22, cy + s * 0.57),
        QPointF(cx - s * 0.06, cy + s * 0.62),
        QPointF(cx - s * 0.06, cy + s * 0.55),
    ))
    p.drawPath(_pth(
        QPointF(cx + s * 0.22, cy + s * 0.57),
        QPointF(cx + s * 0.06, cy + s * 0.62),
        QPointF(cx + s * 0.06, cy + s * 0.55),
    ))
    _solid(p, "#b71c1c")
    p.drawEllipse(QPointF(cx, cy + s * 0.59), s * 0.07, s * 0.07)


# ══════════════════════════════════════════════════════════════════════════════
#  3 – Grinsekatze  (violett-rosa gestreift, riesiges Lächeln)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_cat(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Körper
    _solid(p, "#7b1fa2")
    p.drawEllipse(QPointF(cx, cy + s * 0.42), s * 0.5, s * 0.42)
    # Kopf (großes violettes Oval)
    _solid(p, "#8e24aa")
    p.drawEllipse(QPointF(cx, cy - s * 0.1), s * 0.58, s * 0.52)
    # Rosa Streifen
    p.setPen(QPen(QColor("#f48fb1"), max(1.5, s * 0.1)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    for dy in [-0.28, -0.08, 0.12]:
        p.drawLine(QPointF(cx - s * 0.48, cy + dy * s),
                   QPointF(cx + s * 0.48, cy + dy * s))
    # Ohren
    _solid(p, "#6a1b9a")
    p.drawPath(_pth(
        QPointF(cx - s * 0.56, cy - s * 0.52),
        QPointF(cx - s * 0.26, cy - s * 0.78),
        QPointF(cx - s * 0.04, cy - s * 0.5),
    ))
    p.drawPath(_pth(
        QPointF(cx + s * 0.56, cy - s * 0.52),
        QPointF(cx + s * 0.26, cy - s * 0.78),
        QPointF(cx + s * 0.04, cy - s * 0.5),
    ))
    # Innen-Ohren
    _solid(p, "#f48fb1")
    p.drawPath(_pth(
        QPointF(cx - s * 0.5,  cy - s * 0.52),
        QPointF(cx - s * 0.27, cy - s * 0.7),
        QPointF(cx - s * 0.1,  cy - s * 0.52),
    ))
    p.drawPath(_pth(
        QPointF(cx + s * 0.5,  cy - s * 0.52),
        QPointF(cx + s * 0.27, cy - s * 0.7),
        QPointF(cx + s * 0.1,  cy - s * 0.52),
    ))
    # RIESIGES Lächeln
    p.setPen(QPen(QColor("#f8bbd9"), max(2.5, s * 0.13)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawArc(QRectF(cx - s * 0.46, cy - s * 0.3, s * 0.92, s * 0.44),
              212 * 16, 116 * 16)
    # Zähne
    _solid(p, "#fafafa")
    for tx in [-0.2, -0.07, 0.07, 0.2]:
        p.drawPath(_pth(
            QPointF(cx + tx * s - s * 0.06,  cy + s * 0.02),
            QPointF(cx + tx * s + s * 0.06,  cy + s * 0.02),
            QPointF(cx + tx * s,              cy + s * 0.1),
        ))
    # Grüne Augen
    _solid(p, "#aed581")
    p.drawEllipse(QPointF(cx - s * 0.24, cy - s * 0.19), s * 0.13, s * 0.12)
    p.drawEllipse(QPointF(cx + s * 0.24, cy - s * 0.19), s * 0.13, s * 0.12)
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx - s * 0.24, cy - s * 0.19), s * 0.05, s * 0.08)
    p.drawEllipse(QPointF(cx + s * 0.24, cy - s * 0.19), s * 0.05, s * 0.08)
    # Schnurrhaare
    p.setPen(QPen(QColor("#e0e0e0"), max(1.0, s * 0.04)))
    for mx, ey_off in [(-0.7, 0.0), (-0.7, 0.06), (0.7, 0.0), (0.7, 0.06)]:
        p.drawLine(QPointF(cx + mx * s, cy + ey_off * s),
                   QPointF(cx + (0.1 if mx > 0 else -0.1) * s, cy + ey_off * s))


# ══════════════════════════════════════════════════════════════════════════════
#  4 – Weißes Kaninchen  (weißes Fell, rote Augen, goldene Taschenuhr)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_rabbit(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Ohren
    _solid(p, "#f5f5f5")
    p.drawEllipse(QPointF(cx - s * 0.2, cy - s * 0.72), s * 0.14, s * 0.44)
    p.drawEllipse(QPointF(cx + s * 0.2, cy - s * 0.72), s * 0.14, s * 0.44)
    _solid(p, "#fce4ec")
    p.drawEllipse(QPointF(cx - s * 0.2, cy - s * 0.72), s * 0.08, s * 0.32)
    p.drawEllipse(QPointF(cx + s * 0.2, cy - s * 0.72), s * 0.08, s * 0.32)
    # Körper
    _solid(p, "#f5f5f5")
    p.drawEllipse(QPointF(cx, cy + s * 0.38), s * 0.44, s * 0.45)
    # Kopf
    p.drawEllipse(QPointF(cx, cy - s * 0.1), s * 0.38, s * 0.34)
    # Kragen / Weste (blau)
    _solid(p, "#1565c0")
    p.drawPath(_pth(
        QPointF(cx - s * 0.3,  cy + s * 0.05),
        QPointF(cx + s * 0.3,  cy + s * 0.05),
        QPointF(cx + s * 0.25, cy + s * 0.28),
        QPointF(cx - s * 0.25, cy + s * 0.28),
    ))
    # Rote Augen
    _solid(p, "#ef5350")
    p.drawEllipse(QPointF(cx - s * 0.14, cy - s * 0.12), s * 0.08, s * 0.08)
    p.drawEllipse(QPointF(cx + s * 0.14, cy - s * 0.12), s * 0.08, s * 0.08)
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx - s * 0.14, cy - s * 0.12), s * 0.04, s * 0.04)
    p.drawEllipse(QPointF(cx + s * 0.14, cy - s * 0.12), s * 0.04, s * 0.04)
    # Nase
    _solid(p, "#f06292")
    p.drawEllipse(QPointF(cx, cy - s * 0.03), s * 0.05, s * 0.04)
    # Taschenuhr (goldener Ring)
    _solid(p, "#f9ca24")
    p.drawEllipse(QPointF(cx + s * 0.3, cy + s * 0.4), s * 0.25, s * 0.25)
    _solid(p, "#fffde7")
    p.drawEllipse(QPointF(cx + s * 0.3, cy + s * 0.4), s * 0.18, s * 0.18)
    # Zeiger
    p.setPen(QPen(QColor("#1a1a2e"), max(1.2, s * 0.06)))
    p.drawLine(QPointF(cx + s * 0.3, cy + s * 0.4),
               QPointF(cx + s * 0.3, cy + s * 0.26))  # Minutenzeiger
    p.drawLine(QPointF(cx + s * 0.3, cy + s * 0.4),
               QPointF(cx + s * 0.42, cy + s * 0.44))  # Stundenzeiger
    # Uhraufzugs-Knopf
    _solid(p, "#f9ca24")
    p.drawRect(QRectF(cx + s * 0.27, cy + s * 0.13, s * 0.06, s * 0.05))


# ══════════════════════════════════════════════════════════════════════════════
#  5 – Herzkönigin  (rote Krone, wütende Augen, rotes Herz)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_queen(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Kronenbasis
    _solid(p, "#c62828")
    p.drawPath(_pth(
        QPointF(cx - s * 0.58, cy - s * 0.02),
        QPointF(cx - s * 0.58, cy - s * 0.55),
        QPointF(cx - s * 0.35, cy - s * 0.3),
        QPointF(cx,             cy - s * 0.72),
        QPointF(cx + s * 0.35, cy - s * 0.3),
        QPointF(cx + s * 0.58, cy - s * 0.55),
        QPointF(cx + s * 0.58, cy - s * 0.02),
    ))
    # Kronenband
    _solid(p, "#b71c1c")
    p.drawRect(QRectF(cx - s * 0.58, cy - s * 0.18, s * 1.16, s * 0.18))
    # Juwelen
    for jx, jc in [(-0.32, "#ffd54f"), (0, "#b2ff59"), (0.32, "#ffd54f")]:
        _solid(p, jc)
        p.drawEllipse(QPointF(cx + jx * s, cy - s * 0.09), s * 0.09, s * 0.09)
    # Herz auf Krone
    _solid(p, "#fce4ec")
    p.drawPath(_heart(cx, cy - s * 0.48, s * 0.16))
    # Gesicht
    _solid(p, "#fce4d6")
    p.drawEllipse(QPointF(cx, cy + s * 0.3), s * 0.4, s * 0.36)
    # Wütende Augenbrauen
    p.setPen(QPen(QColor("#1a1a2e"), max(2.0, s * 0.09)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawLine(QPointF(cx - s * 0.28, cy + s * 0.16),
               QPointF(cx - s * 0.08, cy + s * 0.24))
    p.drawLine(QPointF(cx + s * 0.08, cy + s * 0.24),
               QPointF(cx + s * 0.28, cy + s * 0.16))
    # Augen
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx - s * 0.16, cy + s * 0.3), s * 0.07, s * 0.065)
    p.drawEllipse(QPointF(cx + s * 0.16, cy + s * 0.3), s * 0.07, s * 0.065)
    # Böser Mund
    p.setPen(QPen(QColor("#8b0000"), max(1.5, s * 0.08)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawArc(QRectF(cx - s * 0.2, cy + s * 0.38, s * 0.4, s * 0.12),
              0 * 16, 180 * 16)
    # Roter Mantel
    _solid(p, "#d32f2f")
    p.drawPath(_pth(
        QPointF(cx - s * 0.58, cy + s * 0.85),
        QPointF(cx + s * 0.58, cy + s * 0.85),
        QPointF(cx + s * 0.42, cy + s * 0.62),
        QPointF(cx - s * 0.42, cy + s * 0.62),
    ))
    # Herz auf Mantel
    _solid(p, "#fce4ec")
    p.drawPath(_heart(cx, cy + s * 0.65, s * 0.13))


# ══════════════════════════════════════════════════════════════════════════════
#  6 – Taschenuhr  (goldfarbenes Zifferblatt, Zeiger auf "late")
# ══════════════════════════════════════════════════════════════════════════════
def _draw_watch(p: QPainter, cx: float, cy: float, s: float) -> None:
    cy0 = cy + s * 0.08  # leicht nach unten versetzt
    # Äußerer Gold-Ring (mit Gravur-Effekt)
    grad = QRadialGradient(cx - s * 0.2, cy0 - s * 0.2, s * 0.7)
    grad.setColorAt(0.0, QColor("#ffe082"))
    grad.setColorAt(0.7, QColor("#f9a825"))
    grad.setColorAt(1.0, QColor("#c67c00"))
    p.setBrush(QBrush(grad))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawEllipse(QPointF(cx, cy0), s * 0.62, s * 0.62)
    # Inneres Zifferblatt
    grad2 = QRadialGradient(cx, cy0, s * 0.5)
    grad2.setColorAt(0.0, QColor("#fffde7"))
    grad2.setColorAt(1.0, QColor("#fff8e1"))
    p.setBrush(QBrush(grad2))
    p.drawEllipse(QPointF(cx, cy0), s * 0.5, s * 0.5)
    # Stundenmarkierungen
    p.setPen(QPen(QColor("#6d4c41"), max(1.5, s * 0.065)))
    for i in range(12):
        angle = math.pi * 2 * i / 12 - math.pi / 2
        x1 = cx  + s * 0.38 * math.cos(angle)
        y1 = cy0 + s * 0.38 * math.sin(angle)
        x2 = cx  + s * 0.46 * math.cos(angle)
        y2 = cy0 + s * 0.46 * math.sin(angle)
        p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
    # Minutenzeiger (zeigt auf ~12)
    p.setPen(QPen(QColor("#1a1a2e"), max(2.0, s * 0.09),
                  Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    ang_m = - math.pi / 2   # 12 Uhr
    p.drawLine(QPointF(cx, cy0),
               QPointF(cx + s * 0.38 * math.cos(ang_m),
                       cy0 + s * 0.38 * math.sin(ang_m)))
    # Stundenzeiger (10 Uhr = "it's late!")
    p.setPen(QPen(QColor("#4e342e"), max(2.5, s * 0.11),
                  Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    ang_h = math.pi * 2 * 10 / 12 - math.pi / 2
    p.drawLine(QPointF(cx, cy0),
               QPointF(cx + s * 0.28 * math.cos(ang_h),
                       cy0 + s * 0.28 * math.sin(ang_h)))
    # Mittelpunkt
    _solid(p, "#1a1a2e")
    p.drawEllipse(QPointF(cx, cy0), s * 0.06, s * 0.06)
    # Aufzieh-Knopf oben
    _solid(p, "#f9a825")
    p.drawRect(QRectF(cx - s * 0.09, cy0 - s * 0.72, s * 0.18, s * 0.14))
    # Kette-Clip
    _solid(p, "#c67c00")
    p.drawEllipse(QPointF(cx, cy0 - s * 0.66), s * 0.07, s * 0.07)


# ══════════════════════════════════════════════════════════════════════════════
#  7 – Wunderpilz  (roter Hut, weiße Punkte, weißer Stiel)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_shroom(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Stiel
    _solid(p, "#f5f5dc")
    p.drawPath(_pth(
        QPointF(cx - s * 0.2,  cy + s * 0.88),
        QPointF(cx + s * 0.2,  cy + s * 0.88),
        QPointF(cx + s * 0.17, cy + s * 0.2),
        QPointF(cx - s * 0.17, cy + s * 0.2),
    ))
    # Stielrand
    _solid(p, "#fafafa")
    p.drawEllipse(QPointF(cx, cy + s * 0.22), s * 0.3, s * 0.1)
    # Roter Hutkappe
    _solid(p, "#d32f2f")
    cap = QPainterPath()
    cap.moveTo(cx - s * 0.65, cy + s * 0.22)
    cap.cubicTo(cx - s * 0.72, cy - s * 0.42,
                cx - s * 0.08, cy - s * 0.95,
                cx,            cy - s * 0.92)
    cap.cubicTo(cx + s * 0.08, cy - s * 0.95,
                cx + s * 0.72, cy - s * 0.42,
                cx + s * 0.65, cy + s * 0.22)
    cap.closeSubpath()
    p.drawPath(cap)
    # Weiße Punkte
    _solid(p, "#f5f5f5")
    for dx, dy, r in [(0, -0.5, 0.14), (-0.34, -0.22, 0.1), (0.34, -0.22, 0.1),
                      (-0.18, -0.72, 0.09), (0.22, -0.72, 0.09)]:
        p.drawEllipse(QPointF(cx + dx * s, cy + dy * s), s * r, s * r)
    # Hutunterseite (Schatten)
    _solid(p, "#b71c1c")
    p.drawEllipse(QPointF(cx, cy + s * 0.22), s * 0.65, s * 0.11)


# ══════════════════════════════════════════════════════════════════════════════
#  8 – Teekanne  (blau-weiß bemalt, Dampf, Verrückte-Hutmacher-Party)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_teapot(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Kannenkörper
    grad = QLinearGradient(cx - s * 0.5, cy, cx + s * 0.5, cy)
    grad.setColorAt(0.0, QColor("#1e88e5"))
    grad.setColorAt(1.0, QColor("#0d47a1"))
    p.setBrush(QBrush(grad))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawEllipse(QPointF(cx, cy + s * 0.18), s * 0.5, s * 0.54)
    # Weiße Chinadekor-Punkte
    _solid(p, "#e3f2fd")
    for dx, dy in [(-0.18, -0.06), (0.18, -0.06), (0, 0.16),
                   (-0.28, 0.22), (0.28, 0.22)]:
        p.drawEllipse(QPointF(cx + dx * s, cy + s * 0.18 + dy * s), s * 0.055, s * 0.055)
    # Schnabel
    _solid(p, "#1565c0")
    p.drawPath(_pth(
        QPointF(cx + s * 0.46, cy - s * 0.08),
        QPointF(cx + s * 0.95, cy - s * 0.45),
        QPointF(cx + s * 0.88, cy - s * 0.22),
        QPointF(cx + s * 0.46, cy + s * 0.12),
    ))
    # Henkel
    p.setPen(QPen(QColor("#1565c0"), max(4.0, s * 0.2),
                  Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawArc(QRectF(cx - s * 0.9, cy - s * 0.1, s * 0.42, s * 0.56),
              300 * 16, 120 * 16)
    # Deckel
    _solid(p, "#1e88e5")
    p.drawEllipse(QPointF(cx, cy - s * 0.6), s * 0.32, s * 0.1)
    _solid(p, "#0d47a1")
    p.drawEllipse(QPointF(cx, cy - s * 0.7), s * 0.09, s * 0.1)  # Knauf
    # Dampf-Wellchen
    p.setPen(QPen(QColor(200, 230, 255, 180), max(1.5, s * 0.07)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    for dxd in [-0.1, 0.05, 0.2]:
        p.drawArc(QRectF(cx + dxd * s - s * 0.02, cy - s * 0.98,
                         s * 0.18, s * 0.2),
                  0 * 16, 180 * 16)


# ══════════════════════════════════════════════════════════════════════════════
#  9 – Schlüssel  (verzierter Gold-Schlüssel)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_key(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Gradient-Gold
    grad = QLinearGradient(cx - s * 0.15, 0, cx + s * 0.15, 0)
    grad.setColorAt(0.0, QColor("#ffe082"))
    grad.setColorAt(0.5, QColor("#f9a825"))
    grad.setColorAt(1.0, QColor("#c67c00"))
    # Schaft
    p.setBrush(QBrush(grad))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRect(QRectF(cx - s * 0.09, cy - s * 0.08, s * 0.18, s * 0.9))
    # Zähne
    for by, bh in [(cy + s * 0.42, s * 0.12), (cy + s * 0.6, s * 0.12)]:
        p.drawRect(QRectF(cx + s * 0.09, by, s * 0.2, bh))
    # Ornate Kopf (großer Ring)
    p.drawEllipse(QPointF(cx, cy - s * 0.42), s * 0.44, s * 0.44)
    # Loch im Kopf
    _solid(p, "#0d0422")   # Hintergund-Farbe
    p.drawEllipse(QPointF(cx, cy - s * 0.42), s * 0.26, s * 0.26)
    # Kreuz im Ring
    _solid(p, "#f9a825")
    p.drawRect(QRectF(cx - s * 0.28, cy - s * 0.47, s * 0.56, s * 0.1))
    p.drawRect(QRectF(cx - s * 0.05, cy - s * 0.7,  s * 0.1,  s * 0.56))
    # Outline
    p.setPen(QPen(QColor("#c67c00"), max(1.0, s * 0.05)))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawEllipse(QPointF(cx, cy - s * 0.42), s * 0.44, s * 0.44)


# ══════════════════════════════════════════════════════════════════════════════
#  10 – Rote Rose  (wird rot angemalt — Herzkarte-Gärtner-Szene)
# ══════════════════════════════════════════════════════════════════════════════
def _draw_rose(p: QPainter, cx: float, cy: float, s: float) -> None:
    # Stängel
    _solid(p, "#388e3c")
    p.drawRect(QRectF(cx - s * 0.055, cy + s * 0.12, s * 0.11, s * 0.76))
    # Blätter
    _solid(p, "#43a047")
    p.drawPath(_pth(
        QPointF(cx,            cy + s * 0.48),
        QPointF(cx - s * 0.38, cy + s * 0.3),
        QPointF(cx - s * 0.06, cy + s * 0.55),
    ))
    p.drawPath(_pth(
        QPointF(cx,            cy + s * 0.38),
        QPointF(cx + s * 0.38, cy + s * 0.2),
        QPointF(cx + s * 0.06, cy + s * 0.45),
    ))
    # Blütenblätter (layered)
    _solid(p, "#b71c1c")
    p.drawEllipse(QPointF(cx, cy - s * 0.12), s * 0.48, s * 0.46)
    _solid(p, "#e53935")
    p.drawEllipse(QPointF(cx - s * 0.16, cy - s * 0.25), s * 0.35, s * 0.32)
    p.drawEllipse(QPointF(cx + s * 0.16, cy - s * 0.25), s * 0.35, s * 0.32)
    _solid(p, "#ef5350")
    p.drawEllipse(QPointF(cx, cy - s * 0.35), s * 0.3, s * 0.3)
    # Kern
    _solid(p, "#c62828")
    p.drawEllipse(QPointF(cx, cy - s * 0.2), s * 0.16, s * 0.16)
    # Pinsel (Gärtner malt die Rose rot)
    _solid(p, "#795548")
    p.drawRect(QRectF(cx + s * 0.35, cy - s * 0.68, s * 0.08, s * 0.38))
    _solid(p, "#e53935")
    p.drawEllipse(QPointF(cx + s * 0.39, cy - s * 0.66), s * 0.1, s * 0.08)
    # Farbkleckse
    for dx, dy in [(0.52, -0.52), (0.62, -0.38)]:
        _solid(p, "#ef5350")
        p.drawEllipse(QPointF(cx + dx * s, cy + dy * s), s * 0.06, s * 0.06)


# ── Dispatch-Tabelle ──────────────────────────────────────────────────────────
_DRAW = [
    _draw_wild,     # 0  Wild
    _draw_alice,    # 1  Alice (Scatter)
    _draw_hatter,   # 2  Mad Hatter
    _draw_cat,      # 3  Cheshire Cat
    _draw_rabbit,   # 4  White Rabbit
    _draw_queen,    # 5  Queen of Hearts
    _draw_watch,    # 6  Pocket Watch
    _draw_shroom,   # 7  Mushroom
    _draw_teapot,   # 8  Teapot
    _draw_key,      # 9  Key
    _draw_rose,     # 10 Red Rose
]


def draw_symbol(p: QPainter, sym_idx: int,
                cx: float, cy: float, size: float,
                alpha: float = 1.0) -> None:
    """Zeichnet Symbol sym_idx zentriert bei (cx, cy); size = halbe Symbolhöhe."""
    if not 0 <= sym_idx < len(_DRAW):
        return
    p.save()
    p.setOpacity(p.opacity() * max(0.0, min(1.0, alpha)))
    _DRAW[sym_idx](p, cx, cy, size)
    p.restore()
