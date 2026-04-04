"""
gui/slot_machine.py  –  🎩 Alice im Wunderland — Wunderrad  v3.1
5 Reels · 3 Reihen · 243 Ways · Wild
• Einsatzstufen  5 · 10 · 20 · 50 · 100 Credits
• 3 Sammelpods   🔴 Herzkönigin · 🔵 Wunderland · ⭐ Goldschatz
• Bonusbälle     landen per Zufall – füllen Pods  
• Holding Spin   5+ Bälle → 5 Respins (Reset bei neuem Ball)
• Free Games     3× Alice → 10 FS · 2× · Sticky Wilds
• Pod-Boni       Herzkönigin-Karten · Wunderland-FS · Goldschatz-Jackpot
"""
from __future__ import annotations

import math
import random

from PySide6.QtCore import Qt, QRectF, QPointF, QTimer
from PySide6.QtGui import (
    QColor, QBrush, QPen, QPainter, QFont,
    QLinearGradient, QRadialGradient,
)
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QWidget, QFrame,
)

# ── Symbole ───────────────────────────────────────────────────────────────────
WILD_IDX   = 0
ALICE_IDX  = 1
BALL_R_IDX = 11        # ROT  → Pod 0 (Herzkönigin)
BALL_B_IDX = 12        # BLAU → Pod 1 (Wunderland)
BALL_G_IDX = 13        # GOLD → Pod 2 (Goldschatz)
IS_BALL:     frozenset[int] = frozenset({BALL_R_IDX, BALL_B_IDX, BALL_G_IDX})
BALL_TO_POD: dict[int, int] = {BALL_R_IDX: 0, BALL_B_IDX: 1, BALL_G_IDX: 2}

# (emoji, name, farbe, gewicht, win3, win4, win5)
SYMBOLS: list[tuple] = [
    ("🪄", "Wild",              "#ffffff",   4,  0,    0,  5000),  # 0
    ("👸", "Alice (Scatter)",   "#fce4ec",   5, 400, 1000, 3000),  # 1
    ("🎩", "Hutmacher",         "#f9ca24",   9, 120,  320,  800),  # 2
    ("🐱", "Grinsekatze",       "#c39bd3",  11,  80,  220,  550),  # 3
    ("🐰", "Kaninchen",         "#dfe6e9",  13,  55,  160,  380),  # 4
    ("👑", "Herzkönigin",       "#e74c3c",  15,  40,  110,  250),  # 5
    ("⏰", "Taschenuhr",        "#fab1a0",  18,  28,   75,  160),  # 6
    ("🍄", "Wunderpilz",        "#55efc4",  21,  18,   50,  100),  # 7
    ("🫖", "Teekanne",          "#74b9ff",  24,  10,   30,   70),  # 8
    ("🔑", "Schlüssel",         "#fdcb6e",  28,   7,   20,   45),  # 9
    ("🌹", "Rote Rose",         "#fd79a8",  32,   4,   12,   30),  # 10
    ("🔴", "Ball ROT",          "#ef5350",   0,   0,    0,    0),  # 11
    ("🔵", "Ball BLAU",         "#1e88e5",   0,   0,    0,    0),  # 12
    ("⭐", "Ball GOLD",         "#f9a825",   0,   0,    0,    0),  # 13
]
_POOL = sum([[i] * SYMBOLS[i][3] for i in range(11)], [])   # Indices 0-10

# Einsatz
BET_LEVELS   = [5, 10, 20, 50, 100]
_DEF_BET_IDX = 1        # 10 Credits

# Pods  (kein fester Zähler – Bälle erhitzen den Pod, Bonus löst zufällig aus)
POD_NAMES  = ["Herzkönigin", "Wunderland", "Goldschatz"]
POD_EMOJIS = ["🔴", "🔵", "⭐"]
POD_COLS   = ["#c62828", "#1565c0", "#f9a825"]
POD_TXT    = ["#ef9a9a", "#90caf9", "#ffe082"]

# Ball-Werte (Multiplikator × Bet/10) – stark auf niedrige Werte gewichtet
_BV_MUL  = [1,  2,  3,  5,  8, 15, 25, 50]
_BV_WGT  = [52, 26, 12,  6,  3,  2,  1,  1]

# Holding Spin
_HOLD_TRIGGER = 3    # Bälle auf mindestens 3 Walzen → Holding Spin (~1:80 Spins)
_HOLD_RESPINS = 3    # Standard Lightning Link: immer 3 Respins

# Jackpot-Stufen im Holding Spin (Multiplikator × Einsatz)
_JP_LABELS = ["MINI", "MINOR", "MAJOR", "GRAND"]
_JP_MULTS  = [  10,     35,     100,    350  ]
_JP_PROBS  = [ 0.08,  0.03,   0.008,  0.001 ]   # pro Ball im Hold-Spin

# Free Games
_FS_BASE    = 10
_FS_RETRIG  = 5
_FS_MULT    = 2

_ROWS  = 3
_REELS = 5


def _rnd() -> int:
    return random.choice(_POOL)


def _rnd_ball() -> int:
    return random.choice([BALL_R_IDX, BALL_B_IDX, BALL_G_IDX])


def _ball_val(bet: int) -> int:
    m = random.choices(_BV_MUL, weights=_BV_WGT, k=1)[0]
    return max(1, bet * m // 10)


def _ball_prob(bet_idx: int) -> float:
    return 0.10 + bet_idx * 0.012


def _ball_val_hold(bet: int) -> tuple[int, str | None]:
    """Ball-Wert während Hold & Respin – inkl. zufälliger Jackpot-Stufe."""
    r = random.random()
    cumulative = 0.0
    for label, mult, prob in zip(_JP_LABELS, _JP_MULTS, _JP_PROBS):
        cumulative += prob
        if r < cumulative:
            return (bet * mult, label)
    m = random.choices(_BV_MUL, weights=_BV_WGT, k=1)[0]
    return (max(1, bet * m // 10), None)


def _hold_col(locked_rows: set[int]) -> list[int]:
    """Hold-Spin Walze: jede Reihe unabhängig 28% Ball-Chance oder Rose (Blank)."""
    col: list[int] = []
    for row in range(_ROWS):
        if row in locked_rows:
            col.append(_rnd_ball())   # gesperrte Reihe – Overlay überlagert es
        else:
            col.append(_rnd_ball() if random.random() < 0.28 else 10)  # 10 = Rose als Blank
    return col


def _spin_col(bet_idx: int, extra_ball: float = 0.0) -> list[int]:
    """Erzeugt 3 Symbole für eine Walze — max. 1 Ball pro Walze."""
    bp = _ball_prob(bet_idx) + extra_ball
    syms: list[int] = []
    has_ball = False
    for _ in range(_ROWS):
        if not has_ball and random.random() < bp:
            syms.append(_rnd_ball())
            has_ball = True
        else:
            syms.append(_rnd())
    return syms


# ── 243-Ways Auswertung ───────────────────────────────────────────────────────
def evaluate_ways(grid: list[list[int]],
                  mult: float = 1.0) -> tuple[int, list[tuple]]:
    """
    grid[reel][row]  —  5 Reels × 3 Reihen.
    Balls und Scatter werden ignoriert. mult skaliert Gewinne mit Einsatz.
    """
    wins: list[tuple] = []
    total = 0

    candidates = {s for col in grid for s in col
                  if s not in (WILD_IDX, ALICE_IDX) and s not in IS_BALL}

    for sym in candidates:
        ways = 1
        length = 0
        for col in grid:
            m = sum(1 for s in col if s == sym or s == WILD_IDX)
            if m == 0:
                break
            ways   *= m
            length += 1
        if length >= 3:
            se = SYMBOLS[sym]
            bw = se[6] if length == 5 else se[5] if length == 4 else se[4]
            prize = max(1, int(bw * ways * mult))
            total += prize
            wins.append((sym, length, ways, prize))

    if sum(1 for col in grid if all(s == WILD_IDX for s in col)) == _REELS:
        jp = int(SYMBOLS[WILD_IDX][6] * mult)
        total += jp
        wins.append((WILD_IDX, 5, 1, jp))

    return total, wins


# ── _PodWidget ────────────────────────────────────────────────────────────────
class _PodWidget(QWidget):
    """Pod: Bälle erhitzen ihn – je heißer desto wahrscheinlicher der Bonus."""

    def __init__(self, idx: int, parent=None) -> None:
        super().__init__(parent)
        self.idx         = idx
        self._heat       = 0.0   # 0.0 – 1.0
        self._lit        = False
        self._pulse      = 0.0
        self._pulse_dir  = 1
        self.setFixedSize(118, 54)
        self.setToolTip(
            f"{POD_EMOJIS[idx]} {POD_NAMES[idx]}\n"
            "Bälle dieser Farbe erhitzen den Pod.\n"
            "Je heißer, desto wahrscheinlicher löst er aus!"
        )
        self._ptimer = QTimer(self)
        self._ptimer.timeout.connect(self._do_pulse)
        self._ptimer.start(40)

    def _do_pulse(self) -> None:
        if self._heat < 0.04 and not self._lit:
            return
        spd = 0.03 + self._heat * 0.09
        self._pulse += spd * self._pulse_dir
        if self._pulse >= 1.0:
            self._pulse = 1.0
            self._pulse_dir = -1
        elif self._pulse <= 0.0:
            self._pulse = 0.0
            self._pulse_dir = 1
        self.update()

    @property
    def heat(self) -> float:
        return self._heat

    @heat.setter
    def heat(self, v: float) -> None:
        self._heat = max(0.0, min(1.0, float(v)))
        self.update()

    def flash_trigger(self) -> None:
        """Kurzes Aufleuchten wenn Bonus ausgelöst wird."""
        self._lit = True
        self.update()
        QTimer.singleShot(2400, self._unlit)

    def _unlit(self) -> None:
        self._lit = False
        self.update()

    def reset(self) -> None:
        self._heat      = 0.0
        self._lit       = False
        self._pulse     = 0.0
        self._pulse_dir = 1
        self.update()

    def paintEvent(self, _) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        W, H   = self.width(), self.height()
        base   = QColor(POD_COLS[self.idx])
        txt_c  = QColor(POD_TXT[self.idx])
        glow   = self._pulse * self._heat

        # Hintergrund-Glow
        bg = QColor(base)
        bg.setAlpha(200 if self._lit else int(12 + self._heat * 100 + glow * 60))
        p.fillRect(0, 0, W, H, QBrush(bg))

        # Rand – pulsiert
        bw = 1.0 + self._heat * 2.5 + glow * 2.0
        bc = QColor(base)
        bc.setAlpha(int(100 + self._heat * 155 + glow * 60))
        p.setPen(QPen(bc, bw))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(QRectF(1, 1, W - 2, H - 2), 7, 7)

        # Name
        f = QFont("Segoe UI", 7, QFont.Weight.Bold)
        p.setFont(f)
        p.setPen(QPen(txt_c))
        p.drawText(QRectF(2, 2, W - 4, 14), Qt.AlignmentFlag.AlignCenter,
                   f"{POD_EMOJIS[self.idx]} {POD_NAMES[self.idx]}")

        # Status-Text
        if self._lit:
            status = "💥 BONUS!"
        elif self._heat >= 0.75:
            status = "🔥 HEISS!"
        elif self._heat >= 0.4:
            status = f"♨ {int(self._heat * 100)}%"
        elif self._heat > 0.0:
            status = f"· {int(self._heat * 100)}%"
        else:
            status = "kalt"
        f2 = QFont("Segoe UI", 7, QFont.Weight.Bold)
        p.setFont(f2)
        p.setPen(QPen(txt_c))
        p.drawText(QRectF(2, 17, W - 4, 13), Qt.AlignmentFlag.AlignCenter, status)

        # Wärme-Balken
        bx, by, bh2, bww = 5.0, 33.0, 8.0, float(W - 10)
        bg2 = QColor(base); bg2.setAlpha(28)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QBrush(bg2))
        p.drawRoundedRect(QRectF(bx, by, bww, bh2), 3, 3)
        if self._heat > 0.0:
            fill_c = QColor(base); fill_c.setAlpha(int(160 + glow * 90))
            p.setBrush(QBrush(fill_c))
            p.drawRoundedRect(QRectF(bx, by, bww * self._heat, bh2), 3, 3)

        # Glüh-Halo bei hoher Hitze
        if self._heat >= 0.5:
            halo = QColor(base)
            halo.setAlpha(int(glow * self._heat * 55))
            p.setBrush(QBrush(halo))
            halo_r = 6.0 + glow * 8.0
            p.drawEllipse(QPointF(W / 2.0, by + bh2 / 2), bww * self._heat / 2.0 + halo_r, bh2 + halo_r)
        p.end()


# ── _WinOverlay ───────────────────────────────────────────────────────────────
class _WinOverlay(QWidget):
    """Animiertes Bonus-Win-Banner — erscheint zentriert über dem Dialog."""

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setFixedSize(parent.width(), parent.height())
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._text:  str    = ""
        self._color: QColor = QColor("#c9a227")
        self._alpha: float  = 0.0
        self._ring:  float  = 0.0
        self._phase: str    = "out"
        self._hld:   int    = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._step)
        self.hide()

    def show_anim(self, text: str, color: str = "#c9a227") -> None:
        self._text  = text
        self._color = QColor(color)
        self._alpha = 0.0
        self._ring  = 0.0
        self._phase = "in"
        self._hld   = 0
        self.raise_()
        self.show()
        if not self._timer.isActive():
            self._timer.start(22)

    def _step(self) -> None:
        if self._phase == "in":
            self._alpha = min(1.0, self._alpha + 0.08)
            self._ring  = min(1.0, self._ring  + 0.05)
            if self._alpha >= 1.0:
                self._phase = "hold"
        elif self._phase == "hold":
            self._ring = (self._ring + 0.025) % 1.0
            self._hld += 1
            if self._hld >= 90:
                self._phase = "out"
        elif self._phase == "out":
            self._alpha = max(0.0, self._alpha - 0.06)
            if self._alpha <= 0.0:
                self._timer.stop()
                self.hide()
        self.update()

    def paintEvent(self, _) -> None:
        if self._alpha <= 0.0:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        W, H  = self.width(), self.height()
        cx    = W / 2.0
        cy    = H * 0.42
        bw    = W - 80
        bh    = 86

        # Hintergrund
        p.setOpacity(self._alpha * 0.88)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QBrush(QColor(4, 0, 18, 245)))
        p.drawRoundedRect(QRectF(cx - bw / 2, cy - bh / 2, bw, bh), 16, 16)

        # Leuchtender Rand
        c_pen = QColor(self._color)
        c_pen.setAlpha(int(self._alpha * 210))
        p.setPen(QPen(c_pen, 2.0))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(QRectF(cx - bw / 2, cy - bh / 2, bw, bh), 16, 16)

        # Expandier-Ring
        ring_r    = 30 + self._ring * 230
        c_ring    = QColor(self._color)
        ring_fade = max(0.0, 1.0 - self._ring) * self._alpha
        c_ring.setAlpha(int(ring_fade * 110))
        p.setPen(QPen(c_ring, max(1.0, 4.5 * (1.0 - self._ring))))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawEllipse(QPointF(cx, cy), ring_r, ring_r)

        # Zweiter, versetzter Ring
        ring_r2   = 30 + ((self._ring + 0.4) % 1.0) * 230
        ring_fade2 = max(0.0, 1.0 - (self._ring + 0.4) % 1.0) * self._alpha * 0.6
        c_ring.setAlpha(int(ring_fade2 * 90))
        p.setPen(QPen(c_ring, max(1.0, 3.0 * (1.0 - (self._ring + 0.4) % 1.0))))
        p.drawEllipse(QPointF(cx, cy), ring_r2, ring_r2)

        # Text (Schatten + Hauptfarbe)
        p.setOpacity(self._alpha)
        f = QFont("Segoe UI", 17, QFont.Weight.Bold)
        p.setFont(f)
        p.setPen(QPen(QColor(0, 0, 0, 170)))
        p.drawText(QRectF(cx - bw / 2 + 2, cy - bh / 2 + 2, bw, bh),
                   Qt.AlignmentFlag.AlignCenter, self._text)
        p.setPen(QPen(self._color))
        p.drawText(QRectF(cx - bw / 2, cy - bh / 2, bw, bh),
                   Qt.AlignmentFlag.AlignCenter, self._text)
        p.end()


# ══════════════════════════════════════════════════════════════════════════════
#  _Reel  –  eine Walze (zeigt 3 Reihen)
# ══════════════════════════════════════════════════════════════════════════════
class _Reel(QWidget):
    _SH = 78    # Symbol-Höhe px
    _N  = 24    # Bandlänge

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFixedSize(88, self._SH * _ROWS)
        self._band:   list[int]    = [_rnd() for _ in range(self._N)]
        self._pos:    float        = float(random.randint(0, self._N - 1))
        self._speed:  float        = 0.0
        self._target: float | None = None
        self.stopped: bool         = True
        self._flash_rows: set[int] = set()   # welche Zeilen blinken
        self._flash:  float        = 0.0
        self.held:    bool         = False
        self._ht:     float        = 0.0
        self._ball_vals:    list[int | None]  = [None, None, None]
        self._ball_labels:  list[str | None]  = [None, None, None]  # JP-Label
        self.hold_mode:     bool              = False  # Hold-Spin-Modus: nicht-Bälle abdunkeln
        self.sweat:   bool         = False   # Sweat-Walze (Free Games)
        self._sw:     float        = 0.0    # Animations-Phase
        # Per-Zell-Animation (jede Zelle = unabhängiges 1x1-Reel im Hold & Spin)
        self._cell_phase:   list[float] = [0.0, 0.0, 0.0]  # -1=pending >0=dreht 0=fertig
        self._cell_target:  list[int]   = [-1,  -1,  -1  ]  # Ziel-Symbol pro Zeile

    # ── Sichtbare Symbole ──────────────────────────────────────────────────────
    def visible_symbols(self) -> list[int]:
        """Gibt die 3 sichtbaren Symbole zurück (oben→unten)."""
        base = int(self._pos) % self._N
        return [self._band[(base + r) % self._N] for r in range(_ROWS)]

    # ── Steuerung ──────────────────────────────────────────────────────────────
    def start(self) -> None:
        if self.held:
            return
        self.stopped      = False
        self._speed       = 0.62
        self._target      = None
        self._flash       = 0.0
        self._flash_rows  = set()

    def schedule_stop(self, syms: list[int]) -> None:
        """syms: gewünschte 3 sichtbare Symbole von oben."""
        if self.held:
            return
        # Stelle sicher alle 3 Symbole stehen auf dem Band
        for r, s in enumerate(syms):
            ok = any(self._band[(i + r) % self._N] == s
                     for i in range(self._N))
            if not ok:
                self._band[r % self._N] = s

        # Startposition suchen die zu syms[0] passt
        candidates = [i for i in range(self._N)
                      if self._band[i % self._N] == syms[0]]
        if not candidates:
            self._band[0] = syms[0]
            candidates = [0]

        cur_floor = math.floor(self._pos)
        cur_mod   = cur_floor % self._N
        best      = min(candidates,
                        key=lambda t: (t - cur_mod) % self._N)
        dist = (best - cur_mod) % self._N
        if dist < 4:
            dist += self._N
        self._target = float(cur_floor + dist + self._N)

    def flash_win(self, rows: set[int] | None = None) -> None:
        self._flash      = 1.0
        self._flash_rows = rows or {0, 1, 2}
    def set_ball_val(self, row: int, val: int | None) -> None:
        self._ball_vals[row] = val
        self.update()

    def set_ball_label(self, row: int, label: str | None) -> None:
        self._ball_labels[row] = label
        self.update()

    def clear_ball_vals(self) -> None:
        self._ball_vals   = [None, None, None]
        self._ball_labels = [None, None, None]
        self.update()
    # ── Tick ───────────────────────────────────────────────────────────────────
    def tick(self) -> bool:
        # Per-Zell-Animation – laeuft immer, auch bei held=True
        for _row in range(_ROWS):
            if self._cell_phase[_row] > 0:
                self._cell_phase[_row] -= 1.0
                if self._cell_phase[_row] <= 0:
                    self._cell_phase[_row] = 0.0
                    _base = int(self._pos) % self._N
                    self._band[(_base + _row) % self._N] = self._cell_target[_row]

        if self.held:
            self._ht += 0.09
            if self._flash > 0:
                self._flash = max(0.0, self._flash - 0.025)
            self.update()
            return True

        if self.stopped:
            if self._flash > 0:
                self._flash = max(0.0, self._flash - 0.035)
                self.update()
            return True

        if self._target is None:
            self._pos = (self._pos + self._speed) % self._N
        else:
            rem = self._target - self._pos
            if rem <= 0.02:
                self._pos    = self._target % self._N
                self._speed  = 0.0
                self.stopped = True
                self.update()
                return True
            self._speed = max(0.03, min(0.62, rem * 0.16))
            self._pos  += self._speed

        self.update()
        return False

    # ── Zeichnen ───────────────────────────────────────────────────────────────
    def paintEvent(self, _) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        W, H = self.width(), self.height()
        SH   = self._SH

        # Hintergrund
        bg = QLinearGradient(0, 0, 0, H)
        if self.held:
            bg.setColorAt(0.0, QColor("#261200"))
            bg.setColorAt(0.5, QColor("#3d1f00"))
            bg.setColorAt(1.0, QColor("#261200"))
        else:
            bg.setColorAt(0.0, QColor("#0d0422"))
            bg.setColorAt(0.5, QColor("#1a073d"))
            bg.setColorAt(1.0, QColor("#0d0422"))
        p.fillRect(0, 0, W, H, QBrush(bg))

        frac = self._pos - math.floor(self._pos)
        base = int(self._pos) % self._N

        from gui.slot_symbols import draw_symbol

        for k in range(-1, _ROWS + 1):
            sym_i = self._band[(base + k) % self._N]
            cy    = (k + 0.5 - frac) * SH

            if cy < -SH * 0.75 or cy > H + SH * 0.75:
                continue

            dist_top    = cy
            dist_bottom = H - cy
            fade = min(1.0, dist_top / (SH * 0.48), dist_bottom / (SH * 0.48))
            fade = max(0.0, fade)

            # Per-Zell-Animation ueberschreibt Symbol + Helligkeit
            sym_fade = fade
            if 0 <= k < _ROWS:
                _ph = self._cell_phase[k]
                if _ph < 0:           # pending: noch nicht gestartet
                    sym_i    = _POOL[int(self._ht * 8) % len(_POOL)]
                    sym_fade = fade * 0.28
                elif _ph > 0:         # aktiv drehend: Symbol-Zyklieren
                    sym_i    = _POOL[int(_ph * 3.5) % len(_POOL)]
                    sym_fade = fade * 0.55
                elif (self.hold_mode and sym_i not in IS_BALL
                        and self._ball_vals[k] is None):
                    sym_fade = fade * 0.15
            draw_symbol(p, sym_i, W / 2.0, cy, SH * 0.38, sym_fade)

            # Ball-Kreditwert-Badge mit optionalem Jackpot-Label
            if 0 <= k < _ROWS and self._ball_vals[k] is not None and fade > 0.25:
                bv    = self._ball_vals[k]
                label = self._ball_labels[k]
                is_jp = label is not None
                jp_colors = {"MINI": "#80cbc4", "MINOR": "#a5d6a7",
                             "MAJOR": "#fff59d", "GRAND": "#ff8a65"}
                badge_col = jp_colors.get(label, "#f9a825") if is_jp else "#f9a825"
                bw, bh    = (48, 14) if is_jp else (38, 14)
                p.setOpacity(fade * 0.95)
                p.setPen(Qt.PenStyle.NoPen)
                p.setBrush(QBrush(QColor(60, 0, 40, 210) if is_jp else QColor(0, 0, 0, 190)))
                p.drawRoundedRect(QRectF(W / 2 - bw / 2, cy + SH * 0.26, bw, bh), 4, 4)
                bf = QFont("Segoe UI", 7, QFont.Weight.Bold)
                p.setFont(bf)
                p.setPen(QPen(QColor(badge_col)))
                p.drawText(QRectF(W / 2 - bw / 2, cy + SH * 0.26, bw, bh),
                           Qt.AlignmentFlag.AlignCenter,
                           label if is_jp else f"+{bv}")

        p.setOpacity(1.0)

        # Flash-Overlay auf Gewinn-Zeilen
        if self._flash > 0:
            for row in self._flash_rows:
                fy = row * SH
                fc = QColor(201, 162, 39, int(self._flash * 140))
                p.fillRect(0, fy, W, SH, QBrush(fc))

        # HOLD-Puls
        if self.held:
            pulse = 0.55 + 0.45 * math.sin(self._ht)
            p.setOpacity(pulse)
            hf = QFont("Segoe UI", 7, QFont.Weight.Bold)
            p.setFont(hf)
            p.setPen(QPen(QColor("#c9a227")))
            p.drawText(QRectF(0, 4, W, 15),
                       Qt.AlignmentFlag.AlignHCenter, "HOLD")
            p.setOpacity(1.0)

        # SWEAT-Glut – Feuer-Überlagerung auf der ausgewählten Walze
        if self.sweat:
            sw_pulse = 0.5 + 0.5 * math.sin(self._sw)
            fg = QLinearGradient(0, 0, 0, H)
            fg.setColorAt(0.0, QColor(255, 107, 0, int(sw_pulse * 95)))
            fg.setColorAt(0.5, QColor(255, 180, 0, int(sw_pulse * 38)))
            fg.setColorAt(1.0, QColor(255, 107, 0, int(sw_pulse * 95)))
            p.setOpacity(1.0)
            p.fillRect(0, 0, W, H, QBrush(fg))
            p.setOpacity(1.0)

        # Gewinnlinien-Ticks
        p.setPen(QPen(QColor(201, 162, 39, 60), 1, Qt.PenStyle.DotLine))
        for row in range(1, _ROWS):
            p.drawLine(0, row * SH, W, row * SH)

        # Rahmen
        if self.sweat:
            sw_p = 0.5 + 0.5 * math.sin(self._sw)
            p.setPen(QPen(QColor(255, int(107 + sw_p * 80), 0, 230), 3.5))
        elif self.held:
            p.setPen(QPen(QColor("#c9a227"), 3))
        else:
            p.setPen(QPen(QColor("#6c3483"), 2))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(QRectF(1, 1, W - 2, H - 2), 7, 7)
        p.end()


# ══════════════════════════════════════════════════════════════════════════════
#  SlotMachineDialog
# ══════════════════════════════════════════════════════════════════════════════
class SlotMachineDialog(QDialog):
    """🎩 Alice im Wunderland — 5×3 · 243 Ways · Wild · Bonus Balls · Free Games"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("🎩 Alice's Wunderrad")
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(610, 830)

        # ── Basis-State ───────────────────────────────────────────────────────
        self._credits   = 200
        self._bet_idx   = _DEF_BET_IDX
        self._mode      = "idle"    # idle|spinning|freespinning|holdspinning|waiting
        self._drag_pos  = None

        # Pods
        self._pod_heat: list[float] = [0.0, 0.0, 0.0]

        # Holding Spin
        self._hold_reels: set[int]                        = set()
        self._hold_vals:  dict[int, list[tuple[int,int]]] = {}   # reel→[(row,val)]
        self._hold_jps:   dict[tuple[int,int], str | None] = {}  # (ri,row)→JP-Label
        self._hold_respins: int = 0
        self._hold_bet:     int = 0
        self._wait_ticks:   int = 0  # Watchdog-Zähler für stuck "waiting"-Mode

        # Free Games
        self._fs_left:      int              = 0
        self._fs_total:     int              = 0
        self._fs_wins:      int              = 0   # akkumulierte Credits in FG
        self._sticky_wilds: set[tuple[int,int]] = set()

        self._reels:      list[_Reel]    = []
        self._pods_ui:    list[_PodWidget] = []
        self._stop_timers: list[QTimer]  = []
        self._overlay: _WinOverlay | None = None

        self._tick_timer = QTimer(self)
        self._tick_timer.timeout.connect(self._tick)
        self._tick_timer.start(26)

        self._build_ui()
        self._update_credits()
        self._update_bet_ui()

    # ── UI ─────────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)

        card = QWidget()
        card.setObjectName("card")
        card.setStyleSheet("""
            QWidget#card {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #07011a, stop:1 #130535);
                border: 2px solid #6c3483;
                border-radius: 18px;
            }
        """)
        outer.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(14, 10, 14, 12)
        lay.setSpacing(6)

        # ── Titelzeile ──────────────────────────────────────────────────────
        tr = QHBoxLayout()
        tl = QLabel("🎩  Alice's Wunderrad  🐇")
        tl.setStyleSheet(
            "color:#c9a227; font:bold 16px 'Segoe UI'; background:transparent;"
        )
        tr.addWidget(tl)
        tr.addStretch()
        wl = QLabel("2 4 3  W A Y S")
        wl.setStyleSheet(
            "color:#8e44ad; font:bold 10px 'Segoe UI'; background:transparent;"
        )
        tr.addWidget(wl)
        tr.addSpacing(8)
        bx = QPushButton("✕")
        bx.setFixedSize(24, 24)
        bx.setStyleSheet("""
            QPushButton {
                background:#2d0a50; color:#c9a0e0;
                border:none; border-radius:12px; font-size:11px;
            }
            QPushButton:hover { background:#c0392b; color:white; }
        """)
        bx.clicked.connect(self.close)
        tr.addWidget(bx)
        lay.addLayout(tr)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#6c348350;")
        lay.addWidget(sep)

        # ── Credits + Einsatz ────────────────────────────────────────────────
        cr = QHBoxLayout()
        self._cred_lbl = QLabel()
        self._cred_lbl.setStyleSheet("""
            QLabel {
                color:#c9a227; font:bold 14px 'Segoe UI';
                background:#0f0230; border:1px solid #6c3483;
                border-radius:7px; padding:2px 12px;
            }
        """)
        cr.addWidget(self._cred_lbl)
        cr.addStretch()

        bet_lbl_prefix = QLabel("Einsatz:")
        bet_lbl_prefix.setStyleSheet(
            "color:#8e44ad; font:10px 'Segoe UI'; background:transparent;"
        )
        bet_minus = QPushButton("▼")
        bet_minus.setFixedSize(22, 22)
        bet_minus.setStyleSheet("""
            QPushButton {
                background:#1a073d; color:#c9a0e0;
                border:1px solid #6c3483; border-radius:4px; font-size:9px;
            }
            QPushButton:hover { background:#6c3483; }
        """)
        bet_minus.clicked.connect(self._bet_down)
        self._bet_lbl = QLabel()
        self._bet_lbl.setFixedWidth(66)
        self._bet_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._bet_lbl.setStyleSheet("""
            QLabel {
                color:#c9a227; font:bold 10px 'Segoe UI';
                background:#0f0230; border:1px solid #6c348360;
                border-radius:4px; padding:1px 4px;
            }
        """)
        bet_plus = QPushButton("▲")
        bet_plus.setFixedSize(22, 22)
        bet_plus.setStyleSheet("""
            QPushButton {
                background:#1a073d; color:#c9a0e0;
                border:1px solid #6c3483; border-radius:4px; font-size:9px;
            }
            QPushButton:hover { background:#6c3483; }
        """)
        bet_plus.clicked.connect(self._bet_up)
        cr.addWidget(bet_lbl_prefix)
        cr.addSpacing(4)
        cr.addWidget(bet_minus)
        cr.addWidget(self._bet_lbl)
        cr.addWidget(bet_plus)
        lay.addLayout(cr)

        # ── 3 Sammelpods ─────────────────────────────────────────────────────
        pod_row = QHBoxLayout()
        pod_row.setSpacing(5)
        pod_row.addStretch()
        for i in range(3):
            pw = _PodWidget(i)
            self._pods_ui.append(pw)
            pod_row.addWidget(pw)
        pod_row.addStretch()
        lay.addLayout(pod_row)

        # ── Reel-Rahmen ──────────────────────────────────────────────────────
        rf = QFrame()
        rf.setStyleSheet("""
            QFrame {
                background: rgba(10,2,30,200);
                border: 2px solid #4a235a;
                border-radius: 10px;
            }
        """)
        rl = QHBoxLayout(rf)
        rl.setContentsMargins(8, 8, 8, 8)
        rl.setSpacing(5)
        rl.addStretch()
        for _ in range(_REELS):
            r = _Reel()
            self._reels.append(r)
            rl.addWidget(r)
        rl.addStretch()
        lay.addWidget(rf)

        # ── Ergebnis-Label ───────────────────────────────────────────────────
        self._res_lbl = QLabel("")
        self._res_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._res_lbl.setFixedHeight(36)
        self._res_lbl.setWordWrap(True)
        self._res_lbl.setStyleSheet(
            "color:white; font:bold 12px 'Segoe UI'; background:transparent;"
        )
        lay.addWidget(self._res_lbl)

        # ── Modus-Banner ─────────────────────────────────────────────────────
        self._mode_lbl = QLabel("")
        self._mode_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._mode_lbl.setFixedHeight(24)
        self._mode_lbl.setStyleSheet("""
            QLabel {
                color:#c9a227; font:bold 10px 'Segoe UI';
                background:rgba(201,162,39,12);
                border:1px solid #c9a22750; border-radius:5px;
            }
        """)
        self._mode_lbl.hide()
        lay.addWidget(self._mode_lbl)

        # ── Spin-Button ──────────────────────────────────────────────────────
        self._spin_btn = QPushButton("🎰   DREHEN")
        self._spin_btn.setFixedHeight(44)
        self._spin_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #e74c3c, stop:1 #922b21);
                color:white; font:bold 13px 'Segoe UI';
                border:none; border-radius:10px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ff6b6b, stop:1 #e74c3c);
            }
            QPushButton:disabled { background:#1a0535; color:#4a2a6a; }
        """)
        self._spin_btn.clicked.connect(self._spin)
        lay.addWidget(self._spin_btn)

        # ── +Credits ─────────────────────────────────────────────────────────
        btn_add = QPushButton("🪙  +200 Credits")
        btn_add.setStyleSheet("""
            QPushButton {
                background:transparent; color:#c9a227; font:10px 'Segoe UI';
                border:1px solid #c9a22740; border-radius:5px; padding:2px 10px;
            }
            QPushButton:hover { background:rgba(201,162,39,18); }
        """)
        btn_add.clicked.connect(self._add_credits)
        lay.addWidget(btn_add, alignment=Qt.AlignmentFlag.AlignHCenter)

        # ── Info-Legende ─────────────────────────────────────────────────────
        info = QFrame()
        info.setStyleSheet("""
            QFrame {
                background:rgba(15,3,35,180);
                border:1px solid #4a235a; border-radius:7px;
            }
        """)
        iv = QVBoxLayout(info)
        iv.setContentsMargins(5, 3, 5, 3)
        iv.setSpacing(1)
        for row_sl in [SYMBOLS[:6], SYMBOLS[6:11]]:
            row = QHBoxLayout()
            row.setSpacing(0)
            for em, name, col, _, w3, w4, w5 in row_sl:
                is_w = (em == SYMBOLS[WILD_IDX][0])
                is_s = (em == SYMBOLS[ALICE_IDX][0])
                badge = " W" if is_w else (" S" if is_s else "")
                lb = QLabel(f"{em}{badge} {w3}·{w4}·{w5}")
                lb.setStyleSheet(
                    f"color:{'#ffffff' if is_w else col};"
                    f"{'font-weight:bold;' if is_w or is_s else ''}"
                    f"font-size:7px; font-family:'Segoe UI'; background:transparent;"
                )
                lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
                row.addWidget(lb)
            iv.addLayout(row)
        note = QLabel(
            "🔴🔵⭐ Bonusbälle → Pods  ·  3+ Bälle → Holding Spin (5 Respins)  ·  "
            "3×👸 → 10 Free Games  ·  Pods voll → Bonus"
        )
        note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        note.setStyleSheet(
            "color:#8e44ad70; font:7px 'Segoe UI'; background:transparent;"
        )
        iv.addWidget(note)
        lay.addWidget(info)

        foot = QLabel("🐇  Folge dem weißen Kaninchen  🎩")
        foot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        foot.setStyleSheet(
            "color:#4a235a80; font:italic 8px 'Segoe UI'; background:transparent;"
        )
        lay.addWidget(foot)

        # Overlay (immer ganz oben, keine Mausinteraktion)
        self._overlay = _WinOverlay(self)

    # ── Hilfsmethoden ──────────────────────────────────────────────────────────
    def _update_credits(self) -> None:
        self._cred_lbl.setText(f"🪙  {self._credits}  Credits")

    def _update_bet_ui(self) -> None:
        bet = BET_LEVELS[self._bet_idx]
        self._bet_lbl.setText(f"  {bet} Cr.")
        self._spin_btn.setText(f"🎰   DREHEN  –  {bet} Cr.")

    def _bet_down(self) -> None:
        if self._mode != "idle":
            return
        self._bet_idx = max(0, self._bet_idx - 1)
        self._update_bet_ui()

    def _bet_up(self) -> None:
        if self._mode != "idle":
            return
        self._bet_idx = min(len(BET_LEVELS) - 1, self._bet_idx + 1)
        self._update_bet_ui()

    def _add_credits(self) -> None:
        self._credits += 200
        self._update_credits()
        self._res_lbl.setText("🪙  +200 Credits aufgeladen!")
        self._check_can_spin()

    def _check_can_spin(self) -> None:
        can = (self._mode == "idle"
               and self._fs_left <= 0
               and not self._hold_reels
               and self._credits >= BET_LEVELS[self._bet_idx])
        self._spin_btn.setEnabled(can)

    def _show_win_anim(self, text: str, color: str = "#c9a227") -> None:
        """Zeigt das animierte Bonus-Banner."""
        if self._overlay:
            self._overlay.show_anim(text, color)

    def _do_sweat(self, reel: "_Reel") -> None:
        """Markiert Sweat-Walze + zeigt Spannungs-Animation."""
        reel.sweat = True
        self._show_win_anim("🌡  SWEAT!", "#ff6b00")

    def _grid(self) -> list[list[int]]:
        return [r.visible_symbols() for r in self._reels]

    # ── Spin ───────────────────────────────────────────────────────────────────
    def _spin(self) -> None:
        if self._mode != "idle":
            return
        bet = BET_LEVELS[self._bet_idx]
        if self._credits < bet:
            self._res_lbl.setText("❌  Nicht genug Credits!")
            return

        self._credits -= bet
        self._spin_btn.setEnabled(False)
        self._mode = "spinning"
        self._res_lbl.setText("")
        self._update_credits()

        results = [_spin_col(self._bet_idx) for _ in range(_REELS)]
        self._start_reels(results)

    def _start_reels(self, results: list[list[int]],
                     delays: list[int] | None = None) -> None:
        """Startet Walzen und stoppt sie gestaffelt."""
        for ri, r in enumerate(self._reels):
            if ri not in self._hold_reels:
                r.start()
        if delays is None:
            delays = [350 + i * 280 for i in range(_REELS)]
        for idx, (reel, syms) in enumerate(zip(self._reels, results)):
            if idx in self._hold_reels:
                continue
            d = delays[idx] if idx < len(delays) else delays[-1] + idx * 200
            t = QTimer(self)
            t.setSingleShot(True)
            t.timeout.connect(lambda r=reel, s=syms: r.schedule_stop(s))
            t.start(d)
            self._stop_timers.append(t)

    # ── Tick ───────────────────────────────────────────────────────────────────
    def _tick(self) -> None:
        for r in self._reels:
            r.tick()

        if self._mode == "spinning":
            if all(r.stopped for r in self._reels):
                self._stop_timers.clear()
                self._mode = "waiting"
                self._evaluate()

        elif self._mode == "freespinning":
            if all(r.stopped for r in self._reels):
                self._stop_timers.clear()
                self._mode = "waiting"
                self._evaluate_freespin()

        elif self._mode == "holdspinning":
            # Fertig wenn alle Zellen aller Reels ihre Animation abgeschlossen haben
            if all(r.cell_anim_done for r in self._reels):
                self._stop_timers.clear()
                self._mode = "waiting"
                self._evaluate_holdspin()

        # Watchdog: zu lange in "waiting" → Notfall-Reset (~6 s bei 26 ms Tick)
        if self._mode == "waiting":
            self._wait_ticks += 1
            if self._wait_ticks > 230:
                self._force_idle()
        else:
            self._wait_ticks = 0

    def _force_idle(self) -> None:
        """Notfall-Reset: setzt das Spiel aus festegefahrenem 'waiting'-Mode zurück."""
        self._stop_timers.clear()
        self._hold_reels.clear()
        self._hold_vals.clear()
        self._hold_jps.clear()
        self._hold_respins = 0
        self._fs_left  = 0
        self._fs_total = 0
        self._mode_lbl.hide()
        for r in self._reels:
            r.held      = False
            r.hold_mode = False
            r.clear_ball_vals()
        self._mode       = "idle"
        self._wait_ticks = 0
        self._res_lbl.setText("⚡  Bitte erneut drehen")
        self._check_can_spin()

    # ── Ball-Verarbeitung ──────────────────────────────────────────────────────
    def _process_balls(self, grid: list[list[int]], bet: int
                       ) -> tuple[dict[int, list[tuple[int,int]]], list[int]]:
        """Wertet alle Bälle im Grid aus → (ball_data, getriggerte_pods)."""
        ball_data: dict[int, list[tuple[int,int]]] = {}
        triggered: list[int] = []
        for ri, col in enumerate(grid):
            for row, sym in enumerate(col):
                if sym in IS_BALL:
                    val = _ball_val(bet)
                    ball_data.setdefault(ri, []).append((row, val))
                    pod = BALL_TO_POD[sym]
                    # Jeder Ball erhitzt den Pod um 12–30 %
                    self._pod_heat[pod] = min(
                        1.0,
                        self._pod_heat[pod] + random.uniform(0.12, 0.30)
                    )
                    self._pods_ui[pod].heat = self._pod_heat[pod]
                    # Zufalls-Trigger: Wahrscheinlichkeit steigt quadratisch mit Hitze
                    if (random.random() < self._pod_heat[pod] ** 1.6
                            and pod not in triggered):
                        self._pods_ui[pod].flash_trigger()
                        self._pod_heat[pod] = 0.0
                        QTimer.singleShot(100, lambda pi=pod: setattr(
                            self._pods_ui[pi], 'heat', 0.0)
                        )
                        triggered.append(pod)
        return ball_data, triggered

    # ── Normale Auswertung ─────────────────────────────────────────────────────
    def _evaluate(self) -> None:
        grid = self._grid()
        bet  = BET_LEVELS[self._bet_idx]
        mult = bet / 10.0

        ball_data, pod_bonuses = self._process_balls(grid, bet)

        alice_cnt = sum(1 for col in grid for s in col if s == ALICE_IDX)

        total, wins = evaluate_ways(grid, mult)
        if total > 0:
            self._credits += total
            for sym, L, ways, prize in wins:
                for ri in range(L):
                    hit = {r for r, s in enumerate(grid[ri])
                           if s == sym or s == WILD_IDX}
                    self._reels[ri].flash_win(hit)
            parts = [f"{SYMBOLS[s][0]}×{l}({w}w)=+{pr}"
                     for s, l, w, pr in wins[:3]]
            msg = "  ".join(parts)
            if len(wins) > 3:
                msg += f"  (+{len(wins)-3})"
            self._res_lbl.setText(f"🎊  {msg}")
            # Großer Gewinn-Flash ab 5× Einsatz
            if total >= BET_LEVELS[self._bet_idx] * 5:
                self._show_win_anim(f"🎊 +{total} CREDITS!", "#c9a227")
        else:
            self._res_lbl.setText("😿  Kein Treffer.")
        self._update_credits()

        # Pod-Boni mit kurzer Verzögerung
        for i, pb in enumerate(pod_bonuses):
            QTimer.singleShot(400 + i * 600, lambda p=pb: self._award_pod_bonus(p))

        # Trigger-Priorität: Holding Spin > Free Games
        ball_count = sum(len(v) for v in ball_data.values())
        # Fester Trigger: 3+ Bälle auf verschiedenen Walzen
        # Zufälliger Trigger: ab 1 Ball mit steigender Wahrscheinlichkeit
        rand_hold_prob = {1: 0.08, 2: 0.35}.get(ball_count, 0.0)
        do_hold = (ball_count >= _HOLD_TRIGGER
                   or (ball_count >= 1 and random.random() < rand_hold_prob))
        if do_hold and ball_data:
            self._show_win_anim("🔮 HOLDING SPIN!", "#7e57c2")
            QTimer.singleShot(1400, lambda: self._trigger_holdspin(ball_data, bet))
            return

        if alice_cnt >= 3:
            QTimer.singleShot(800, lambda: self._trigger_freegames(alice_cnt))
            return

        self._mode = "idle"
        self._check_can_spin()

    # ── Holding Spin  (Lightning Link Stil) ───────────────────────────────────
    def _trigger_holdspin(self, ball_data: dict[int, list[tuple[int,int]]],
                          bet: int) -> None:
        self._hold_reels   = set(ball_data.keys())
        self._hold_vals    = {ri: list(cells) for ri, cells in ball_data.items()}
        self._hold_jps     = {}
        self._hold_respins = _HOLD_RESPINS
        self._hold_bet     = bet

        for ri, cells in ball_data.items():
            r = self._reels[ri]
            r.held      = True
            r.hold_mode = True
            for row, val in cells:
                r.set_ball_val(row, val)
                r.set_ball_label(row, None)
                self._hold_jps[(ri, row)] = None

        for ri in range(_REELS):
            if ri not in self._hold_reels:
                self._reels[ri].hold_mode = True

        ball_count = sum(len(v) for v in ball_data.values())
        total_val  = sum(val for cells in ball_data.values() for _, val in cells)
        self._mode_lbl.setText(
            f"🔮  HOLD & RESPIN!  {ball_count} Bälle  ·  +{total_val}  ·  {_HOLD_RESPINS} Respins"
        )
        self._mode_lbl.show()
        self._res_lbl.setText(f"🔮  {ball_count} Bonusbälle!  Hold & Respin startet…")
        QTimer.singleShot(1600, self._run_holdspin)

    def _run_holdspin(self) -> None:
        """Hold & Spin: jede freie Zelle ist ein unabhaengiges 1x1-Reel."""
        total_locked = sum(len(c) for c in self._hold_vals.values())
        if total_locked >= _REELS * _ROWS:
            self._end_holdspin()
            return

        total_val = sum(val for cells in self._hold_vals.values() for _, val in cells)
        self._mode_lbl.setText(
            f"🔮  HOLD  ·  {total_locked}/{_REELS * _ROWS} Zellen  "
            f"·  +{total_val}  ·  {self._hold_respins} Respins"
        )

        # Pro Zelle Ergebnis-Symbol berechnen
        all_cells: list[tuple[int, int, int]] = []
        for ri in range(_REELS):
            locked_rows = {row for row, _ in self._hold_vals.get(ri, [])}
            col = _hold_col(locked_rows)
            for row in range(_ROWS):
                if row not in locked_rows:
                    all_cells.append((ri, row, col[row]))

        # Zellen als pending markieren bevor Mode gesetzt wird
        for ri, row, _s in all_cells:
            self._reels[ri].mark_cell_pending(row)

        self._mode = "holdspinning"

        # Gestaffelt starten: links->rechts, oben->unten (55 ms Abstand)
        step_ms = 55
        for idx, (ri, row, sym) in enumerate(all_cells):
            spin_ticks = 12 + ri * 2 + row   # 12-22 ticks = ~312-572 ms Drehdauer
            t = QTimer(self)
            t.setSingleShot(True)
            t.timeout.connect(
                lambda _r=self._reels[ri], _rw=row, _s=sym, _td=spin_ticks:
                    _r.start_cell_anim(_rw, _s, _td)
            )
            t.start(idx * step_ms)
            self._stop_timers.append(t)

    def _evaluate_holdspin(self) -> None:
        grid = self._grid()
        new_balls: list[tuple[int, int, int, str | None]] = []

        for ri in range(_REELS):
            # Auch auf teilweise gehaltenen Reels koennen weitere Bälle landen
            locked_rows = {row for row, _ in self._hold_vals.get(ri, [])}
            for row, sym in enumerate(grid[ri]):
                if row in locked_rows:
                    continue
                if sym in IS_BALL:
                    val, jp_label = _ball_val_hold(self._hold_bet)
                    new_balls.append((ri, row, val, jp_label))
                    pod = BALL_TO_POD[sym]
                    self._pod_heat[pod] = min(
                        1.0, self._pod_heat[pod] + random.uniform(0.12, 0.30)
                    )
                    self._pods_ui[pod].heat = self._pod_heat[pod]
                    if random.random() < self._pod_heat[pod] ** 1.6:
                        self._pods_ui[pod].flash_trigger()
                        self._pod_heat[pod] = 0.0
                        QTimer.singleShot(100, lambda pi=pod: setattr(
                            self._pods_ui[pi], 'heat', 0.0))
                        QTimer.singleShot(200, lambda p=pod: self._award_pod_bonus(p))

        if new_balls:
            for ri, row, val, lbl in new_balls:
                self._hold_reels.add(ri)
                self._hold_vals.setdefault(ri, []).append((row, val))
                self._hold_jps[(ri, row)] = lbl
                self._reels[ri].held      = True
                self._reels[ri].hold_mode = True
                self._reels[ri].set_ball_val(row, val)
                self._reels[ri].set_ball_label(row, lbl)
            self._hold_respins = _HOLD_RESPINS
            total_cells = sum(len(c) for c in self._hold_vals.values())
            self._res_lbl.setText(
                f"🔮  +{len(new_balls)} Ball/Bälle  ·  {total_cells}/{_REELS * _ROWS}"
                f"  ·  Respins reset → {_HOLD_RESPINS}"
            )
        else:
            self._hold_respins -= 1
            txt = (f"⏳  {self._hold_respins} Respin(s) verbleibend"
                   if self._hold_respins > 0 else "🔮  Hold & Respin endet…")
            self._res_lbl.setText(txt)

        total_val   = sum(val for cells in self._hold_vals.values() for _, val in cells)
        total_cells = sum(len(c) for c in self._hold_vals.values())
        full_board  = (total_cells >= _REELS * _ROWS)

        if self._hold_respins <= 0 or full_board:
            delay = 300 if full_board else 700
            QTimer.singleShot(delay, self._end_holdspin)
        else:
            self._mode_lbl.setText(
                f"🔮  HOLD  ·  {total_cells}/{_REELS * _ROWS} Zellen  "
                f"·  +{total_val}  ·  {self._hold_respins} Respins"
            )
            QTimer.singleShot(700 if new_balls else 450, self._run_holdspin)

    def _end_holdspin(self) -> None:
        total_val   = sum(val for cells in self._hold_vals.values() for _, val in cells)
        ball_cnt    = sum(len(c) for c in self._hold_vals.values())
        full_board  = (ball_cnt >= _REELS * _ROWS)

        # Grand-Jackpot-Bonus bei vollem Brett (alle 15 Zellen)
        grand_bonus = 0
        if full_board:
            grand_bonus  = BET_LEVELS[self._bet_idx] * 500
            total_val   += grand_bonus

        self._credits += total_val

        for r in self._reels:
            r.held      = False
            r.hold_mode = False
            r.clear_ball_vals()
            if total_val > 0:
                r.flash_win()

        self._hold_reels.clear()
        self._hold_vals.clear()
        self._hold_respins = 0
        self._mode_lbl.hide()

        # JP-Tier höchste Stufe ermitteln
        top_jp = None
        for lbl in self._hold_jps.values():
            if lbl == "GRAND":
                top_jp = "GRAND"; break
            elif lbl == "MAJOR" and top_jp not in ("GRAND",):
                top_jp = "MAJOR"
            elif lbl == "MINOR" and top_jp not in ("GRAND", "MAJOR"):
                top_jp = "MINOR"
            elif lbl == "MINI" and top_jp is None:
                top_jp = "MINI"
        self._hold_jps.clear()

        if full_board:
            msg = f"🏆  GRAND JACKPOT!  {ball_cnt}/15 Zellen  =  +{total_val} Credits!"
            self._show_win_anim(f"🏆 GRAND JACKPOT! +{total_val}!", "#ff8a65")
        elif top_jp == "MAJOR":
            msg = f"💎  MAJOR!  {ball_cnt} Bälle  =  +{total_val} Credits!"
            self._show_win_anim(f"💎 MAJOR! +{total_val}!", "#fff59d")
        elif ball_cnt >= _REELS:
            msg = f"🎊  MEGA HOLD!  {ball_cnt} Bälle  =  +{total_val} Credits!  🔥"
            self._show_win_anim(f"🔥 MEGA HOLD! +{total_val}!", "#ff6b35")
        elif total_val >= BET_LEVELS[self._bet_idx] * 3:
            msg = f"✨  SUPER HOLD!  {ball_cnt} Bälle  =  +{total_val} Credits!"
            self._show_win_anim(f"🔮 HOLD +{total_val} Credits!", "#c9a227")
        else:
            msg = f"🔮  Hold & Respin: {ball_cnt} Bälle  =  +{total_val} Credits"

        self._res_lbl.setText(msg)
        self._update_credits()
        self._mode = "idle"
        self._check_can_spin()

    # ── Free Games ─────────────────────────────────────────────────────────────
    def _trigger_freegames(self, alice_cnt: int) -> None:
        extra = max(0, alice_cnt - 2) * _FS_RETRIG
        self._fs_total = _FS_BASE + extra
        self._fs_left  = self._fs_total
        self._fs_wins  = 0
        self._sticky_wilds.clear()
        self._mode_lbl.setText(
            f"👸  FREE GAMES!  {self._fs_total} Spins  ·  {_FS_MULT}×  ·  Sticky Wilds"
        )
        self._mode_lbl.show()
        self._res_lbl.setText(
            f"👸  {alice_cnt}× Alice!  {self._fs_total} Free Games starten…"
        )
        self._show_win_anim(
            f"👸  {self._fs_total} FREE GAMES!",
            "#e91e63"
        )
        QTimer.singleShot(1800, self._run_freespin)

    def _run_freespin(self) -> None:
        if self._fs_left <= 0:
            self._end_freegames()
            return

        self._fs_left -= 1

        # Sweat: ~22% Chance – eine zufällige Walze spinnt als letzte
        # und enthüllt eine volle Wild-Walze
        sweat_ri = -1
        if random.random() < 0.22:
            sweat_ri = random.randrange(_REELS)

        # Ergebnisse erzeugen: Sticky Wilds + erhöhte Wild-Chance
        results = []
        for ri in range(_REELS):
            if ri == sweat_ri:
                col = [WILD_IDX, WILD_IDX, WILD_IDX]   # Sweat → volle Wild-Walze
            else:
                col = _spin_col(self._bet_idx, extra_ball=0.02)
                for row in range(_ROWS):
                    if (ri, row) in self._sticky_wilds:
                        col[row] = WILD_IDX
                    elif col[row] not in (ALICE_IDX,) and random.random() < 0.05:
                        col[row] = WILD_IDX
            results.append(col)

        for r in self._reels:
            r.sweat = False
            r.start()

        self._mode = "freespinning"

        # Normale Walzen stoppen (Sweat-Walze ausgelassen → kommt dramatisch als letzte)
        for idx, (reel, syms) in enumerate(zip(self._reels, results)):
            if idx == sweat_ri:
                continue
            t = QTimer(self)
            t.setSingleShot(True)
            t.timeout.connect(lambda r=reel, s=syms: r.schedule_stop(s))
            t.start(260 + idx * 200)
            self._stop_timers.append(t)

        if sweat_ri >= 0:
            # Sweat-Walze stoppt 1000ms nach der letzten normalen
            last_normal_ms = 260 + (_REELS - 1) * 200
            sweat_stop_ms  = last_normal_ms + 1000
            sr = self._reels[sweat_ri]
            ss = results[sweat_ri]
            # 600ms vor Stopp: Feuer-Animation + Overlay
            QTimer.singleShot(sweat_stop_ms - 600,
                             lambda r=sr: self._do_sweat(r))
            t_sw = QTimer(self)
            t_sw.setSingleShot(True)
            t_sw.timeout.connect(lambda r=sr, s=ss: r.schedule_stop(s))
            t_sw.start(sweat_stop_ms)
            self._stop_timers.append(t_sw)
            # Sweat-Flag nach dem Stopp löschen
            QTimer.singleShot(sweat_stop_ms + 1200,
                             lambda r=sr: setattr(r, 'sweat', False))

        self._mode_lbl.setText(
            f"👸  FREE GAMES  ·  {self._fs_left + 1}/{self._fs_total}  "
            f"·  {_FS_MULT}×  ·  Stickies: {len(self._sticky_wilds)}"
        )

    def _evaluate_freespin(self) -> None:
        grid = self._grid()
        bet  = BET_LEVELS[self._bet_idx]
        mult = bet / 10.0 * _FS_MULT

        ball_data, pod_bonuses = self._process_balls(grid, bet)
        for i, pb in enumerate(pod_bonuses):
            QTimer.singleShot(200 + i * 400, lambda p=pb: self._award_pod_bonus(p))

        # Retrigger ab 2× Alice
        alice_cnt = sum(1 for col in grid for s in col if s == ALICE_IDX)
        if alice_cnt >= 2:
            bonus = _FS_RETRIG * alice_cnt
            self._fs_left += bonus
            self._show_win_anim(f"👸 +{bonus} SPINS!", "#e91e63")
            self._res_lbl.setText(f"👸  Retrigger! +{bonus} Free Spins!")

        # Sticky Wilds
        for ri, col in enumerate(grid):
            for row, sym in enumerate(col):
                if sym == WILD_IDX:
                    self._sticky_wilds.add((ri, row))

        total, wins = evaluate_ways(grid, mult)
        if total > 0:
            self._credits += total
            self._fs_wins  += total
            for sym, L, ways, prize in wins:
                for ri in range(L):
                    hit = {r for r, s in enumerate(grid[ri])
                           if s == sym or s == WILD_IDX}
                    self._reels[ri].flash_win(hit)
            parts = [f"{SYMBOLS[s][0]}×{l}=+{pr}" for s, l, w, pr in wins[:3]]
            if alice_cnt < 2:
                self._res_lbl.setText(f"✨{_FS_MULT}×  " + "  ".join(parts))
        else:
            if alice_cnt < 2:
                self._res_lbl.setText(f"⭐  {self._fs_left} Free Game(s) verbleibend")

        self._update_credits()

        if self._fs_left <= 0:
            QTimer.singleShot(950, self._end_freegames)
        else:
            QTimer.singleShot(650, self._run_freespin)

    def _end_freegames(self) -> None:
        for r in self._reels:
            r.flash_win()
        self._sticky_wilds.clear()
        self._mode_lbl.hide()
        spins = self._fs_total
        wins  = self._fs_wins
        self._fs_total = 0
        self._fs_left  = 0
        self._fs_wins  = 0
        self._res_lbl.setText(
            f"👸  Free Games beendet!  {spins} Spins  ·  Gewinn: +{wins} Credits!"
        )
        self._show_win_anim(f"👸 FREE GAMES  +{wins} Credits!", "#e91e63")
        self._mode = "idle"
        self._check_can_spin()

    # ── Pod-Boni ───────────────────────────────────────────────────────────────
    def _award_pod_bonus(self, pod_idx: int) -> None:
        """Vergibt den Bonus eines gefüllten Pods."""
        bet = BET_LEVELS[self._bet_idx]

        if pod_idx == 0:
            # 👑 Herzkönigin: 3 Karten-Picks → bester Gewinn
            picks = sorted(
                [random.randint(bet * 3, bet * 30) for _ in range(3)],
                reverse=True
            )
            win = picks[0]
            self._credits += win
            self._res_lbl.setText(
                f"👑  Herzkönigin-Bonus!  Karten: {picks}  →  +{win} Credits!"
            )
            self._show_win_anim(f"👑 +{win} Herzkönigin!", "#c62828")

        elif pod_idx == 1:
            # 🔵 Wunderland: Free Games
            bonus_fs = _FS_BASE if self._fs_left == 0 else _FS_RETRIG
            if self._fs_left > 0:
                self._fs_left += bonus_fs
                self._mode_lbl.setText(
                    f"👸  FREE GAMES  ·  +{bonus_fs} Spins  "
                    f"·  {self._fs_left} verbleibend  ·  {_FS_MULT}×"
                )
            else:
                self._fs_total = bonus_fs
                self._fs_left  = bonus_fs
                self._sticky_wilds.clear()
                self._mode_lbl.setText(
                    f"👸  FREE GAMES!  {bonus_fs} Spins  ·  {_FS_MULT}× Multiplikator"
                )
                self._mode_lbl.show()
                QTimer.singleShot(1200, self._run_freespin)
            self._res_lbl.setText(f"🔵  Wunderland-Bonus!  +{bonus_fs} Free Games!")
            self._show_win_anim(f"🔵 +{bonus_fs} FREE GAMES!", "#1565c0")

        elif pod_idx == 2:
            # ⭐ Goldschatz: Jackpot-Rad
            tiers   = [bet * 10, bet * 20, bet * 50, bet * 100, bet * 200]
            weights = [40, 28, 18, 10, 4]
            win = random.choices(tiers, weights=weights, k=1)[0]
            self._credits += win
            self._res_lbl.setText(f"⭐  Goldschatz-Jackpot!  +{win} Credits!  🎡")
            self._show_win_anim(f"⭐ JACKPOT! +{win} Credits!", "#f9a825")

        self._update_credits()

    # ── Draggable ──────────────────────────────────────────────────────────────
    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e) -> None:
        if e.buttons() & Qt.MouseButton.LeftButton and self._drag_pos is not None:
            self.move(self.pos() + e.globalPosition().toPoint() - self._drag_pos)
            self._drag_pos = e.globalPosition().toPoint()

    def mouseReleaseEvent(self, e) -> None:
        self._drag_pos = None

    def closeEvent(self, e) -> None:
        self._tick_timer.stop()
        for t in self._stop_timers:
            t.stop()
        super().closeEvent(e)
