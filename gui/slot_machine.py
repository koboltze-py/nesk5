"""
gui/slot_machine.py  –  🎩 Alice im Wunderland — Glücksrad
5 Reels · 3 Reihen · 243 Ways to Win (wie Buffalo) · Wild · Alice-Bonus
"""
from __future__ import annotations

import math
import random
from itertools import product

from PySide6.QtCore import Qt, QRectF, QPointF, QTimer
from PySide6.QtGui import (
    QColor, QBrush, QPen, QPainter, QFont,
    QLinearGradient, QRadialGradient,
)
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QWidget, QFrame,
)

# ══════════════════════════════════════════════════════════════════════════════
#  Symbol-Tabelle
#  Index  Emoji   Name              Farbe     Gewicht  win×3 win×4 win×5
# ══════════════════════════════════════════════════════════════════════════════
WILD_IDX  = 0   # Wild ersetzt alle außer Alice-Scatter und Wild selbst
ALICE_IDX = 1   # Scatter → Bonus

SYMBOLS: list[tuple] = [
    # idx  emoji   name               farbe      wt  w3   w4    w5
    ("🪄", "Wild",              "#ffffff",   4,  0,    0,  5000),  # 0 Wild
    ("👸", "Alice (Scatter)",   "#fce4ec",   6, 400, 1000, 3000),  # 1 Scatter
    ("🎩", "Hutmacher",         "#f9ca24",   9, 120,  320,  800),  # 2
    ("🐱", "Grinsekatze",       "#c39bd3",  11,  80,  220,  550),  # 3
    ("🐰", "Weißes Kaninchen",  "#dfe6e9",  13,  55,  160,  380),  # 4
    ("👑", "Herzkönigin",       "#e74c3c",  15,  40,  110,  250),  # 5
    ("⏰", "Taschenuhr",        "#fab1a0",  18,  28,   75,  160),  # 6
    ("🍄", "Wunderpilz",        "#55efc4",  21,  18,   50,  100),  # 7
    ("🫖", "Teekanne",          "#74b9ff",  24,  10,   30,   70),  # 8
    ("🔑", "Schlüssel",         "#fdcb6e",  28,   7,   20,   45),  # 9
    ("🌹", "Rote Rose",         "#fd79a8",  32,   4,   12,   30),  # 10
]

_POOL = sum([[i] * SYMBOLS[i][3] for i in range(len(SYMBOLS))], [])
_COST        = 10
_BONUS_SPINS = 5
_BONUS_WINS  = {3: 300, 4: 1000, 5: 4000}
_ROWS        = 3    # sichtbare Zeilen pro Reel
_REELS       = 5


def _rnd() -> int:
    return random.choice(_POOL)


# ══════════════════════════════════════════════════════════════════════════════
#  243-Ways Auswertung
# ══════════════════════════════════════════════════════════════════════════════
def evaluate_ways(grid: list[list[int]]) -> tuple[int, list[tuple]]:
    """
    grid[reel][row]  —  5 Reels × 3 Reihen.
    Gibt (Gesamtgewinn, [(sym, länge, multiplikator), …]) zurück.
    Wild (WILD_IDX) ersetzt jedes Symbol außer ALICE_IDX.
    """
    wins: list[tuple] = []
    total = 0

    # Alle möglichen Symbolkandidaten (ohne Scatter, ohne Wild als Basiswert)
    base_syms = {s for col in grid for s in col
                 if s not in (WILD_IDX, ALICE_IDX)}

    checked: set[int] = set()
    for sym in base_syms:
        if sym in checked:
            continue
        checked.add(sym)

        # Wie viele Wege gibt es für dieses Symbol links→rechts?
        ways = 1
        length = 0
        for reel_idx in range(_REELS):
            col = grid[reel_idx]
            matches = sum(1 for s in col if s == sym or s == WILD_IDX)
            if matches == 0:
                break
            ways   *= matches
            length += 1

        if length >= 3:
            sym_entry = SYMBOLS[sym]
            if length == 5:
                base_win = sym_entry[6]
            elif length == 4:
                base_win = sym_entry[5]
            else:
                base_win = sym_entry[4]
            prize = base_win * ways
            total += prize
            wins.append((sym, length, ways, prize))

    # Wilds alleine (nur 5× Wild → Jackpot)
    wild_reels = sum(1 for col in grid if all(s == WILD_IDX for s in col))
    if wild_reels == 5:
        total += SYMBOLS[WILD_IDX][6]
        wins.append((WILD_IDX, 5, 1, SYMBOLS[WILD_IDX][6]))

    return total, wins


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

    # ── Tick ───────────────────────────────────────────────────────────────────
    def tick(self) -> bool:
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

            draw_symbol(p, sym_i, W / 2.0, cy, SH * 0.38, fade)

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

        # Gewinnlinien-Ticks
        p.setPen(QPen(QColor(201, 162, 39, 60), 1, Qt.PenStyle.DotLine))
        for row in range(1, _ROWS):
            p.drawLine(0, row * SH, W, row * SH)

        # Rahmen
        p.setPen(QPen(QColor("#c9a227") if self.held else QColor("#6c3483"),
                      3 if self.held else 2))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(QRectF(1, 1, W - 2, H - 2), 7, 7)
        p.end()


# ══════════════════════════════════════════════════════════════════════════════
#  SlotMachineDialog
# ══════════════════════════════════════════════════════════════════════════════
class SlotMachineDialog(QDialog):
    """🎩 Alice im Wunderland — 5×3 · 243 Ways · Wild · Alice-Bonus"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("🎩 Alice im Wunderland — Glücksrad")
        self.setWindowFlags(
            Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(580, 700)

        self._credits          = 100
        self._mode             = "idle"
        self._bonus_spins_left = 0
        self._held_reels: set[int]    = set()
        self._reels:      list[_Reel] = []
        self._stop_timers: list[QTimer] = []
        self._drag_pos = None

        self._tick_timer = QTimer(self)
        self._tick_timer.timeout.connect(self._tick)
        self._tick_timer.start(26)

        self._build_ui()
        self._update_credits()

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
        lay.setContentsMargins(16, 12, 16, 14)
        lay.setSpacing(7)

        # ── Titelzeile ──────────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title_lbl = QLabel("🎩  Alice's Wunderrad  🐇")
        title_lbl.setStyleSheet(
            "color:#c9a227; font:bold 17px 'Segoe UI'; background:transparent;"
        )
        title_row.addWidget(title_lbl)
        ways_lbl = QLabel("2 4 3  W A Y S")
        ways_lbl.setStyleSheet(
            "color:#8e44ad; font:bold 11px 'Segoe UI'; background:transparent;"
        )
        title_row.addStretch()
        title_row.addWidget(ways_lbl)
        title_row.addSpacing(10)
        btn_x = QPushButton("✕")
        btn_x.setFixedSize(26, 26)
        btn_x.setStyleSheet("""
            QPushButton {
                background:#2d0a50; color:#c9a0e0;
                border:none; border-radius:13px; font-size:12px;
            }
            QPushButton:hover { background:#c0392b; color:white; }
        """)
        btn_x.clicked.connect(self.close)
        title_row.addWidget(btn_x)
        lay.addLayout(title_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#6c348350;")
        lay.addWidget(sep)

        # ── Credits ─────────────────────────────────────────────────────────
        cr = QHBoxLayout()
        cr.addStretch()
        self._cred_lbl = QLabel()
        self._cred_lbl.setStyleSheet("""
            QLabel {
                color:#c9a227; font:bold 15px 'Segoe UI';
                background:#0f0230; border:1px solid #6c3483;
                border-radius:8px; padding:2px 14px;
            }
        """)
        cr.addWidget(self._cred_lbl)
        cr.addStretch()
        lay.addLayout(cr)

        # ── Reel-Rahmen ─────────────────────────────────────────────────────
        reel_frame = QFrame()
        reel_frame.setStyleSheet("""
            QFrame {
                background: rgba(10,2,30,200);
                border: 2px solid #4a235a;
                border-radius: 10px;
            }
        """)
        reel_lay = QHBoxLayout(reel_frame)
        reel_lay.setContentsMargins(8, 8, 8, 8)
        reel_lay.setSpacing(5)
        reel_lay.addStretch()
        for _ in range(_REELS):
            r = _Reel()
            self._reels.append(r)
            reel_lay.addWidget(r)
        reel_lay.addStretch()
        lay.addWidget(reel_frame)

        # ── Ergebnis-Label ───────────────────────────────────────────────────
        self._res_lbl = QLabel("")
        self._res_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._res_lbl.setFixedHeight(38)
        self._res_lbl.setWordWrap(True)
        self._res_lbl.setStyleSheet(
            "color:white; font:bold 13px 'Segoe UI'; background:transparent;"
        )
        lay.addWidget(self._res_lbl)

        # ── Bonus-Banner ─────────────────────────────────────────────────────
        self._bonus_lbl = QLabel("")
        self._bonus_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._bonus_lbl.setFixedHeight(26)
        self._bonus_lbl.setStyleSheet("""
            QLabel {
                color:#c9a227; font:bold 11px 'Segoe UI';
                background:rgba(201,162,39,15);
                border:1px solid #c9a22750;
                border-radius:5px;
            }
        """)
        self._bonus_lbl.hide()
        lay.addWidget(self._bonus_lbl)

        # ── Spin-Button ──────────────────────────────────────────────────────
        self._spin_btn = QPushButton(f"🎰   DREHEN   –  {_COST} Credits")
        self._spin_btn.setFixedHeight(46)
        self._spin_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #e74c3c, stop:1 #922b21);
                color:white; font:bold 14px 'Segoe UI';
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
        btn_add = QPushButton("🪙  +100 Credits")
        btn_add.setStyleSheet("""
            QPushButton {
                background:transparent; color:#c9a227; font:11px 'Segoe UI';
                border:1px solid #c9a22740; border-radius:5px; padding:2px 12px;
            }
            QPushButton:hover { background:rgba(201,162,39,18); }
        """)
        btn_add.clicked.connect(self._add_credits)
        lay.addWidget(btn_add, alignment=Qt.AlignmentFlag.AlignHCenter)

        # ── Legende ──────────────────────────────────────────────────────────
        leg = QFrame()
        leg.setStyleSheet("""
            QFrame {
                background:rgba(15,3,35,180);
                border:1px solid #4a235a; border-radius:7px;
            }
        """)
        lv = QVBoxLayout(leg)
        lv.setContentsMargins(6, 4, 6, 4)
        lv.setSpacing(2)

        # Zeile 1 — erste 6 Symbole
        for row_slice in [SYMBOLS[:6], SYMBOLS[6:]]:
            row = QHBoxLayout()
            row.setSpacing(0)
            for em, name, col, _, w3, w4, w5 in row_slice:
                is_wild    = (em == SYMBOLS[WILD_IDX][0])
                is_scatter = (em == SYMBOLS[ALICE_IDX][0])
                badge = " W" if is_wild else (" S" if is_scatter else "")
                lb = QLabel(f"{em}{badge}  {w3}·{w4}·{w5}")
                lb.setStyleSheet(
                    f"color:{'#ffffff' if is_wild else col};"
                    f"{'font-weight:bold;' if is_wild or is_scatter else ''}"
                    f"font-size:8px; font-family:'Segoe UI'; background:transparent;"
                )
                lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
                row.addWidget(lb)
            lv.addLayout(row)

        note = QLabel(
            "🪄 Wild ersetzt alle  ·  👸 3+ Scatter = BONUS  ·  243 Ways = alle Kombi-Treffer l→r"
        )
        note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        note.setStyleSheet(
            "color:#8e44ad70; font:7px 'Segoe UI'; background:transparent;"
        )
        lv.addWidget(note)
        lay.addWidget(leg)

        # ── Fußzeile ─────────────────────────────────────────────────────────
        foot = QLabel("🐇  Folge dem weißen Kaninchen  🎩")
        foot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        foot.setStyleSheet(
            "color:#4a235a90; font:italic 9px 'Segoe UI'; background:transparent;"
        )
        lay.addWidget(foot)

    # ── Hilfsmethoden ──────────────────────────────────────────────────────────
    def _update_credits(self) -> None:
        self._cred_lbl.setText(f"🪙  {self._credits}  Credits")

    def _add_credits(self) -> None:
        self._credits += 100
        self._update_credits()
        self._res_lbl.setText("🪙  +100 Credits aufgeladen!")
        if self._mode == "idle":
            self._spin_btn.setEnabled(True)

    def _grid(self) -> list[list[int]]:
        """Gibt aktuelle 5×3 Grid zurück: grid[reel][row]."""
        return [r.visible_symbols() for r in self._reels]

    # ── Spin ───────────────────────────────────────────────────────────────────
    def _spin(self) -> None:
        if self._mode != "idle":
            return
        if self._credits < _COST:
            self._res_lbl.setText("❌  Nicht genug Credits!")
            return

        self._credits -= _COST
        self._spin_btn.setEnabled(False)
        self._mode = "spinning"
        self._res_lbl.setText("")
        self._update_credits()

        # Zufalls-Ergebnisse: 3 Symbole je Reel
        results = [[_rnd() for _ in range(_ROWS)] for _ in range(_REELS)]

        for r in self._reels:
            r.start()

        # Gestaffeltes Stoppen: 350 · 650 · 950 · 1250 · 1550 ms (schneller!)
        for idx, (reel, syms) in enumerate(zip(self._reels, results)):
            t = QTimer(self)
            t.setSingleShot(True)
            t.timeout.connect(lambda r=reel, s=syms: r.schedule_stop(s))
            t.start(350 + idx * 300)
            self._stop_timers.append(t)

    # ── Tick ───────────────────────────────────────────────────────────────────
    def _tick(self) -> None:
        for r in self._reels:
            r.tick()

        if self._mode == "spinning":
            if all(r.stopped for r in self._reels):
                self._stop_timers.clear()
                self._mode = "idle"
                self._evaluate()

        elif self._mode == "bonus_spinning":
            non_held = [i for i in range(_REELS) if i not in self._held_reels]
            if all(self._reels[i].stopped for i in non_held):
                self._stop_timers.clear()
                self._mode = "idle"
                self._evaluate_bonus_spin()

    # ── Auswertung ─────────────────────────────────────────────────────────────
    def _evaluate(self) -> None:
        grid       = self._grid()
        flat       = [s for col in grid for s in col]
        alice_cnt  = sum(1 for s in flat if s == ALICE_IDX)

        if alice_cnt >= 3:
            self._start_bonus(grid)
            return

        total, wins = evaluate_ways(grid)

        if total > 0:
            self._credits += total
            # Gewinn-Flash: markiere Zeilen die zu Gewinnen beitragen
            for sym, length, ways, prize in wins:
                for reel_idx in range(length):
                    cols = grid[reel_idx]
                    hit_rows = {r for r, s in enumerate(cols)
                                if s == sym or s == WILD_IDX}
                    self._reels[reel_idx].flash_win(hit_rows)

            parts = [
                f"{SYMBOLS[sym][0]}×{length}({ways}w)=+{prize}"
                for sym, length, ways, prize in wins[:3]
            ]
            msg = "  ".join(parts)
            if len(wins) > 3:
                msg += f"  (+{len(wins)-3} weitere)"
            self._res_lbl.setText(f"🎊  {msg}")
        else:
            self._res_lbl.setText("😿  Kein Treffer.")

        self._update_credits()
        self._spin_btn.setEnabled(self._credits >= _COST)

    # ── Bonus ──────────────────────────────────────────────────────────────────
    def _start_bonus(self, grid: list[list[int]]) -> None:
        self._held_reels       = set()
        self._bonus_spins_left = _BONUS_SPINS

        for reel_idx, col in enumerate(grid):
            if ALICE_IDX in col:
                self._held_reels.add(reel_idx)
                row_set = {r for r, s in enumerate(col) if s == ALICE_IDX}
                self._reels[reel_idx].held = True
                self._reels[reel_idx].flash_win(row_set)

        n = len(self._held_reels)
        self._bonus_lbl.setText(
            f"👸  ALICE-BONUS!  {n} Reel(s) mit Alice  —  {_BONUS_SPINS} Holding Spins"
        )
        self._bonus_lbl.show()
        self._res_lbl.setText(f"👸  {n}× Alice gefunden!  Bonus startet…")
        QTimer.singleShot(1200, self._run_bonus_spin)

    def _run_bonus_spin(self) -> None:
        if self._bonus_spins_left <= 0 or len(self._held_reels) >= _REELS:
            self._end_bonus()
            return

        self._bonus_spins_left -= 1
        non_held = [i for i in range(_REELS) if i not in self._held_reels]

        # Erhöhte Alice-Wahrscheinlichkeit (28%)
        results = {
            i: [ALICE_IDX if random.random() < 0.28 else _rnd()
                for _ in range(_ROWS)]
            for i in non_held
        }

        for i in non_held:
            self._reels[i].start()

        self._mode = "bonus_spinning"

        for j, i in enumerate(non_held):
            t = QTimer(self)
            t.setSingleShot(True)
            t.timeout.connect(
                lambda r=self._reels[i], s=results[i]: r.schedule_stop(s)
            )
            t.start(350 + j * 280)
            self._stop_timers.append(t)

        self._bonus_lbl.setText(
            f"👸  HOLDING SPINS  —  {len(self._held_reels)} Hold  "
            f"—  noch {self._bonus_spins_left} Spin(s)"
        )
        self._res_lbl.setText("🎠  …")

    def _evaluate_bonus_spin(self) -> None:
        grid = self._grid()
        new_holds: list[int] = []

        for i in range(_REELS):
            if i not in self._held_reels:
                col = grid[i]
                if ALICE_IDX in col:
                    new_holds.append(i)
                    self._held_reels.add(i)
                    rows = {r for r, s in enumerate(col) if s == ALICE_IDX}
                    self._reels[i].held = True
                    self._reels[i].flash_win(rows)

        if new_holds:
            self._res_lbl.setText(
                f"👸  {len(new_holds)} neue Alice!  ({len(self._held_reels)}/{_REELS} Reels)"
            )
        else:
            self._res_lbl.setText(
                f"⏳  {self._bonus_spins_left} Spin(s) verbleibend…"
                if self._bonus_spins_left > 0 else "🎩  Bonus endet…"
            )

        delay = 900 if new_holds else 500
        if self._bonus_spins_left <= 0 or len(self._held_reels) >= _REELS:
            QTimer.singleShot(delay, self._end_bonus)
        else:
            QTimer.singleShot(delay, self._run_bonus_spin)

    def _end_bonus(self) -> None:
        alice_n = len(self._held_reels)
        win     = _BONUS_WINS.get(alice_n, alice_n * 80)
        self._credits += win

        for r in self._reels:
            r.held = False
            if win > 0:
                r.flash_win()
        self._held_reels.clear()
        self._bonus_lbl.hide()

        if alice_n == 5:
            msg = f"👑  MEGA-BONUS!  5 Reels voll Alice  =  +{win} Credits!  👑"
        elif alice_n == 4:
            msg = f"🎊  SUPER-BONUS!  4×👸  =  +{win} Credits!"
        else:
            msg = f"✨  Bonus: {alice_n}×👸  =  +{win} Credits"

        self._res_lbl.setText(msg)
        self._update_credits()
        self._spin_btn.setEnabled(self._credits >= _COST)

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

