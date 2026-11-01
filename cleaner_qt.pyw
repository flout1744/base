"""Base — очистка и аналитика телефонных баз (версия на PySide6/Qt).

Перенос с CustomTkinter на Qt: окно масштабируется нативно и плавно,
без рывков при изменении размера. Вся бизнес-логика и вшитые базы
(города, телефонные диапазоны, специализации, триггер-слова) берутся
из модуля cleaner_core без изменений."""

import os
import re
import sys
import threading
import datetime
import random
import webbrowser

import pandas as pd

from PySide6.QtCore import (Qt, QObject, Signal, QTimer, QPropertyAnimation,
                           QEasingCurve, QSize, QPoint, QPointF, QRect, QRectF, QEvent,
                           QSequentialAnimationGroup, QParallelAnimationGroup, QVariantAnimation)
from PySide6.QtGui import (QFont, QColor, QKeySequence, QShortcut, QAction,
                           QPainter, QPen, QBrush, QPainterPath, QIcon, QFontDatabase,
                           QLinearGradient, QRadialGradient, QPixmap, QPalette, QCursor,
                           QFontMetrics)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit, QPlainTextEdit, QCheckBox,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressBar, QButtonGroup, QStackedWidget, QFrame, QDialog, QAbstractItemView,
    QMenu, QSizePolicy, QScrollArea, QGraphicsOpacityEffect, QLayout,
    QToolTip, QToolButton, QComboBox, QStyle, QStyleOptionButton, QAbstractButton,
    QSizeGrip, QListWidget, QListWidgetItem, QRadioButton,
)

try:
    import qtawesome as qta
    _QTA = True
except ImportError:
    _QTA = False


def _icon(name, color=None, fallback=""):
    """Векторная иконка для кнопок.

    В ICONS у каждой роли лежит не одно имя, а список кандидатов через «|»
    (сначала скруглённый набор mdi6, затем mdi, затем fa5s). Берётся первый,
    который есть в установленной версии qtawesome — так новый набор иконок
    не превращается в безликие кружки-заглушки на старых сборках библиотеки.
    """
    if _QTA and name:
        for cand in str(name).split("|"):
            cand = cand.strip()
            if not cand:
                continue
            try:
                return qta.icon(cand, color=color or FG)
            except Exception:
                continue
        try:
            return qta.icon("fa5s.circle", color=color or FG)
        except Exception:
            pass
    return fallback


import cleaner_core as C
try:
    import updater as U
except Exception:  # обновлялка не должна мешать запуску программы
    U = None

# ── палитра и токены оформления из ядра (тема одна — Vortex) ──
BG, BG2, BG3 = C.BG, C.BG2, C.BG3
NAV_BG = C.NAV_BG
SURFACE, BORDER, DIVIDER = C.SURFACE, C.BORDER, C.DIVIDER
ACCENT, ACCENT_HV = C.ACCENT, C.ACCENT_HV
ACCENT_DEEP, ACCENT_SOFT = C.ACCENT_DEEP, C.ACCENT_SOFT
BLUE = C.BLUE
DUPE_RED, DUPE_RED_HV, GREEN, GREEN_HV = C.DUPE_RED, C.DUPE_RED_HV, C.GREEN, C.GREEN_HV
DANGER_BG, DANGER_FG = C.DANGER_BG, C.DANGER_FG
GRAY_BTN, GRAY_HV = C.GRAY_BTN, C.GRAY_HV
FG, FG2, FG3, LOG_FG = C.FG, C.FG2, C.FG3, C.LOG_FG
ROW_ODD, HOVER_ROW = C.ROW_ODD, C.HOVER_ROW
ICONS = C.get_icon_set()   # роль -> имя(имена) иконки qtawesome
STYLE = C.get_style_set()  # шрифт, скругления, рамки, размеры хрома окна


def _hex_rgb(hex_color):
    h = str(hex_color).lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _rgba(hex_color, alpha):
    """'#3d74e4', 0.4 -> 'rgba(61,116,228,0.4)'."""
    r, g, b = _hex_rgb(hex_color)
    return f"rgba({r},{g},{b},{round(alpha, 3)})"


def _mix(c1, c2, t):
    """Смешивает два цвета (t=0 -> c1, t=1 -> c2). Нужно, чтобы получать
    промежуточные оттенки подложек, не заводя под каждый свою константу."""
    r1, g1, b1 = _hex_rgb(c1)
    r2, g2, b2 = _hex_rgb(c2)
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return f"#{r:02x}{g:02x}{b:02x}"


def _glow(hex_color, strong=0.55, mid=0.22):
    """Радиальное свечение из центра кнопки (целиком через QSS).

    Специально не через QGraphicsDropShadowEffect: этот эффект в паре с уже
    имеющимися QGraphicsOpacityEffect (fade вкладок) даёт серьёзные
    визуальные баги — см. предупреждение в App.__init__ про content_card."""
    return (f"qradialgradient(cx:0.5, cy:0.45, radius:1.15, fx:0.5, fy:0.45, "
            f"stop:0 {_rgba(hex_color, strong)}, stop:0.55 {_rgba(hex_color, mid)}, "
            f"stop:1 {_rgba(hex_color, 0.0)})")


def _glow_ring(hex_color, ring=0.55, center=0.08, edge=0.04):
    """Свечение-«кольцо» для маленьких квадратных кнопок-иконок: центр
    притушен (под ним сидит сама иконка), яркость вынесена на кольцо."""
    return (f"qradialgradient(cx:0.5, cy:0.5, radius:0.95, fx:0.5, fy:0.5, "
            f"stop:0 {_rgba(hex_color, center)}, stop:0.62 {_rgba(hex_color, ring)}, "
            f"stop:1 {_rgba(hex_color, edge)})")


def _vgrad(top, bottom):
    """Вертикальный градиент сверху вниз."""
    return f"qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {top}, stop:1 {bottom})"


def _luma(hex_color):
    """Яркость цвета 0..1 — по ней определяем, светлая тема или тёмная."""
    r, g, b = _hex_rgb(hex_color)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


# Одна и та же вёрстка обслуживает обе темы, поэтому все «вуали» (тонкие
# рамки, разделители, подложки под наведением) задаются не белым цветом, а
# через _veil(): на тёмной теме вуаль светлая, на светлой — тёмная. Иначе
# белые волосяные рамки на белом фоне просто исчезли бы.
IS_LIGHT = _luma(BG) > 0.5


def _veil(alpha):
    """Полупрозрачная вуаль поверх подложки, подходящая текущей теме."""
    if IS_LIGHT:
        return _rgba("#0d1424", min(0.9, alpha * 1.3))
    return _rgba("#ffffff", alpha)


def _sunken(color, amount=0.25):
    """«Утопленная» поверхность — поля ввода и таблицы.

    На тёмной теме она темнее подложки (углубление), на светлой — наоборот,
    остаётся чистой белой. Если бы формула была общей, на светлой теме поля
    и таблицы залились бы грязно-серым (белый, смешанный с чёрным).
    """
    if IS_LIGHT:
        return _mix(color, "#ffffff", min(1.0, amount * 0.6))
    return _mix(color, "#000000", amount)


def _press(color, amount=0.35):
    """Цвет нажатой кнопки: на тёмной теме заметно темнее, на светлой —
    лёгкое холодное затемнение вместо провала в серый."""
    if IS_LIGHT:
        return _mix(color, "#0d1424", min(1.0, amount * 0.28))
    return _mix(color, "#000000", amount)


def _sheen(base, hi=0.06, lo=0.0):
    """«Стеклянная» кнопка: светлый блик сверху, лёгкая тень снизу —
    ровно то, за счёт чего плоские кнопки выглядят объёмными и дорогими."""
    shade = 0.055 if IS_LIGHT else 0.10
    return (f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
            f"stop:0 {_rgba('#ffffff', hi if not IS_LIGHT else hi * 1.6)}, "
            f"stop:0.45 {_rgba('#ffffff', lo)}, "
            f"stop:1 {_rgba('#000000', shade)}), {base}")


def _shift_grad(light, mid, deep, shift=0.0):
    """Диагональный градиент светлое→тёмное со сдвигаемой светлой зоной.
    На нём построены и заливка акцентных кнопок, и «живой» блик при
    наведении: цвета не меняются, двигается только светлая зона."""
    a = max(0.0, min(1.0, 0.18 + shift))
    b = max(a + 0.01, min(0.99, 0.62 + shift))
    return (f"qlineargradient(x1:0, y1:0, x2:1, y2:1, "
            f"stop:0 {light}, stop:{a:.3f} {mid}, "
            f"stop:{b:.3f} {_mix(mid, deep, 0.55)}, stop:1 {deep})")


# Нижний (тёмный) стоп акцентного градиента.
# Если акцент сам по себе светлый (сталь, янтарь), текст на кнопке становится
# тёмным — и на глубоко-тёмном низу градиента он был бы нечитаем. Поэтому для
# светлых акцентов низ притемняем лишь слегка.
ACCENT_GRAD_DEEP = (ACCENT_DEEP if _luma(ACCENT) <= 0.5
                    else _mix(ACCENT, ACCENT_DEEP, 0.30))


def _accent_grad(shift=0.0):
    """Заливка акцентных элементов — диагональный перелив: тёмно-синий угол
    внизу справа плавно переходит в светлую лазурь вверху слева.

    shift сдвигает светлую зону вдоль градиента: на этом построена «живая»
    подсветка кнопки при наведении (см. ShimmerFilter) — цвета не меняются,
    двигается только сам блик."""
    return _shift_grad(ACCENT_HV, ACCENT, ACCENT_GRAD_DEEP, shift)


def _accent_grad_hover():
    """Тот же перелив, но заметно светлее — состояние наведения."""
    return (f"qlineargradient(x1:0, y1:0, x2:1, y2:1, "
            f"stop:0 {_mix(ACCENT_HV, '#ffffff', 0.35)}, stop:0.35 {ACCENT_HV}, "
            f"stop:0.75 {ACCENT}, stop:1 {_mix(ACCENT, ACCENT_GRAD_DEEP, 0.5)})")


def _green_grad():
    return (f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
            f"stop:0 {GREEN_HV}, stop:1 {GREEN})")


def _preview_panel_qss(bg=None):
    """Стиль панели предпросмотра (вкладка «Очистка»).

    Селектор QFrame#preview_panel обязателен: без него Qt применил бы рамку
    и скругление ко всем подписям внутри панели, и вокруг каждой из них
    появилась бы своя маленькая рамка-квадратик.
    """
    return (f"QFrame#preview_panel {{ background-color:{bg or PANEL_BG};"
            f" border:1px solid {PANEL_BORDER}; border-radius:14px; }}")


def _card_bg():
    """Подложка карточки контента: градиент с еле уловимым синим свечением
    из левого верхнего угла — «дорогая» глубина вместо плоской заливки.
    На светлой теме затемнение книзу сильно мягче, иначе белая карточка
    выглядела бы грязной."""
    dark_stop = 0.05 if IS_LIGHT else 0.25
    return (f"qlineargradient(x1:0, y1:0, x2:0.9, y2:1, "
            f"stop:0 {_mix(BG3, ACCENT, 0.045)}, stop:0.55 {BG3}, "
            f"stop:1 {_mix(BG3, '#000000', dark_stop)})")


def _panel_bg():
    """Подложка бокового меню — вертикальный градиент с лёгким затемнением книзу."""
    dark_stop = 0.05 if IS_LIGHT else 0.4
    return (f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
            f"stop:0 {_mix(NAV_BG, ACCENT, 0.035)}, stop:0.4 {NAV_BG}, "
            f"stop:1 {_mix(NAV_BG, '#000000', dark_stop)})")


def _rel_lum(hex_color):
    """Относительная яркость по формуле WCAG (не то же самое, что «на глаз»)."""
    def ch(v):
        v /= 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = _hex_rgb(hex_color)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def _contrast(c1, c2):
    """Коэффициент контраста двух цветов (1 — одинаковые, 21 — максимум)."""
    l1, l2 = _rel_lum(c1), _rel_lum(c2)
    if l1 < l2:
        l1, l2 = l2, l1
    return (l1 + 0.05) / (l2 + 0.05)


def _best_text_color(bg_hex, dark="#141414", light="#f7f7f8"):
    """Выбирает читаемый цвет текста для сплошной заливки.

    Сравниваются реальные коэффициенты контраста, а не яркость «по порогу»:
    на средних по светлоте акцентах (бирюза, сталь) простой порог ошибался и
    ставил белый текст там, где тёмный читается почти втрое лучше.
    """
    return dark if _contrast(dark, bg_hex) >= _contrast(light, bg_hex) else light


# Текст на синей заливке — светлый: сама заливка тёмная, тёмный текст на ней
# был бы нечитаем (в янтарной версии оформления было наоборот).
ON_ACCENT = _best_text_color(_mix(ACCENT, ACCENT_DEEP, 0.18),
                            dark="#0b1225", light="#f4f8ff")
# Подложка вложенных панелек внутри карточки контента. Отдельный оттенок нужен
# потому, что сама карточка уже покрашена в BG3: панель того же цвета на ней
# просто растворялась бы, а стеклянная (чуть светлее + волосяная рамка) даёт
# ту самую многослойность, из-за которой интерфейс читается как «дорогой».
PANEL_BG = "rgba(255,255,255,0.045)"
PANEL_BORDER = "rgba(255,255,255,0.07)"
FONT_STACK = f"'{STYLE['font']}', '{STYLE.get('font_fallback', 'Segoe UI')}'"


def _ui_font(size=10.5, weight=None, family=None):
    """Шрифт интерфейса с правильными настройками сглаживания.

    Без явного PreferAntialias/PreferQuality Qt на Windows включает жёсткий
    хинтинг — тонкие штрихи Manrope при этом «прилипают» к пиксельной сетке
    и текст выглядит рваным, как будто пиксельным. Отключённый хинтинг даёт
    ровные, «печатные» буквы на любом масштабе экрана."""
    f = QFont(family or STYLE["font"])
    f.setPointSizeF(float(size))
    if weight is not None:
        f.setWeight(weight)
    f.setStyleStrategy(QFont.PreferAntialias | QFont.PreferQuality)
    f.setHintingPreference(QFont.PreferNoHinting)
    return f


APP_NAME = "Base"


def C_clean_transcript(text):
    """Оставляет только реплики клиента, убирая строки робота/оператора и таймкоды."""
    if text is None:
        return ""
    out = []
    for raw in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw.strip()
        if not line:
            continue
        low = line.lower()
        if "робот" in low or "оператор" in low:
            continue
        line = re.sub(r"^\s*\d{1,2}:\d{2}(:\d{2})?\s*", "", line)
        line = re.sub(r"^(клиент|client|абонент|человек)\s*:\s*", "", line, flags=re.IGNORECASE)
        line = line.strip(" -–—:•").strip()
        if line:
            out.append(line)
    return "   ·   ".join(out)


# ══════════════════════════════════════════════════════════════════════════════
#                         ОФОРМЛЕНИЕ  ·  VORTEX OBSIDIAN
#   Тёмный обсидиан + янтарь. Тонкие волосяные рамки, стеклянные блики,
#   крупные скругления, много воздуха, заглавные разрядки в микро-подписях.
# ══════════════════════════════════════════════════════════════════════════════
def build_qss():
    R_SM, R_MD, R_LG = STYLE["radius_sm"], STYLE["radius_md"], STYLE["radius_lg"]
    ICR = STYLE["icon_btn_radius"]
    BW = STYLE["border_w"]
    BTNW = STYLE["btn_weight"]

    # Волосяные рамки строятся не из BORDER, а из полупрозрачного белого:
    # такая рамка одинаково аккуратно ложится и на карточку, и на подложку,
    # и не «чернеет» на градиентах.
    HAIR = _veil(0.07)
    HAIR_HI = _veil(0.13)
    GHOST = _sheen(_vgrad(GRAY_HV, GRAY_BTN))

    return f"""
    /* ─────────────────────────  БАЗА  ───────────────────────── */
    QWidget {{
        background-color: transparent;
        color: {FG};
        font-family: {FONT_STACK};
        font-size: 10.5pt;
    }}
    QMainWindow, QDialog {{ background-color: {BG}; }}
    QWidget#root_bg {{
        background-color: {BG2};
        border: 1px solid {_veil(0.10)};
    }}
    QLabel {{ background: transparent; }}
    QLabel#h1 {{ font-size: 16pt; font-weight: 800; color: {FG}; letter-spacing: 0.3px; }}
    QLabel#h2 {{ font-size: 11pt; font-weight: 700; color: {FG}; }}
    QLabel#muted {{ color: {FG3}; font-size: 9pt; }}
    QLabel#eyebrow {{
        color: {FG3}; font-size: 8pt; font-weight: 800;
        letter-spacing: 2.2px; text-transform: uppercase;
    }}
    QLabel#chip {{
        color: {ACCENT_SOFT};
        background-color: {_rgba(ACCENT, 0.10)};
        border: {BW}px solid {_rgba(ACCENT, 0.28)};
        border-radius: 9px; padding: 3px 10px;
        font-size: 9pt; font-weight: 700;
    }}
    QToolTip {{
        background-color: {_mix(BG3, "#ffffff", 0.05)}; color: {FG};
        border: {BW}px solid {_rgba(ACCENT, 0.45)};
        border-radius: 8px; padding: 6px 9px;
    }}

    /* ─────────────────────  ПОЛЯ ВВОДА  ───────────────────── */
    QLineEdit, QTextEdit, QPlainTextEdit {{
        background-color: {_sunken(SURFACE, 0.25)};
        border: {BW}px solid {HAIR};
        border-radius: {R_MD}px;
        padding: 8px 12px;
        color: {FG};
        selection-background-color: {_rgba(ACCENT, 0.45)};
        selection-color: {FG};
    }}
    QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover {{ border: {BW}px solid {HAIR_HI}; }}
    QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
        border: {BW}px solid {_rgba(ACCENT, 0.75)};
        background-color: {_mix(SURFACE, ACCENT, 0.05)};
    }}
    QLineEdit::placeholder {{ color: {FG3}; }}
    QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled {{ color: {FG3}; }}

    /* ─────────────────────────  КНОПКИ  ───────────────────────── */
    /* Базовая («призрачная») кнопка: стеклянный блик сверху, волосяная рамка. */
    QPushButton {{
        background: {GHOST};
        color: {FG};
        border: {BW}px solid {HAIR};
        border-radius: {R_MD}px;
        padding: 8px 15px;
        font-weight: {BTNW};
    }}
    QPushButton:hover {{
        background: {_glow(ACCENT, 0.22, 0.10)}, {GHOST};
        border: {BW}px solid {_rgba(ACCENT, 0.55)};
        color: {ACCENT_SOFT};
    }}
    QPushButton:pressed {{
        background: {_press(GRAY_BTN, 0.35)};
        border: {BW}px solid {_rgba(ACCENT, 0.85)};
        padding-top: 9px; padding-bottom: 7px;   /* «продавливание» кнопки */
    }}
    QPushButton:disabled {{ color: {FG3}; border: {BW}px solid {_veil(0.04)}; }}

    /* Главная кнопка действия */
    QPushButton#accent {{
        background: {_accent_grad()};
        color: {ON_ACCENT};
        border: {BW}px solid {_rgba("#ffffff", 0.22)};
        border-radius: {R_MD}px;
        font-weight: 800;
        padding: 9px 18px;
    }}
    QPushButton#accent:hover {{
        background: {_accent_grad_hover()};
        border: {BW}px solid {_rgba("#ffffff", 0.40)};
        color: {ON_ACCENT};
    }}
    QPushButton#accent:pressed {{ background: {ACCENT_DEEP}; padding-top: 10px; padding-bottom: 8px; }}
    QPushButton#accent:disabled {{ background: {_rgba(ACCENT, 0.22)}; color: {_rgba("#ffffff", 0.45)}; }}

    QPushButton#green {{
        background: {_green_grad()};
        color: {_best_text_color(GREEN, dark="#06231a", light="#f0fff8")};
        border: {BW}px solid {_rgba("#ffffff", 0.22)};
        border-radius: {R_MD}px;
        font-weight: 800;
        padding: 9px 18px;
    }}
    QPushButton#green:hover {{ background: {_vgrad(_mix(GREEN_HV, "#ffffff", 0.2), GREEN_HV)}; }}
    QPushButton#green:pressed {{ background: {GREEN}; }}

    QPushButton#danger {{
        background: {_rgba(DUPE_RED, 0.12)};
        color: {DANGER_FG};
        border: {BW}px solid {_rgba(DUPE_RED, 0.35)};
        border-radius: {R_MD}px;
    }}
    QPushButton#danger:hover {{
        background: {_glow(DUPE_RED, 0.30, 0.16)}, {_rgba(DUPE_RED, 0.16)};
        color: {DUPE_RED_HV};
        border: {BW}px solid {_rgba(DUPE_RED_HV, 0.75)};
    }}
    QPushButton#danger:pressed {{ background: {_rgba(DUPE_RED, 0.30)}; }}

    /* Сегментированный переключатель — «таблетки» в общей канавке */
    QPushButton#seg {{
        background-color: {_veil(0.03)};
        color: {FG2};
        border: {BW}px solid transparent;
        border-radius: {R_SM}px;
        padding: 7px 14px;
        font-weight: {BTNW};
    }}
    QPushButton#seg:hover {{ color: {FG}; background-color: {_rgba(ACCENT, 0.12)}; }}
    QPushButton#seg:checked {{
        background: {_accent_grad()};
        color: {ON_ACCENT};
        border: {BW}px solid {_rgba("#ffffff", 0.20)};
        font-weight: 800;
    }}

    /* Кнопка с выпадающим списком */
    QToolButton#dropdown_btn {{
        background: {GHOST};
        color: {FG};
        border: {BW}px solid {HAIR};
        border-radius: {R_MD}px;
        padding: 7px 14px;
        font-weight: {BTNW};
    }}
    QToolButton#dropdown_btn::menu-indicator {{ image: none; width: 0; }}
    QToolButton#dropdown_btn:hover {{
        background: {_glow(ACCENT, 0.20, 0.09)}, {GHOST};
        border: {BW}px solid {_rgba(ACCENT, 0.6)};
        color: {ACCENT_SOFT};
    }}
    QToolButton#dropdown_btn:pressed {{ background-color: {BG2}; border: {BW}px solid {ACCENT}; }}

    /* Кнопки-вкладки внутри страниц */
    QPushButton#tab {{
        background: transparent;
        color: {FG2};
        border: {BW}px solid {HAIR};
        border-radius: {R_LG}px;
        padding: 7px 18px;
        font-weight: {BTNW};
    }}
    QPushButton#tab:hover {{ border: {BW}px solid {_rgba(ACCENT, 0.6)}; color: {FG}; }}
    QPushButton#tab:checked {{
        background: {_accent_grad()};
        color: {ON_ACCENT};
        border: {BW}px solid {_rgba("#ffffff", 0.2)};
    }}

    /* Компактные квадратные кнопки-иконки */
    QPushButton#icon_btn, QPushButton#icon_btn_danger, QPushButton#icon_btn_green {{
        background-color: {_veil(0.04)};
        border: {BW}px solid {HAIR};
        border-radius: {ICR}px;
        font-size: 13pt;
        padding: 0px;
        color: {FG};
    }}
    QPushButton#icon_btn:hover {{
        background: {_glow_ring(ACCENT)};
        border: {BW}px solid {_rgba(ACCENT, 0.8)};
        color: {ACCENT_SOFT};
    }}
    QPushButton#icon_btn:pressed {{ background-color: {_rgba(ACCENT, 0.30)}; }}
    QPushButton#icon_btn_danger {{ color: {DUPE_RED}; }}
    QPushButton#icon_btn_danger:hover {{
        background: {_glow_ring(DUPE_RED)};
        border: {BW}px solid {_rgba(DUPE_RED, 0.8)};
        color: {DUPE_RED_HV};
    }}
    QPushButton#icon_btn_danger:pressed {{ background-color: {_rgba(DUPE_RED, 0.30)}; }}
    QPushButton#icon_btn_green {{ color: {GREEN}; }}
    QPushButton#icon_btn_green:hover {{
        background: {_glow_ring(GREEN)};
        border: {BW}px solid {_rgba(GREEN, 0.8)};
        color: {GREEN_HV};
    }}
    QPushButton#icon_btn_green:pressed {{ background-color: {_rgba(GREEN, 0.28)}; }}

    /* ────────────────  КНОПКИ УПРАВЛЕНИЯ ОКНОМ  ──────────────── */
    QPushButton#win_btn, QPushButton#win_close {{
        background: transparent;
        border: none;
        border-radius: 8px;
        padding: 0px;
        color: {FG2};
    }}
    QPushButton#win_btn:hover {{ background-color: {_veil(0.09)}; color: {FG}; }}
    QPushButton#win_btn:pressed {{ background-color: {_veil(0.15)}; }}
    QPushButton#win_close:hover {{ background-color: {DUPE_RED}; color: #ffffff; }}
    QPushButton#win_close:pressed {{ background-color: {_mix(DUPE_RED, "#000000", 0.25)}; }}

    /* ─────────────────────────  ТАБЛИЦЫ  ───────────────────────── */
    QTableWidget {{
        background-color: {_sunken(SURFACE, 0.35)};
        alternate-background-color: {ROW_ODD};
        gridline-color: transparent;
        border: {BW}px solid {HAIR};
        border-radius: {R_LG}px;
        selection-background-color: {_rgba(ACCENT, 0.20)};
        selection-color: {FG};
        outline: none;
    }}
    QTableWidget::item {{
        padding: 6px 8px;
        border: none;
        border-bottom: 1px solid {_veil(0.05)};
    }}
    QTableWidget::item:selected {{
        background-color: {_rgba(ACCENT, 0.20)};
        color: {FG};
    }}
    QTableWidget::item:focus {{ outline: none; }}
    QHeaderView {{ background: transparent; }}
    QHeaderView::section {{
        background-color: {_mix(BG2, "#ffffff", 0.02)};
        color: {FG3};
        padding: 9px 8px;
        border: none;
        border-bottom: {BW}px solid {_rgba(ACCENT, 0.28)};
        font-weight: 800;
        font-size: 8.5pt;
        letter-spacing: 1.4px;
    }}
    QHeaderView::section:hover {{ color: {ACCENT_SOFT}; }}
    QTableCornerButton::section {{ background-color: {BG2}; border: none; }}

    /* ────────────────────────  ПРОГРЕСС  ──────────────────────── */
    QProgressBar {{
        background-color: {_veil(0.05)};
        border: none;
        border-radius: 4px;
        height: 7px;
        text-align: center;
        color: transparent;
    }}
    QProgressBar::chunk {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT_DEEP}, stop:0.45 {ACCENT}, stop:0.75 {ACCENT_HV}, stop:1 {GREEN});
        border-radius: 4px;
    }}

    /* ───────────────────────  ПОЛОСЫ ПРОКРУТКИ  ─────────────────────── */
    QScrollBar:vertical {{ background: transparent; width: 11px; margin: 4px 2px 4px 0px; }}
    QScrollBar::handle:vertical {{
        background: {_veil(0.13)};
        border-radius: 4px; min-height: 34px;
    }}
    QScrollBar::handle:vertical:hover {{ background: {_rgba(ACCENT, 0.75)}; }}
    QScrollBar:horizontal {{ background: transparent; height: 11px; margin: 0px 4px 2px 4px; }}
    QScrollBar::handle:horizontal {{
        background: {_veil(0.13)};
        border-radius: 4px; min-width: 34px;
    }}
    QScrollBar::handle:horizontal:hover {{ background: {_rgba(ACCENT, 0.75)}; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
    QScrollArea {{ background: transparent; border: none; }}

    /* ─────────────────────────  МЕНЮ  ───────────────────────── */
    QMenu {{
        background-color: {_mix(BG3, "#ffffff", 0.02)};
        color: {FG};
        border: {BW}px solid {HAIR_HI};
        border-radius: {R_MD}px;
        padding: 7px;
    }}
    QMenu::item {{
        padding: 8px 18px 8px 14px;
        border-radius: {R_SM}px;
        color: {FG};
    }}
    QMenu::item:selected {{ background-color: {_rgba(ACCENT, 0.18)}; color: {ACCENT_SOFT}; }}
    QMenu::separator {{ height: 1px; background: {_veil(0.08)}; margin: 6px 8px; }}
    QMenu::icon {{ padding-left: 8px; }}

    /* ───────────────────────  ЧЕКБОКСЫ / СПИСКИ  ─────────────────────── */
    QCheckBox {{ spacing: 9px; background: transparent; }}
    QCheckBox::indicator {{
        width: 17px; height: 17px;
        border: {BW}px solid {_veil(0.18)};
        border-radius: 6px;
        background-color: {_veil(0.04)};
    }}
    QCheckBox::indicator:hover {{ border: {BW}px solid {ACCENT}; background-color: {_rgba(ACCENT, 0.14)}; }}
    QCheckBox::indicator:checked {{
        background: {_accent_grad()};
        border: {BW}px solid {_rgba("#ffffff", 0.25)};
    }}
    QComboBox {{
        background: {GHOST};
        color: {FG};
        border: {BW}px solid {HAIR};
        border-radius: {R_MD}px;
        padding: 7px 12px;
    }}
    QComboBox:hover {{ border: {BW}px solid {_rgba(ACCENT, 0.6)}; }}
    QComboBox::drop-down {{ border: none; width: 22px; }}
    QComboBox QAbstractItemView {{
        background-color: {_mix(BG3, "#ffffff", 0.02)};
        border: {BW}px solid {HAIR_HI};
        border-radius: {R_SM}px;
        selection-background-color: {_rgba(ACCENT, 0.20)};
        selection-color: {ACCENT_SOFT};
        padding: 4px;
        outline: none;
    }}
    QMessageBox {{ background-color: {BG3}; }}
    QMessageBox QLabel {{ color: {FG}; font-size: 10.5pt; }}
    QMessageBox QPushButton {{ min-width: 96px; padding: 8px 16px; }}

    /* ═══════════════════  КАРКАС ОКНА  ═══════════════════ */
    /* Своя полоса заголовка вместо системной */
    QFrame#titlebar {{
        background-color: {BG2};
        border: none;
        border-bottom: 1px solid {_veil(0.05)};
    }}
    QLabel#wordmark {{
        color: {FG}; font-size: 12pt; font-weight: 800; letter-spacing: 3px;
    }}
    QLabel#wordmark_sub {{
        color: {FG3}; font-size: 8pt; font-weight: 600; letter-spacing: 1.6px;
    }}

    /* Боковое меню */
    QFrame#nav_panel {{
        background: {_panel_bg()};
        border: none;
        border-right: 1px solid {_veil(0.05)};
    }}
    QFrame#icon_rail {{ background-color: {BG2}; border: none; }}
    QFrame#side_card {{
        background-color: {_veil(0.035)};
        border: {BW}px solid {HAIR};
        border-radius: {R_MD}px;
    }}
    QFrame#nav_sep {{ background-color: {_veil(0.07)}; border: none; }}

    /* Пункт бокового меню */
    QPushButton#navitem {{
        background-color: transparent;
        color: {FG2};
        border: {BW}px solid transparent;
        border-radius: {R_MD}px;
        margin: 2px 10px 2px 12px;
        text-align: left;
        padding: 11px 12px 11px 14px;
        font-size: 11pt;
        font-weight: 600;
        letter-spacing: 0.2px;
    }}
    QPushButton#navitem:hover {{
        background-color: {_veil(0.05)};
        color: {FG};
        border: {BW}px solid {_veil(0.06)};
    }}
    QPushButton#navitem:checked {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {_rgba(ACCENT, 0.20)}, stop:1 {_rgba(ACCENT, 0.02)});
        color: {FG};
        border: {BW}px solid {_rgba(ACCENT, 0.22)};
        font-weight: 800;
    }}

    /* Рабочая область */
    QWidget#content_area {{ background-color: {BG}; }}
    QFrame#topbar {{ background: transparent; border: none; }}
    QFrame#statusbar {{
        background-color: {BG2};
        border: none;
        border-top: 1px solid {_veil(0.05)};
    }}
    QFrame#content_card {{
        background: {_card_bg()};
        border: {BW}px solid {HAIR};
        border-radius: {R_LG}px;
    }}
    """


# ──────────────────────────  ЧИСЛОВОЙ ЭЛЕМЕНТ ТАБЛИЦЫ  ──────────────────────────
class FlowLayout(QLayout):
    """Раскладка, переносящая виджеты на новую строку, когда не хватает
    места по горизонтали (вместо обрезания текста или сжатия шрифта).
    Используется для компактного списка городов-«чипов» над таблицей на
    вкладке «Редактирование» — если городов много (хоть 20), они просто
    переносятся на несколько строк, все остаются видны целиком."""
    def __init__(self, parent=None, margin=0, h_spacing=6, v_spacing=6):
        super().__init__(parent)
        self._items = []
        self._h_spacing = h_spacing
        self._v_spacing = v_spacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index):
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), test_only=True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        m = self.contentsMargins()
        size += QSize(m.left() + m.right(), m.top() + m.bottom())
        return size

    def clear(self):
        while self._items:
            item = self._items.pop()
            w = item.widget()
            if w:
                w.deleteLater()

    def _do_layout(self, rect, test_only):
        x, y = rect.x(), rect.y()
        line_height = 0
        for item in self._items:
            w = item.widget()
            if w is None or not w.isVisible():
                continue
            item_w = item.sizeHint().width()
            item_h = item.sizeHint().height()
            next_x = x + item_w + self._h_spacing
            if next_x - self._h_spacing > rect.right() and line_height > 0:
                x = rect.x()
                y = y + line_height + self._v_spacing
                next_x = x + item_w + self._h_spacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), item.sizeHint()))
            x = next_x
            line_height = max(line_height, item_h)
        return y + line_height - rect.y()


class NumItem(QTableWidgetItem):
    """Ячейка таблицы с числовым ключом сортировки (для колонок «км», «Население»)."""
    def __init__(self, text, sort_value):
        super().__init__(text)
        self._sort = sort_value
        self.setFlags(self.flags() & ~Qt.ItemIsEditable)

    def __lt__(self, other):
        if isinstance(other, NumItem):
            return self._sort < other._sort
        return super().__lt__(other)


def txt_item(text, editable=False):
    it = QTableWidgetItem(str(text))
    if not editable:
        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    return it


# ──────────────────────────  СЕГМЕНТИРОВАННЫЙ ПЕРЕКЛЮЧАТЕЛЬ  ──────────────────────────
class HintPlainTextEdit(QPlainTextEdit):
    """QPlainTextEdit с подсказкой, которая рисуется по центру поля и
    пропадает, как только пользователь начинает печатать (аналог
    placeholder у QLineEdit, но с центрированием текста)."""

    def __init__(self, hint_text="", parent=None):
        super().__init__(parent)
        self._hint = hint_text
        # Qt по умолчанию перерисовывает только маленькую область вокруг
        # курсора при вводе текста — этого недостаточно, чтобы стереть
        # подсказку, нарисованную в центре поля. Поэтому при любом
        # изменении текста форсируем полную перерисовку viewport'а.
        self.textChanged.connect(self.viewport().update)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._hint and not self.toPlainText():
            p = QPainter(self.viewport())
            p.setRenderHint(QPainter.Antialiasing)
            p.setPen(QColor(FG2))
            p.setFont(self.font())
            rect = self.viewport().rect().adjusted(12, 0, -12, 0)
            p.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, self._hint)
            p.end()


class DropdownButton(QToolButton):
    """Компактная кнопка с выпадающим списком (меню всплывает прямо из
    кнопки, без отдельного окна) — замена сегментированному переключателю
    там, где нужно сэкономить место."""
    changed = Signal(object)

    def __init__(self, options, default=None, parent=None, allow_custom=False):
        # options: список кортежей (label, value)
        # allow_custom: добавить пункт «Своё число…» — пользователь вводит
        # любое целое число вручную.
        super().__init__(parent)
        self.setObjectName("dropdown_btn")
        self.setPopupMode(QToolButton.InstantPopup)
        self.setCursor(Qt.PointingHandCursor)
        self._labels = {value: label for label, value in options}
        self._current = default if default in self._labels else (options[0][1] if options else None)

        menu = QMenu(self)
        for label, value in options:
            act = menu.addAction(label)
            act.triggered.connect(lambda checked=False, v=value: self._select(v))
        if allow_custom:
            menu.addSeparator()
            act = menu.addAction("Своё число…")
            act.triggered.connect(self._ask_custom)
        self.setMenu(menu)
        self._update_text()

    def _ask_custom(self):
        from PySide6.QtWidgets import QInputDialog
        cur = self._current if isinstance(self._current, int) else 5000
        value, ok = QInputDialog.getInt(
            self.window(), "Своё число",
            "Сколько номеров в одном файле:",
            cur, 1, 100_000_000, 1000)
        if ok:
            if value not in self._labels:
                self._labels[value] = f"{value:,}".replace(",", " ")
            self._select(value)

    def _select(self, value):
        self._current = value
        self._update_text()
        self.changed.emit(value)

    def _update_text(self):
        self.setText(f"{self._labels.get(self._current, self._current)}  ▾")

    def value(self):
        return self._current


class Segmented(QWidget):
    """Группа кнопок-переключателей (как CTkSegmentedButton)."""
    changed = Signal(object)  # передаёт value выбранной опции

    def __init__(self, options, default=None, parent=None):
        # options: список кортежей (label, value)
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        self._values = {}
        for label, value in options:
            b = QPushButton(label)
            b.setObjectName("seg")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            self.group.addButton(b)
            self._values[b] = value
            lay.addWidget(b)
            if value == default:
                b.setChecked(True)
        self.group.buttonClicked.connect(
            lambda btn: self.changed.emit(self._values[btn]))
        # animate checked state on click for each seg button
        for b in self.group.buttons():
            b.installEventFilter(self)

    def value(self):
        b = self.group.checkedButton()
        return self._values.get(b)


# ──────────────────────────  СТРОКА ВЫБОРА ПАПКИ  ──────────────────────────
# ──────────────────────────  ИКОНКИ ДЛЯ СТРОК ПАПОК  ──────────────────────────
def _undo_icon(size=22, color=None):
    """Рисует иконку «вернуть назад» — закруглённая стрелка против часовой
    стрелки (как ⟲). Символ ↩ не во всех шрифтах отображается аккуратно
    (иногда рисуется просто прямоугольником), поэтому иконка рисуется вручную
    с точным расчётом геометрии дуги и наконечника стрелки."""
    import math
    from PySide6.QtGui import QPixmap, QIcon
    col = QColor(color) if color else QColor(FG)
    px = QPixmap(size, size)
    px.fill(QColor(0, 0, 0, 0))
    p = QPainter(px)
    p.setRenderHint(QPainter.Antialiasing)
    w = h = size
    cx, cy = w * 0.5, h * 0.5
    r = w * 0.30
    lw = max(1.6, w * 0.13)
    pen = QPen(col, lw, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)

    start_deg, span_deg = 250, 260
    rect = QRectF(cx - r, cy - r, r * 2, r * 2)
    p.drawArc(rect, int(start_deg * 16), int(span_deg * 16))

    # Наконечник стрелки строго на конце дуги, по касательной к окружности —
    # координаты вычислены из той же формулы, что использует QPainter.drawArc.
    end_rad = math.radians(start_deg + span_deg)
    ex, ey = cx + r * math.cos(end_rad), cy - r * math.sin(end_rad)
    dx, dy = -math.sin(end_rad), -math.cos(end_rad)   # касательная (направление увеличения угла)
    nx, ny = -dy, dx                                   # нормаль (ширина наконечника)

    L, W, tail = w * 0.12, w * 0.065, 0.5
    tip = (ex + dx * L, ey + dy * L)
    base = (ex - dx * L * tail, ey - dy * L * tail)
    left = (base[0] + nx * W, base[1] + ny * W)
    right = (base[0] - nx * W, base[1] - ny * W)

    p.setPen(Qt.NoPen)
    p.setBrush(col)
    arrow = QPainterPath()
    arrow.moveTo(*tip)
    arrow.lineTo(*left)
    arrow.lineTo(*right)
    arrow.closeSubpath()
    p.drawPath(arrow)
    p.end()
    return QIcon(px)


def _excel_icon(size=20):
    """Иконка Excel для кнопки выбора CSV-файлов: пробует qtawesome,
    иначе рисует зелёный значок с белым «X» вручную."""
    if _QTA:
        try:
            return qta.icon('mdi.file-excel', color='#1e7145')
        except Exception:
            try:
                return qta.icon('ri.file-excel-2-fill', color='#1e7145')
            except Exception:
                pass
    from PySide6.QtGui import QPixmap, QIcon
    px = QPixmap(size, size)
    px.fill(QColor(0, 0, 0, 0))
    p = QPainter(px)
    p.setRenderHint(QPainter.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(QRectF(1, 1, size - 2, size - 2), 4, 4)
    p.fillPath(path, QColor("#1e7145"))
    p.setPen(QPen(QColor("white")))
    f = QFont("Segoe UI", int(size * 0.55), QFont.Bold)
    p.setFont(f)
    p.drawText(px.rect(), Qt.AlignCenter, "X")
    p.end()
    return QIcon(px)


# Свежие векторные иконки (qtawesome) для строк выбора папок во вкладке «Очистка».
_ROW_ICON_MAP = {
    "dupes": ICONS["phone"],     # «Прозвоненные»
    "upload": ICONS["database"],  # «База для очистки»
    "check": ICONS["check"],     # «Результат»
}


class FolderRow(QWidget):
    changed = Signal()

    def __init__(self, label, label_color=None, icon_variant=None, parent=None, allow_files=False,
                 show_folder_btn=True):
        super().__init__(parent)
        self._selected_files = None  # реальные пути CSV-файлов (если выбраны через кнопку CSV)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        self.edit = QLineEdit()
        if label:
            self.edit.setPlaceholderText(label)
        if icon_variant:
            qname = _ROW_ICON_MAP.get(icon_variant)
            fallback = {"dupes": "☎", "upload": "🗄", "check": "✓"}.get(icon_variant, "")
            qic = _icon(qname, FG2, None) if qname else None  # приглушённый цвет (FG2), без окраски по смыслу
            if qic:
                self.edit.addAction(qic, QLineEdit.LeadingPosition)
            elif fallback:
                # запасной вариант без qtawesome — эмодзи в самом плейсхолдере
                self.edit.setPlaceholderText(f"{fallback}  {label}" if label else fallback)
        self.edit.textChanged.connect(lambda *_: self.changed.emit())
        # Если пользователь сам стирает/печатает в поле — забываем список файлов,
        # чтобы get() возвращал именно то, что реально написано.
        self.edit.textEdited.connect(self._on_manual_edit)
        lay.addWidget(self.edit, 1)
        if show_folder_btn:
            _folder_ic = _icon(ICONS["folder"], FG, "")
            btn = QPushButton("" if _folder_ic else "…")
            if _folder_ic:
                btn.setIcon(_folder_ic)
                btn.setIconSize(QSize(15, 15))
            btn.setFixedWidth(38)
            btn.setToolTip("Выбрать папку")
            btn.clicked.connect(self._browse)
            lay.addWidget(btn)
        self.allow_files = allow_files
        if allow_files:
            csv_btn = QPushButton()
            csv_btn.setIcon(_excel_icon(17))
            csv_btn.setIconSize(QSize(17, 17))
            csv_btn.setFixedWidth(38)
            csv_btn.setToolTip("Выбрать один или несколько Excel/CSV-файлов")
            csv_btn.clicked.connect(self._browse_files)
            lay.addWidget(csv_btn)

    def _on_manual_edit(self, _text):
        self._selected_files = None

    def _browse(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            self._selected_files = None
            self.edit.setText(path)
            self.edit.setToolTip(path)

    def _browse_files(self):
        # Всегда открываем диалог начиная с папки «Загрузки» — большинство
        # выгруженных файлов оказываются именно там, так удобнее, чем каждый
        # раз переходить туда вручную (и чем открываться в последней папке,
        # где что-то выбирали для другого поля).
        start_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.isdir(start_dir):
            start_dir = os.path.expanduser("~")
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите Excel/CSV-файл(ы) с прозвоненными номерами",
            start_dir, "Excel/CSV файлы (*.xlsx *.xlsm *.xls *.csv);;Все файлы (*.*)")
        if files:
            self._selected_files = files
            self.edit.setText(self._files_label(len(files)))
            self.edit.setToolTip("\n".join(files))

    @staticmethod
    def _files_label(n):
        """Русское склонение: 1 файл / 2-4 файла / 5+ файлов."""
        if 11 <= n % 100 <= 14:
            word = "файлов"
        else:
            last = n % 10
            if last == 1:
                word = "файл"
            elif 2 <= last <= 4:
                word = "файла"
            else:
                word = "файлов"
        return f"Выбрано {n} {word}"

    def get(self):
        if self._selected_files:
            return "; ".join(self._selected_files)
        return self.edit.text().strip()

    def set(self, v):
        self._selected_files = None
        self.edit.setText(v)


class FolderButton(QPushButton):
    """Кнопка выбора папки без отдельного поля ввода. Путь хранится внутри,
    а на самой кнопке отображается имя выбранной папки (укороченно)."""
    changed = Signal()

    def __init__(self, placeholder="Выбрать папку", parent=None):
        super().__init__(placeholder, parent)
        self._path = ""
        self._placeholder = placeholder
        self.clicked.connect(self._browse)

    def _browse(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            self.set(path)
            self.changed.emit()

    def get(self):
        return self._path

    def set(self, v):
        self._path = v or ""
        if self._path:
            name = os.path.basename(self._path.rstrip("/\\")) or self._path
            self.setText(f" {name}")
            self.setIcon(_icon(ICONS["folder"], FG, ""))
            self.setIconSize(QSize(15, 15))
            self.setToolTip(self._path)
        else:
            self.setText(self._placeholder)
            self.setIcon(QIcon())
            self.setToolTip("")


# ──────────────────────────  ЗНАК ПРОГРАММЫ (ЛОГОТИП)  ──────────────────────────
class _DbBadge(QWidget):
    """Знак программы — стопка из трёх пластин в скруглённой плитке.

    Пластины сужаются книзу (силуэт базы данных), и по ним непрерывно течёт
    световая волна: тёмная синь переливается в светлую лазурь и обратно.
    При наведении поток ускоряется, плитка подсвечивается, стопка чуть
    раздвигается, вокруг появляется мягкий ореол.

    Всё рисуется вектором (QPainter), поэтому знак остаётся резким при любом
    масштабе экрана и не зависит от наличия библиотеки иконок.
    """

    clicked = Signal()

    def __init__(self, parent=None, size=44):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("Быстрый переход  ·  Ctrl+K")
        self.setAttribute(Qt.WA_Hover, True)

        self._phase = -0.35     # положение световой волны
        self._hover = 0.0       # 0..1 — плавная реакция на курсор

        # Перерисовка по таймеру (25 кадров/с) вместо анимации на каждый
        # кадр экрана: знак маленький, а нагрузка получается символической.
        self._timer = QTimer(self)
        self._timer.setInterval(40)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

        self._hover_anim = QVariantAnimation(self)
        self._hover_anim.setDuration(260)
        self._hover_anim.setEasingCurve(QEasingCurve.OutCubic)
        self._hover_anim.valueChanged.connect(self._on_hover)

    def _tick(self):
        self._phase += 0.011 + 0.020 * self._hover
        if self._phase > 1.35:
            self._phase = -0.35
        self.update()

    def _on_hover(self, v):
        self._hover = float(v)
        self.update()

    def _hover_to(self, target):
        self._hover_anim.stop()
        self._hover_anim.setStartValue(self._hover)
        self._hover_anim.setEndValue(target)
        self._hover_anim.start()

    def enterEvent(self, event):
        self._hover_to(1.0)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover_to(0.0)
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)

    @staticmethod
    def _wave(x0, x1, phase, base, light):
        """Градиент со светлой волной в позиции phase (стопы обрезаны в 0..1)."""
        def cl(v):
            return min(1.0, max(0.0, v))

        g = QLinearGradient(x0, 0, x1, 0)
        g.setColorAt(0.0, QColor(base))
        g.setColorAt(cl(phase - 0.30), QColor(base))
        g.setColorAt(cl(phase), QColor(light))
        g.setColorAt(cl(phase + 0.30), QColor(base))
        g.setColorAt(1.0, QColor(base))
        return g

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        r = min(w, h)
        hv = self._hover

        # ── ореол при наведении ──
        if hv > 0.01:
            halo = QColor(ACCENT_HV)
            halo.setAlphaF(0.22 * hv)
            p.setPen(Qt.NoPen)
            p.setBrush(halo)
            p.drawRoundedRect(QRectF(0, 0, w, h), r * 0.34, r * 0.34)

        # ── плитка ──
        inset = r * 0.05
        tile = QPainterPath()
        tile.addRoundedRect(QRectF(inset, inset, w - inset * 2, h - inset * 2),
                            r * 0.30, r * 0.30)
        grad = QLinearGradient(0, 0, w, h)
        grad.setColorAt(0.0, QColor(_mix(BG3, ACCENT_HV, 0.10 + 0.20 * hv)))
        grad.setColorAt(1.0, QColor(_mix(BG2, "#000000", 0.05 if IS_LIGHT else 0.28)))
        p.setPen(Qt.NoPen)
        p.fillPath(tile, QBrush(grad))

        edge = QColor(ACCENT_HV)
        edge.setAlphaF(0.20 + 0.65 * hv)
        pen = QPen(edge)
        pen.setWidthF(1.2)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawPath(tile)

        # ── три пластины ──
        p.setPen(Qt.NoPen)
        plate_h = r * 0.125
        gap = r * 0.085 + r * 0.035 * hv
        widths = (0.56, 0.42, 0.28)
        total = len(widths) * plate_h + (len(widths) - 1) * gap
        y = (h - total) / 2.0
        for i, wf in enumerate(widths):
            pw = r * wf
            x = (w - pw) / 2.0
            g = self._wave(x, x + pw, self._phase - i * 0.16,
                           _mix(ACCENT, ACCENT_DEEP, 0.45), ACCENT_HV)
            p.setBrush(QBrush(g))
            p.drawRoundedRect(QRectF(x, y, pw, plate_h), plate_h / 2.0, plate_h / 2.0)
            y += plate_h + gap
        p.end()


# ──────────────────────────  ПЛАВНЫЙ ТУЛТИП  ──────────────────────────
class BubbleTip(QFrame):
    """Плавно появляющийся тултип-пузырёк. Singleton — один на всё приложение."""
    _shared: "BubbleTip | None" = None

    @classmethod
    def shared(cls) -> "BubbleTip":
        if cls._shared is None:
            cls._shared = cls()
        return cls._shared

    def __init__(self):
        super().__init__(None,
                         Qt.ToolTip | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self._lbl = QLabel("", self)
        self._lbl.setAlignment(Qt.AlignCenter)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 7, 12, 7)
        lay.addWidget(self._lbl)
        self.setStyleSheet(
            f"QFrame {{ background-color:{_mix(BG3, '#ffffff', 0.05)};"
            f"         border:1px solid {_rgba(ACCENT, 0.55)};"
            "          border-radius:11px; }"
            f"QLabel {{ color:{FG}; font-size:10.5pt; font-weight:600;"
            "          background:transparent; border:none; }}"
        )
        self._eff = QGraphicsOpacityEffect(self)
        self._eff.setOpacity(0.0)
        self.setGraphicsEffect(self._eff)
        self._anim = QPropertyAnimation(self._eff, b"opacity", self)
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)

    def _disconnect_anim(self):
        try:
            self._anim.finished.disconnect()
        except RuntimeError:
            pass

    def show_for(self, btn: QWidget, prefer="above"):
        self._lbl.setText(btn.toolTip())
        self.adjustSize()

        win = btn.window()
        win_left = win_top = win_right = win_bottom = None
        if win is not None:
            top_left = win.mapToGlobal(QPoint(0, 0))
            win_left, win_top = top_left.x(), top_left.y()
            win_right = win_left + win.width()
            win_bottom = win_top + win.height()

        gp_top = btn.mapToGlobal(QPoint(btn.width() // 2, 0))
        x_above = gp_top.x() - self.width() // 2
        y_above = gp_top.y() - self.height() - 14

        gp_bot = btn.mapToGlobal(QPoint(btn.width() // 2, btn.height() + 14))
        x_below = gp_bot.x() - self.width() // 2
        y_below = gp_bot.y()

        # Ограничиваем тултип границами ОКНА ПРОГРАММЫ, а не всего монитора —
        # иначе при длинном тексте подсказка «вылезала» за пределы приложения.
        if win_top is not None:
            above_fits = y_above >= win_top + 4
            below_fits = (y_below + self.height()) <= win_bottom - 4
            if prefer == "below":
                x, y = (x_below, y_below) if below_fits else (x_above, y_above)
            else:
                x, y = (x_above, y_above) if above_fits else (x_below, y_below)
            x = max(win_left + 4, min(x, win_right - self.width() - 4))
        else:
            x, y = (x_below, y_below) if prefer == "below" else (x_above, y_above)
            screen = btn.screen() if hasattr(btn, "screen") else None
            if screen is not None:
                geo = screen.availableGeometry()
                x = max(geo.left() + 4, min(x, geo.right() - self.width() - 4))
        self.move(x, y)
        self._disconnect_anim()
        self._anim.stop()
        if not self.isVisible():
            self.show()
        self._anim.setStartValue(float(self._eff.opacity()))
        self._anim.setEndValue(1.0)
        self._anim.start()

    def hide_tip(self):
        self._disconnect_anim()
        self._anim.stop()
        self._anim.setStartValue(float(self._eff.opacity()))
        self._anim.setEndValue(0.0)

        def _done():
            self._disconnect_anim()
            if self._eff.opacity() < 0.05:
                self.hide()

        self._anim.finished.connect(_done)
        self._anim.start()


# ──────────────────────────  МЕЛКИЕ ДЕКОРАТИВНЫЕ ЭЛЕМЕНТЫ  ──────────────────────────
class PulseDot(QWidget):
    """Пульсирующая точка-индикатор (живой «пульс» ядра). Рисуется вручную:
    сама точка + мягкий ореол, который медленно разгорается и гаснет."""

    def __init__(self, color=None, size=12, parent=None):
        super().__init__(parent)
        self._color = color or GREEN
        self.setFixedSize(size, size)
        self._t = 0.0
        up = QVariantAnimation(self)
        up.setStartValue(0.0)
        up.setEndValue(1.0)
        up.setDuration(1100)
        up.setEasingCurve(QEasingCurve.InOutSine)
        down = QVariantAnimation(self)
        down.setStartValue(1.0)
        down.setEndValue(0.0)
        down.setDuration(1600)
        down.setEasingCurve(QEasingCurve.InOutSine)
        for a in (up, down):
            a.valueChanged.connect(self._on_v)
        self._anim = QSequentialAnimationGroup(self)
        self._anim.addAnimation(up)
        self._anim.addAnimation(down)
        self._anim.setLoopCount(-1)
        self._anim.start()

    def _on_v(self, v):
        self._t = float(v)
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor(self._color)
        cx, cy = self.width() / 2.0, self.height() / 2.0
        halo = QColor(c)
        halo.setAlphaF(0.10 + 0.30 * self._t)
        p.setPen(Qt.NoPen)
        p.setBrush(halo)
        p.drawEllipse(QPointF(cx, cy), self.width() * 0.5, self.height() * 0.5)
        core = QColor(c)
        core.setAlphaF(0.65 + 0.35 * self._t)
        p.setBrush(core)
        p.drawEllipse(QPointF(cx, cy), self.width() * 0.22, self.height() * 0.22)
        p.end()


class FlashOverlay(QWidget):
    """Мягкая вспышка поверх виджета (например, таблицы) при обновлении данных.

    Раньше здесь был QWidget с QGraphicsOpacityEffect, но этот виджет лежит
    внутри страницы, у которой уже есть свой QGraphicsOpacityEffect (плавное
    переключение вкладок). Вложенные друг в друга графические эффекты Qt
    рисует некорректно: консоль засыпается предупреждениями QPainter, а на
    экране возможны артефакты перерисовки. Поэтому прозрачность рисуется
    вручную в paintEvent — эффект не нужен вовсе.
    """

    def __init__(self, parent=None, color=None, radius=10):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._color = color or ACCENT
        self._radius = radius
        self._alpha = 0.0
        self._anim = QVariantAnimation(self)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.valueChanged.connect(self._on_v)

    def flash(self, peak=0.28, ms=900):
        self._anim.stop()
        self._alpha = peak
        self.update()
        self._anim.setStartValue(peak)
        self._anim.setEndValue(0.0)
        self._anim.setDuration(ms)
        self._anim.start()

    def _on_v(self, v):
        self._alpha = float(v)
        self.update()

    def paintEvent(self, _e):
        if self._alpha <= 0.004:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = QColor(self._color)
        c.setAlphaF(min(1.0, max(0.0, self._alpha)))
        p.setPen(Qt.NoPen)
        p.setBrush(c)
        p.drawRoundedRect(QRectF(self.rect()), self._radius, self._radius)
        p.end()


class GradientLabel(QWidget):
    """Заголовок раздела, по которому непрерывно течёт светлая волна.

    Обычный QLabel так не умеет: QSS не заливает текст градиентом. Поэтому
    текст рисуется вручную пером-градиентом. При смене раздела волна
    запускается заново — получается «проявление» нового названия.
    """

    def __init__(self, text="", point_size=16.0, parent=None):
        super().__init__(parent)
        self._text = text
        self._font = _ui_font(point_size, QFont.ExtraBold)
        self._phase = -0.4
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._timer = QTimer(self)
        self._timer.setInterval(45)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    def setText(self, text):
        if text == self._text:
            return
        self._text = text
        self._phase = -0.4
        self.updateGeometry()
        self.update()

    def text(self):
        return self._text

    def sizeHint(self):
        fm = QFontMetrics(self._font)
        return QSize(fm.horizontalAdvance(self._text) + 6, fm.height() + 4)

    def minimumSizeHint(self):
        return self.sizeHint()

    def _tick(self):
        self._phase += 0.007
        if self._phase > 1.4:
            self._phase = -0.4
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        p.setFont(self._font)

        def cl(v):
            return min(1.0, max(0.0, v))

        base = QColor(FG)
        hi = QColor(ACCENT if IS_LIGHT else ACCENT_HV)
        # Градиент строим по ширине САМОГО ТЕКСТА, а не всего виджета: иначе
        # на широком заголовке волна пробегала бы мимо коротких названий.
        text_w = max(1, QFontMetrics(self._font).horizontalAdvance(self._text))
        g = QLinearGradient(0, 0, text_w, 0)
        g.setColorAt(0.0, base)
        g.setColorAt(cl(self._phase - 0.22), base)
        g.setColorAt(cl(self._phase), hi)
        g.setColorAt(cl(self._phase + 0.22), base)
        g.setColorAt(1.0, base)
        p.setPen(QPen(QBrush(g), 1))
        p.drawText(self.rect(), Qt.AlignLeft | Qt.AlignVCenter, self._text)
        p.end()


class Ripple(QWidget):
    """Круг, расходящийся от точки нажатия на кнопку.

    Живёт доли секунды и удаляет себя сам. Мышь не перехватывает, в раскладку
    не попадает — на поведение кнопки никак не влияет.
    """

    def __init__(self, parent, center, color, radius):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setGeometry(parent.rect())
        self._center = QPointF(center)
        self._color = color
        self._radius = max(12.0, float(radius))
        self._t = 0.0

        self._anim = QVariantAnimation(self)
        self._anim.setDuration(420)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.setStartValue(0.0)
        self._anim.setEndValue(1.0)
        self._anim.valueChanged.connect(self._on_v)
        self._anim.finished.connect(self.deleteLater)
        self.show()
        self.raise_()
        self._anim.start()

    def _on_v(self, v):
        self._t = float(v)
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        # Обрезаем по скруглённому прямоугольнику кнопки, иначе круг вылезал
        # бы за её углы.
        clip = QPainterPath()
        clip.addRoundedRect(QRectF(self.rect()), STYLE["radius_md"], STYLE["radius_md"])
        p.setClipPath(clip)
        col = QColor(self._color)
        col.setAlphaF(0.32 * (1.0 - self._t))
        p.setPen(Qt.NoPen)
        p.setBrush(col)
        rr = self._radius * (0.18 + 0.95 * self._t)
        p.drawEllipse(self._center, rr, rr)
        p.end()


class RippleFilter(QObject):
    """Вешает круг-отклик на нажатие любой кнопки."""

    def eventFilter(self, obj, event):
        if (event.type() == QEvent.MouseButtonPress
                and isinstance(obj, QAbstractButton)
                and obj.isEnabled()
                # У кнопок с графическим эффектом (пункты меню) свой набор
                # анимаций — вложенные эффекты Qt рисует некорректно.
                and obj.graphicsEffect() is None):
            name = obj.objectName()
            color = ON_ACCENT if name in ("accent", "green") else (
                ACCENT if IS_LIGHT else ACCENT_HV)
            try:
                pos = event.position().toPoint()
            except Exception:
                pos = obj.rect().center()
            Ripple(obj, pos, color, max(obj.width(), obj.height()))
        return False


class SweepLine(QWidget):
    """Тонкая линия под заголовком раздела. В покое — почти незаметная
    полоска, при смене вкладки по ней пробегает янтарный блик. Именно такие
    мелочи и отличают «дорогой» интерфейс от просто тёмного."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(2)
        self._phase = None
        self._flow = -0.4          # медленное фоновое течение
        self._flow_timer = QTimer(self)
        self._flow_timer.setInterval(50)
        self._flow_timer.timeout.connect(self._flow_tick)
        self._flow_timer.start()
        self._anim = QVariantAnimation(self)
        self._anim.setStartValue(-0.35)
        self._anim.setEndValue(1.35)
        self._anim.setDuration(620)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.valueChanged.connect(self._on_v)
        self._anim.finished.connect(self._on_done)

    def sweep(self):
        self._anim.stop()
        self._anim.start()

    def _on_v(self, v):
        self._phase = float(v)
        self.update()

    def _on_done(self):
        self._phase = None
        self.update()

    def _flow_tick(self):
        self._flow += 0.005
        if self._flow > 1.4:
            self._flow = -0.4
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        base = QColor(FG)
        base.setAlpha(30 if IS_LIGHT else 16)
        p.fillRect(0, 0, w, h, base)

        def cl(v):
            return min(1.0, max(0.0, v))

        # постоянная мягкая волна — линия «дышит», а не просто лежит
        if w > 0:
            soft = QLinearGradient(0, 0, w, 0)
            c_soft = QColor(ACCENT_HV)
            c_soft.setAlpha(0)
            c_mid = QColor(ACCENT_HV)
            c_mid.setAlpha(90 if not IS_LIGHT else 120)
            soft.setColorAt(0.0, c_soft)
            soft.setColorAt(cl(self._flow - 0.25), c_soft)
            soft.setColorAt(cl(self._flow), c_mid)
            soft.setColorAt(cl(self._flow + 0.25), c_soft)
            soft.setColorAt(1.0, c_soft)
            p.fillRect(0, 0, w, h, QBrush(soft))
        if self._phase is not None and w > 0:
            def _cl(v):
                # Qt ругается на позицию стопа вне 0..1, а фаза специально
                # выходит за края (блик «вылетает» за пределы линии) — поэтому
                # каждую позицию обрезаем перед передачей в градиент.
                return min(1.0, max(0.0, v))

            g = QLinearGradient(0, 0, w, 0)
            c = QColor(ACCENT)
            c0 = QColor(ACCENT)
            c0.setAlpha(0)
            mid = _cl(self._phase)
            left = _cl(self._phase - 0.18)
            right = _cl(self._phase + 0.18)
            g.setColorAt(0.0, c0)
            g.setColorAt(left, c0)
            g.setColorAt(mid, c)
            g.setColorAt(right, c0)
            g.setColorAt(1.0, c0)
            p.fillRect(0, 0, w, h, QBrush(g))
        p.end()


class NavIndicator(QWidget):
    """Скользящий янтарный маркер слева от активного пункта меню.
    Не пересоздаётся при переключении — просто плавно переезжает
    на новое место (анимация геометрии)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._glow = 0.55
        up = QVariantAnimation(self)
        up.setStartValue(0.35)
        up.setEndValue(1.0)
        up.setDuration(1500)
        up.setEasingCurve(QEasingCurve.InOutSine)
        down = QVariantAnimation(self)
        down.setStartValue(1.0)
        down.setEndValue(0.35)
        down.setDuration(1900)
        down.setEasingCurve(QEasingCurve.InOutSine)
        for a in (up, down):
            a.valueChanged.connect(self._on_v)
        self._pulse = QSequentialAnimationGroup(self)
        self._pulse.addAnimation(up)
        self._pulse.addAnimation(down)
        self._pulse.setLoopCount(-1)
        self._pulse.start()

    def _on_v(self, v):
        self._glow = float(v)
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        # мягкий ореол
        halo = QColor(ACCENT_HV)
        halo.setAlphaF(0.10 + 0.18 * self._glow)
        p.setPen(Qt.NoPen)
        p.setBrush(halo)
        p.drawRoundedRect(QRectF(0, h * 0.12, w, h * 0.76), w / 2.0, w / 2.0)
        # сам маркер
        bar_w = 3.0
        # Сам маркер «переливается»: светлая лазурь и тёмная синь медленно
        # меняются местами по вертикали в такт пульсации ореола.
        g = QLinearGradient(0, 0, 0, h)
        t = self._glow
        g.setColorAt(0.0, QColor(_mix(ACCENT, ACCENT_HV, t)))
        g.setColorAt(0.5, QColor(_mix(ACCENT_HV, ACCENT, t * 0.6)))
        g.setColorAt(1.0, QColor(_mix(ACCENT_DEEP, ACCENT, t)))
        p.setBrush(QBrush(g))
        p.drawRoundedRect(QRectF((w - bar_w) / 2.0, 0, bar_w, h), bar_w / 2.0, bar_w / 2.0)
        p.end()


# ══════════════════════════════════════════════════════════════════════════════
#                     ВСПЛЫВАЮЩИЕ УВЕДОМЛЕНИЯ (ТОСТЫ)
# ══════════════════════════════════════════════════════════════════════════════
class Toast(QFrame):
    """Карточка-уведомление в правом нижнем углу.

    Заменяет модальные окна «Готово!» там, где не нужен ответ пользователя:
    сообщение показывается поверх интерфейса, само уезжает через несколько
    секунд и не перехватывает управление. Клик по карточке закрывает её сразу.
    """

    closed = Signal()
    WIDTH = 360

    def __init__(self, parent, title, text="", kind="info", ms=4200):
        super().__init__(parent)
        color, icon_role = {
            "ok":   (GREEN, "check"),
            "err":  (DUPE_RED, "warning"),
            "info": (ACCENT if IS_LIGHT else ACCENT_HV, "info"),
        }.get(kind, (ACCENT, "info"))

        self.setObjectName("toast")
        self.setFixedWidth(self.WIDTH)
        self.setCursor(Qt.PointingHandCursor)
        # Селектор с именем обязателен: иначе рамка и скругление достались бы
        # ещё и всем подписям внутри карточки.
        self.setStyleSheet(
            f"QFrame#toast {{ background-color:{'#ffffff' if IS_LIGHT else _mix(BG3, '#ffffff', 0.07)};"
            f" border:1px solid {_rgba(color, 0.5)}; border-radius:14px; }}")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 12, 16, 12)
        lay.setSpacing(12)

        bar = QFrame()
        bar.setFixedWidth(3)
        bar.setStyleSheet(f"background-color:{color}; border-radius:2px;")
        lay.addWidget(bar)

        ic = QLabel()
        qic = _icon(ICONS.get(icon_role, icon_role), color, None)
        if isinstance(qic, QIcon):
            ic.setPixmap(qic.pixmap(20, 20))
        ic.setFixedWidth(22)
        lay.addWidget(ic, 0, Qt.AlignTop)

        col = QVBoxLayout()
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(3)
        t = QLabel(title)
        t.setStyleSheet(f"color:{FG}; font-weight:800; font-size:10.5pt; background:transparent;")
        t.setWordWrap(True)
        col.addWidget(t)
        if text:
            sub = QLabel(text)
            sub.setStyleSheet(f"color:{FG2}; font-size:9.5pt; background:transparent;")
            sub.setWordWrap(True)
            col.addWidget(sub)
        lay.addLayout(col, 1)

        self._eff = QGraphicsOpacityEffect(self)
        self._eff.setOpacity(0.0)
        self.setGraphicsEffect(self._eff)
        self._group = None
        self._life = QTimer(self)
        self._life.setSingleShot(True)
        self._life.timeout.connect(self.dismiss)
        self._ms = ms
        self._dying = False

        # Тонкая полоска внизу карточки показывает, сколько ей осталось жить
        self._life_bar = QFrame(self)
        self._life_bar.setStyleSheet(
            f"background-color:{_rgba(color, 0.6)}; border:none; border-radius:1px;")
        self.adjustSize()

    # ── появление: одновременно всплывает и проявляется ──
    def appear(self, target_pos):
        self.move(target_pos + QPoint(0, 18))
        self.show()
        self.raise_()
        slide = QPropertyAnimation(self, b"pos", self)
        slide.setDuration(280)
        slide.setEasingCurve(QEasingCurve.OutCubic)
        slide.setStartValue(self.pos())
        slide.setEndValue(target_pos)
        fade = QPropertyAnimation(self._eff, b"opacity", self)
        fade.setDuration(240)
        fade.setStartValue(0.0)
        fade.setEndValue(1.0)
        self._group = QParallelAnimationGroup(self)
        self._group.addAnimation(slide)
        self._group.addAnimation(fade)
        self._group.start()
        self._life.start(self._ms)

        bar_w = max(10, self.width() - 28)
        bar_y = self.height() - 7
        self._life_bar.setGeometry(14, bar_y, bar_w, 2)
        shrink = QPropertyAnimation(self._life_bar, b"geometry", self)
        shrink.setDuration(self._ms)
        shrink.setEasingCurve(QEasingCurve.Linear)
        shrink.setStartValue(QRect(14, bar_y, bar_w, 2))
        shrink.setEndValue(QRect(14, bar_y, 0, 2))
        shrink.start()
        self._life_anim = shrink

    def slide_to(self, pos):
        """Плавно переезжает на новое место, когда соседние тосты исчезают."""
        if self._dying:
            return
        anim = QPropertyAnimation(self, b"pos", self)
        anim.setDuration(220)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.setStartValue(self.pos())
        anim.setEndValue(pos)
        anim.start()
        self._move_anim = anim

    def dismiss(self):
        if self._dying:
            return
        self._dying = True
        self._life.stop()
        fade = QPropertyAnimation(self._eff, b"opacity", self)
        fade.setDuration(220)
        fade.setStartValue(self._eff.opacity())
        fade.setEndValue(0.0)
        fade.finished.connect(self._finish)
        fade.start()
        self._fade_out = fade

    def _finish(self):
        self.closed.emit()
        self.hide()
        self.deleteLater()

    def mousePressEvent(self, e):
        self.dismiss()
        super().mousePressEvent(e)


# ══════════════════════════════════════════════════════════════════════════════
#                    ЗАГЛУШКА ДЛЯ ПУСТОЙ ТАБЛИЦЫ
# ══════════════════════════════════════════════════════════════════════════════
class EmptyState(QWidget):
    """Подсказка по центру пустой таблицы вместо голой сетки.

    Живёт внутри viewport таблицы, сама подстраивается под его размер и
    показывается ровно тогда, когда строк нет. Мышь не перехватывает.
    """

    def __init__(self, parent, icon_role, title, hint):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._table = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)
        lay.addStretch()

        ic = QLabel()
        ic.setAlignment(Qt.AlignCenter)
        qic = _icon(ICONS.get(icon_role, icon_role), _mix(FG3, BG3, 0.35), None)
        if isinstance(qic, QIcon):
            ic.setPixmap(qic.pixmap(46, 46))
        lay.addWidget(ic)

        t = QLabel(title)
        t.setAlignment(Qt.AlignCenter)
        t.setStyleSheet(f"color:{FG2}; font-size:11pt; font-weight:700; background:transparent;")
        lay.addWidget(t)

        h = QLabel(hint)
        h.setAlignment(Qt.AlignCenter)
        h.setWordWrap(True)
        h.setStyleSheet(f"color:{FG3}; font-size:9.5pt; background:transparent;")
        lay.addWidget(h)
        lay.addStretch()

    @classmethod
    def attach(cls, table, icon_role, title, hint):
        es = cls(table.viewport(), icon_role, title, hint)
        es._table = table
        table._empty_state = es          # держим ссылку, иначе соберёт сборщик
        model = table.model()
        for sig in (model.rowsInserted, model.rowsRemoved,
                    model.modelReset, model.layoutChanged):
            sig.connect(lambda *a, w=es: w.sync())
        table.viewport().installEventFilter(es)
        es.sync()
        return es

    def sync(self):
        if self._table is None:
            return
        try:
            vp = self._table.viewport()
            self.resize(vp.size())
            self.setVisible(self._table.rowCount() == 0)
            if self.isVisible():
                self.raise_()
        except RuntimeError:
            # Таблицу уже удалил Qt (закрытие программы), а сигнал модели
            # долетел следом — просто игнорируем.
            self._table = None

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Resize:
            self.resize(obj.size())
        return False


# ══════════════════════════════════════════════════════════════════════════════
#                         КОМАНДНАЯ ПАЛИТРА  (Ctrl+K)
# ══════════════════════════════════════════════════════════════════════════════
class CommandPalette(QDialog):
    """Быстрый переход по разделам и действиям: Ctrl+K, набрать пару букв, Enter."""

    def __init__(self, parent, commands):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setModal(True)
        self.setFixedSize(560, 430)
        self._commands = list(commands)
        self._chosen = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        card = QFrame()
        card.setObjectName("palette_card")
        card.setStyleSheet(
            f"QFrame#palette_card {{ background-color:{'#ffffff' if IS_LIGHT else _mix(BG3, '#ffffff', 0.05)};"
            f" border:1px solid {_rgba(ACCENT, 0.45)}; border-radius:16px; }}")
        outer.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(10)

        row = QHBoxLayout()
        row.setSpacing(10)
        ic = QLabel()
        qic = _icon(ICONS.get("search", ""), ACCENT if IS_LIGHT else ACCENT_HV, None)
        if isinstance(qic, QIcon):
            ic.setPixmap(qic.pixmap(18, 18))
        row.addWidget(ic)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Раздел или действие…")
        self.search.setStyleSheet("border:none; background:transparent; font-size:13pt; padding:4px 0;")
        row.addWidget(self.search, 1)
        hint = QLabel("ESC")
        hint.setObjectName("eyebrow")
        row.addWidget(hint)
        lay.addLayout(row)

        line = QFrame()
        line.setFixedHeight(1)
        line.setStyleSheet(f"background-color:{_veil(0.09)}; border:none;")
        lay.addWidget(line)

        self.list = QListWidget()
        self.list.setStyleSheet(
            "QListWidget { background:transparent; border:none; outline:none; }"
            f"QListWidget::item {{ color:{FG}; padding:9px 12px; border-radius:9px; }}"
            f"QListWidget::item:selected {{ background-color:{_rgba(ACCENT, 0.20)}; color:{FG}; }}")
        self.list.setIconSize(QSize(17, 17))
        lay.addWidget(self.list, 1)

        self.search.textChanged.connect(self._refill)
        self.list.itemActivated.connect(lambda _i: self._run())
        self.search.installEventFilter(self)
        self._refill("")

    def _refill(self, text=""):
        q = (text or "").strip().lower()
        self.list.clear()
        for title, subtitle, icon_role, fn in self._commands:
            if q and q not in title.lower() and q not in (subtitle or "").lower():
                continue
            it = QListWidgetItem(f"{title}      {subtitle}" if subtitle else title)
            qic = _icon(ICONS.get(icon_role, icon_role), FG2, None)
            if isinstance(qic, QIcon):
                it.setIcon(qic)
            it.setData(Qt.UserRole, fn)
            self.list.addItem(it)
        if self.list.count():
            self.list.setCurrentRow(0)

    def eventFilter(self, obj, event):
        # Стрелки и Enter работают, пока курсор остаётся в строке поиска —
        # так не приходится переключаться между полем и списком.
        if obj is self.search and event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Down, Qt.Key_Up):
                row = self.list.currentRow() + (1 if key == Qt.Key_Down else -1)
                if 0 <= row < self.list.count():
                    self.list.setCurrentRow(row)
                return True
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self._run()
                return True
        return False

    def _run(self):
        it = self.list.currentItem()
        if it is None:
            return
        self._chosen = it.data(Qt.UserRole)
        self.accept()

    @property
    def chosen(self):
        return self._chosen


# ══════════════════════════════════════════════════════════════════════════════
#                       СВОДКА СЕССИИ  (клик по знаку программы)
# ══════════════════════════════════════════════════════════════════════════════
class ThemeTile(QWidget):
    """Плитка-превью темы: миниатюра окна, нарисованная ЕЁ собственными
    цветами, а не цветами текущей темы. Так видно, как программа будет
    выглядеть, ещё до перезапуска."""

    clicked = Signal(str)
    W, H = 178, 134

    def __init__(self, key, data, is_current=False, parent=None):
        super().__init__(parent)
        self.setFixedSize(self.W, self.H)
        self.setCursor(Qt.PointingHandCursor)
        self.setAttribute(Qt.WA_Hover, True)
        self._key = key
        self._d = data
        self._cur = is_current
        self._hover = False
        self.setToolTip(data.get("label", key))

    def enterEvent(self, e):
        self._hover = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hover = False
        self.update()
        super().leaveEvent(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton and self.rect().contains(e.pos()):
            self.clicked.emit(self._key)
        super().mouseReleaseEvent(e)

    def paintEvent(self, _e):
        d = self._d
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w = self.W
        ph = 88                       # высота миниатюры

        prev = QRectF(1, 1, w - 2, ph)
        path = QPainterPath()
        path.addRoundedRect(prev, 12, 12)
        p.save()
        p.setClipPath(path)

        p.fillRect(prev, QColor(d["BG"]))
        # боковое меню
        p.fillRect(QRectF(1, 1, 46, ph), QColor(d["NAV_BG"]))
        # пункты меню, первый — активный
        for i in range(3):
            y = 16 + i * 16
            col = QColor(d["ACCENT"]) if i == 0 else QColor(d["FG3"])
            col.setAlphaF(0.95 if i == 0 else 0.5)
            p.setPen(Qt.NoPen)
            p.setBrush(col)
            p.drawRoundedRect(QRectF(9, y, 28 - i * 4, 5), 2.5, 2.5)
        # карточка контента
        p.setBrush(QColor(d["BG3"]))
        p.drawRoundedRect(QRectF(56, 12, w - 70, ph - 24), 8, 8)
        # строки текста
        for i, bw in enumerate((72, 54, 62)):
            col = QColor(d["FG2"])
            col.setAlphaF(0.55 - i * 0.12)
            p.setBrush(col)
            p.drawRoundedRect(QRectF(66, 22 + i * 11, bw, 4), 2, 2)
        # акцентная кнопка с градиентом
        g = QLinearGradient(66, 0, 66 + 46, 0)
        g.setColorAt(0.0, QColor(d["ACCENT_HV"]))
        g.setColorAt(1.0, QColor(d["ACCENT_DEEP"]))
        p.setBrush(QBrush(g))
        p.drawRoundedRect(QRectF(66, ph - 30, 46, 13), 6, 6)
        p.restore()

        # рамка миниатюры
        edge = QColor(d["ACCENT"]) if (self._hover or self._cur) else QColor(d["FG3"])
        edge.setAlphaF(1.0 if self._hover else (0.8 if self._cur else 0.35))
        pen = QPen(edge)
        pen.setWidthF(2.0 if (self._hover or self._cur) else 1.0)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawPath(path)

        # подпись + отметка текущей темы
        label = d.get("label", self._key)
        p.setFont(_ui_font(9.5, QFont.DemiBold))
        p.setPen(QColor(FG if (self._hover or self._cur) else FG2))
        p.drawText(QRectF(0, ph + 6, w, 22), Qt.AlignCenter, label)
        if self._cur:
            dot = QColor(d["ACCENT"])
            p.setPen(Qt.NoPen)
            p.setBrush(dot)
            p.drawEllipse(QPointF(w - 14, 14), 5, 5)
        p.end()


class ThemePicker(QDialog):
    """Окно выбора оформления: живые миниатюры всех тем."""

    def __init__(self, parent, current):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setModal(True)
        self.chosen = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        card = QFrame()
        card.setObjectName("theme_card")
        card.setStyleSheet(
            f"QFrame#theme_card {{ background-color:{'#ffffff' if IS_LIGHT else _mix(BG3, '#ffffff', 0.05)};"
            f" border:1px solid {_rgba(ACCENT, 0.35)}; border-radius:18px; }}")
        outer.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(4)

        ttl = QLabel("Оформление")
        ttl.setStyleSheet(f"color:{FG}; font-size:14pt; font-weight:800; background:transparent;")
        lay.addWidget(ttl)
        sub = QLabel("тема применится после перезапуска программы")
        sub.setObjectName("muted")
        lay.addWidget(sub)
        lay.addSpacing(16)

        grid = QGridLayout()
        grid.setSpacing(14)
        for i, key in enumerate(C.THEME_ORDER):
            data = C.THEMES.get(key)
            if not data:
                continue
            tile = ThemeTile(key, data, is_current=(key == current))
            tile.clicked.connect(self._pick)
            grid.addWidget(tile, i // 3, i % 3)
        lay.addLayout(grid)

        lay.addSpacing(16)
        row = QHBoxLayout()
        row.addStretch()
        close = QPushButton("Отмена")
        close.setCursor(Qt.PointingHandCursor)
        close.clicked.connect(self.reject)
        row.addWidget(close)
        lay.addLayout(row)

    def _pick(self, key):
        self.chosen = key
        self.accept()


class AboutCard(QDialog):
    """Компактная карточка «о программе»: версия, дата сборки и что нового.

    Открывается из командной палитры (Ctrl+K → «О программе»).
    """

    def __init__(self, parent, version_line, changes=()):
        super().__init__(parent)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setModal(True)
        self.setFixedWidth(460)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        card = QFrame()
        card.setObjectName("about_card")
        card.setStyleSheet(
            f"QFrame#about_card {{ background-color:{'#ffffff' if IS_LIGHT else _mix(BG3, '#ffffff', 0.05)};"
            f" border:1px solid {_rgba(ACCENT, 0.35)}; border-radius:18px; }}")
        outer.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(4)

        head = QHBoxLayout()
        head.setSpacing(14)
        head.addWidget(_DbBadge(size=46))
        col = QVBoxLayout()
        col.setSpacing(1)
        ttl = QLabel("BASE")
        ttl.setStyleSheet(f"color:{FG}; font-size:15pt; font-weight:800;"
                          " letter-spacing:3px; background:transparent;")
        col.addWidget(ttl)
        sub = QLabel("очистка и аналитика телефонных баз")
        sub.setObjectName("muted")
        col.addWidget(sub)
        head.addLayout(col)
        head.addStretch()
        lay.addLayout(head)
        lay.addSpacing(16)

        ver = QLabel(version_line)
        ver.setStyleSheet(f"color:{ACCENT if IS_LIGHT else ACCENT_SOFT}; font-size:11pt;"
                          " font-weight:700; background:transparent;")
        lay.addWidget(ver)

        theme_name = C.THEMES.get(getattr(C, "_active_theme_name", C.DEFAULT_THEME), {}).get("label", "Vortex")
        th = QLabel(f"Тема: {theme_name}  ·  переключается в шапке окна")
        th.setObjectName("muted")
        lay.addWidget(th)

        if changes:
            lay.addSpacing(14)
            line = QFrame()
            line.setFixedHeight(1)
            line.setStyleSheet(f"background-color:{_veil(0.09)}; border:none;")
            lay.addWidget(line)
            lay.addSpacing(10)
            cap = QLabel("ЧТО ИЗМЕНИЛОСЬ")
            cap.setObjectName("eyebrow")
            lay.addWidget(cap)
            lay.addSpacing(4)
            ch = QLabel("<br>".join(f"·&nbsp;{c}" for c in changes))
            ch.setTextFormat(Qt.RichText)
            ch.setWordWrap(True)
            ch.setStyleSheet(f"color:{FG2}; font-size:9.5pt; background:transparent;")
            lay.addWidget(ch)

        lay.addSpacing(16)
        row = QHBoxLayout()
        row.addStretch()
        ok = QPushButton("Закрыть")
        ok.setObjectName("accent")
        ok.setCursor(Qt.PointingHandCursor)
        ok.clicked.connect(self.accept)
        row.addWidget(ok)
        lay.addLayout(row)


# ──────────────────────────  ПУНКТ БОКОВОГО МЕНЮ  ──────────────────────────
class NavItemBtn(QPushButton):
    """Пункт бокового меню: иконка + название во всю ширину панели.

    При выборе иконка перекрашивается в акцент, текст становится ярче и
    плотнее, а сам пункт получает стеклянную янтарную заливку. Подсветка
    задаётся inline-стилем (а не через :checked в общем QSS) — inline-стиль
    в Qt имеет наивысший приоритет и гарантированно перекрывает капризы
    каскада, из-за которых :checked иногда не применялся.
    """

    def __init__(self, icon_name, text, parent=None, fallback="•"):
        super().__init__(parent)
        self._label_text = text
        self.setText(f"   {text}")
        self.setObjectName("navitem")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(46)

        self._ic_off = _icon(icon_name, FG2, fallback)
        self._ic_on = _icon(icon_name, ACCENT_HV, fallback)
        if isinstance(self._ic_off, QIcon):
            self.setIcon(self._ic_off)
            self.setIconSize(QSize(19, 19))
        else:
            # запасной вариант без qtawesome — символ прямо в тексте
            self.setText(f"  {self._ic_off}   {text}")

        self.setFont(_ui_font(11, QFont.DemiBold))

        # Плавное «проявление» неактивных пунктов при наведении
        self._eff = QGraphicsOpacityEffect(self)
        self._eff.setOpacity(0.68)
        self.setGraphicsEffect(self._eff)
        self._anim = QPropertyAnimation(self._eff, b"opacity", self)
        self._anim.setDuration(170)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        # QGraphicsOpacityEffect не всегда перерисовывается сам на каждом
        # кадре — форсируем repaint.
        self._anim.valueChanged.connect(lambda _v: self.update())
        self._anim.finished.connect(self.update)

        self.toggled.connect(self._on_toggle)

    # ── состояние ──
    def _on_toggle(self, checked):
        self._fade_to(1.0 if checked else 0.68)
        if isinstance(self._ic_off, QIcon):
            self.setIcon(self._ic_on if checked else self._ic_off)
            if checked:
                self._pop_icon()
        if checked:
            self.setStyleSheet(
                "QPushButton#navitem {"
                f"  background: qlineargradient(x1:0, y1:0, x2:1, y2:0,"
                f"      stop:0 {_rgba(ACCENT_HV, 0.24)}, stop:0.75 {_rgba(ACCENT, 0.06)}, stop:1 transparent);"
                f"  color: {FG};"
                f"  border: 1px solid {_rgba(ACCENT_HV, 0.28)};"
                f"  font-weight: 800;"
                "}")
        else:
            self.setStyleSheet("")

    def _pop_icon(self):
        """Короткий «подскок» иконки в момент выбора раздела."""
        a = QVariantAnimation(self)
        a.setDuration(340)
        a.setEasingCurve(QEasingCurve.OutCubic)
        a.setStartValue(19.0)
        a.setKeyValueAt(0.45, 24.0)
        a.setEndValue(19.0)
        a.valueChanged.connect(lambda v: self.setIconSize(QSize(int(v), int(v))))
        a.start()
        self._pop = a

    def _fade_to(self, target):
        self._anim.stop()
        self._anim.setStartValue(self._eff.opacity())
        self._anim.setEndValue(target)
        self._anim.start()

    def enterEvent(self, e):
        if not self.isChecked():
            self._fade_to(1.0)
        super().enterEvent(e)

    def leaveEvent(self, e):
        if not self.isChecked():
            self._fade_to(0.68)
        super().leaveEvent(e)


# ──────────────────────────  БОКОВОЕ МЕНЮ  ──────────────────────────
class Sidebar(QWidget):
    """Боковая панель: знак программы, список разделов со скользящим
    маркером и карточка состояния внизу."""

    _ICONS = {
        "Очистка": (ICONS["broom"], "◫"),
        "Города": (ICONS["map"], "◎"),
        "Редактирование": (ICONS["edit"], "✎"),
        "Генерация": (ICONS["bolt"], "⚡"),
        "Вклады": (ICONS["piggy"], "◈"),
    }

    def __init__(self, tab_names, on_tab, parent=None):
        super().__init__(parent)
        self.setFixedWidth(STYLE.get("sidebar_w", 262))
        self.setStyleSheet("background: transparent;")

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        panel = QFrame()
        panel.setObjectName("nav_panel")
        root.addWidget(panel, 1)

        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 18, 0, 14)
        lay.setSpacing(0)

        # ── знак + название ──
        head = QHBoxLayout()
        head.setContentsMargins(16, 0, 16, 0)
        head.setSpacing(12)
        self.badge = _DbBadge()
        head.addWidget(self.badge)

        titles = QVBoxLayout()
        titles.setContentsMargins(0, 2, 0, 0)
        titles.setSpacing(1)
        name_row = QHBoxLayout()
        name_row.setContentsMargins(0, 0, 0, 0)
        name_row.setSpacing(0)
        name = QLabel("BASE")
        name.setObjectName("wordmark")
        dot = QLabel(".")
        dot.setStyleSheet(f"color:{ACCENT}; font-size:12pt; font-weight:800; background:transparent;")
        # Мигающая точка с плавным затуханием — «живой» акцент в статичном углу
        dot_eff = QGraphicsOpacityEffect(dot)
        dot.setGraphicsEffect(dot_eff)
        d_up = QVariantAnimation(dot)
        d_up.setStartValue(0.15)
        d_up.setEndValue(1.0)
        d_up.setDuration(700)
        d_up.setEasingCurve(QEasingCurve.InOutSine)
        d_down = QVariantAnimation(dot)
        d_down.setStartValue(1.0)
        d_down.setEndValue(0.15)
        d_down.setDuration(1500)
        d_down.setEasingCurve(QEasingCurve.InOutSine)
        for a in (d_up, d_down):
            a.valueChanged.connect(lambda v, eff=dot_eff: eff.setOpacity(float(v)))
        self._title_dot_anim = QSequentialAnimationGroup(dot)
        self._title_dot_anim.addAnimation(d_up)
        self._title_dot_anim.addAnimation(d_down)
        self._title_dot_anim.setLoopCount(-1)
        self._title_dot_anim.start()
        name_row.addWidget(name)
        name_row.addWidget(dot)
        name_row.addStretch()
        titles.addLayout(name_row)
        sub = QLabel("ТЕЛЕФОННЫЕ БАЗЫ")
        sub.setObjectName("wordmark_sub")
        titles.addWidget(sub)
        head.addLayout(titles)
        head.addStretch()
        lay.addLayout(head)

        lay.addSpacing(18)
        sep = QFrame()
        sep.setObjectName("nav_sep")
        sep.setFixedHeight(1)
        sep_wrap = QHBoxLayout()
        sep_wrap.setContentsMargins(16, 0, 16, 0)
        sep_wrap.addWidget(sep)
        lay.addLayout(sep_wrap)
        lay.addSpacing(16)

        eyebrow = QLabel("РАЗДЕЛЫ")
        eyebrow.setObjectName("eyebrow")
        eb_wrap = QHBoxLayout()
        eb_wrap.setContentsMargins(20, 0, 16, 0)
        eb_wrap.addWidget(eyebrow)
        eb_wrap.addStretch()
        lay.addLayout(eb_wrap)
        lay.addSpacing(8)

        # ── список разделов ──
        self._host = QWidget()
        host_lay = QVBoxLayout(self._host)
        host_lay.setContentsMargins(0, 0, 0, 0)
        host_lay.setSpacing(2)

        self.tab_buttons = {}
        self.tab_group = QButtonGroup(self)
        self.tab_group.setExclusive(True)
        for nm in tab_names:
            qname, fallback = self._ICONS.get(nm, (ICONS["list"], "•"))
            b = NavItemBtn(qname, nm, fallback=fallback)
            b.clicked.connect(lambda _=False, n=nm: on_tab(n))
            self.tab_group.addButton(b)
            self.tab_buttons[nm] = b
            host_lay.addWidget(b)

        self._indicator = NavIndicator(self._host)
        self._indicator.setGeometry(QRect(3, 0, 7, 22))
        self._indicator.hide()
        lay.addWidget(self._host)

        lay.addStretch()

        # ── карточка состояния внизу ──
        card = QFrame()
        card.setObjectName("side_card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(12, 10, 12, 10)
        cl.setSpacing(4)
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)
        row.addWidget(PulseDot(GREEN, 12))
        st = QLabel("Ядро активно")
        st.setStyleSheet(f"color:{FG}; font-size:9.5pt; font-weight:700; background:transparent;")
        row.addWidget(st)
        row.addStretch()
        cl.addLayout(row)
        self.side_hint = QLabel(f"Vortex UI · версия {C.APP_VERSION}")
        self.side_hint.setObjectName("muted")
        cl.addWidget(self.side_hint)
        card_wrap = QHBoxLayout()
        card_wrap.setContentsMargins(14, 0, 14, 0)
        card_wrap.addWidget(card)
        lay.addLayout(card_wrap)

        # первый раздел активен сразу
        if tab_names:
            first = self.tab_buttons[tab_names[0]]
            first.setChecked(True)
            QTimer.singleShot(0, lambda: self._move_indicator(first, animate=False))

    # ── скользящий маркер ──
    def _move_indicator(self, btn, animate=True):
        if btn is None:
            return
        h = 22
        target = QRect(3, btn.y() + (btn.height() - h) // 2, 7, h)
        if not self._indicator.isVisible():
            self._indicator.setGeometry(target)
            self._indicator.show()
            self._indicator.raise_()
            return
        self._indicator.raise_()
        if not animate:
            self._indicator.setGeometry(target)
            return
        anim = QPropertyAnimation(self._indicator, b"geometry", self)
        anim.setDuration(280)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.setStartValue(self._indicator.geometry())
        anim.setEndValue(target)
        anim.start()
        self._ind_anim = anim   # держим ссылку, иначе анимация умрёт сборщиком

    def select(self, name):
        btn = self.tab_buttons.get(name)
        if btn is None:
            return
        btn.setChecked(True)
        self._move_indicator(btn)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        checked = self.tab_group.checkedButton()
        if checked is not None:
            QTimer.singleShot(0, lambda b=checked: self._move_indicator(b, animate=False))


# Псевдоним для обратной совместимости
Header = Sidebar


# ──────────────────────────  СВОЙ ЗАГОЛОВОК ОКНА  ──────────────────────────
class TitleBar(QFrame):
    """Полоса заголовка вместо системной: название программы слева,
    кнопки управления окном справа. Перетаскивание — через нативный
    startSystemMove (плавно, без «прилипания» курсора), с ручным
    запасным вариантом на случай, если платформа его не поддерживает."""

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.setObjectName("titlebar")
        self._win = win
        self._drag_pos = None
        self.setFixedHeight(STYLE.get("titlebar_h", 46))

        lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 0, 8, 0)
        lay.setSpacing(10)

        mark = QLabel("BASE")
        mark.setObjectName("wordmark")
        mark.setStyleSheet("font-size:11pt; letter-spacing:3px;")
        lay.addWidget(mark)

        lay.addStretch()

        # Кнопка «развернуть» есть только у растягиваемого окна: при
        # фиксированном размере разворачивать нечего, и лишняя кнопка,
        # которая ничего не делает, выглядела бы как недоделка.
        self._resizable = bool(STYLE.get("resizable", True))

        # Тем стало пять, перебирать их по кругу неудобно — кнопка открывает
        # окно выбора с миниатюрами.
        self.btn_theme = self._win_btn(ICONS["palette"], "◑", "Оформление")
        self.btn_theme.clicked.connect(self._win._switch_theme)
        lay.addWidget(self.btn_theme)

        self.btn_full = self._win_btn(ICONS["fullscreen"], "⛶",
                                      "Полноэкранный режим  ·  F11")
        self.btn_full.clicked.connect(self._win.toggle_fullscreen)
        lay.addWidget(self.btn_full)

        self.btn_min = self._win_btn(ICONS["win_min"], "—", "Свернуть")
        self.btn_min.clicked.connect(self._win.showMinimized)
        lay.addWidget(self.btn_min)

        self.btn_max = None
        if self._resizable:
            self.btn_max = self._win_btn(ICONS["win_max"], "□", "Развернуть")
            self.btn_max.clicked.connect(self.toggle_max)
            lay.addWidget(self.btn_max)

        self.btn_close = self._win_btn(ICONS["win_close"], "✕", "Закрыть", close=True)
        self.btn_close.clicked.connect(self._win.close)
        lay.addWidget(self.btn_close)

    def _win_btn(self, icon_name, fallback, tip, close=False):
        b = QPushButton()
        b.setObjectName("win_close" if close else "win_btn")
        ic = _icon(icon_name, FG2, None)
        if isinstance(ic, QIcon):
            b.setIcon(ic)
            b.setIconSize(QSize(15, 15))
        else:
            b.setText(fallback)
        b.setFixedSize(38, 30)
        b.setCursor(Qt.PointingHandCursor)
        b.setToolTip(tip)
        return b

    def toggle_max(self):
        if not getattr(self, "_resizable", True):
            return
        if self._win.isMaximized():
            self._win.showNormal()
        else:
            self._win.showMaximized()
        self.sync_max_icon()

    def sync_state_icons(self):
        """Обновляет иконку кнопки полноэкранного режима (развернуть/свернуть)."""
        full = self._win.isFullScreen()
        ic = _icon(ICONS["fullscreen_exit" if full else "fullscreen"], FG2, None)
        if isinstance(ic, QIcon):
            self.btn_full.setIcon(ic)
        self.btn_full.setToolTip(
            "Выйти из полноэкранного режима  ·  F11" if full
            else "Полноэкранный режим  ·  F11")

    def sync_max_icon(self):
        if self.btn_max is None:
            return
        role = "win_rest" if self._win.isMaximized() else "win_max"
        ic = _icon(ICONS[role], FG2, None)
        if isinstance(ic, QIcon):
            self.btn_max.setIcon(ic)
        self.btn_max.setToolTip("Восстановить" if self._win.isMaximized() else "Развернуть")

    # ── перетаскивание окна ──
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and not self._win.isFullScreen():
            handle = self._win.windowHandle()
            if handle is not None and hasattr(handle, "startSystemMove"):
                try:
                    if handle.startSystemMove():
                        return
                except Exception:
                    pass
            self._drag_pos = e.globalPosition().toPoint() - self._win.frameGeometry().topLeft()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self._drag_pos is not None and (e.buttons() & Qt.LeftButton):
            if self._win.isMaximized():
                self._win.showNormal()
                self.sync_max_icon()
            self._win.move(e.globalPosition().toPoint() - self._drag_pos)
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        # У окна фиксированного размера двойной клик по шапке разворачивает
        # программу на весь экран — привычный жест не пропадает впустую.
        if e.button() == Qt.LeftButton:
            if self._resizable:
                self.toggle_max()
            else:
                self._win.toggle_fullscreen()
        super().mouseDoubleClickEvent(e)


class ResizeFrame(QWidget):
    """Корневой виджет окна без системной рамки. По периметру оставлена
    полоска в несколько пикселей (за счёт отступов внутреннего layout) —
    именно за неё окно тянется мышью за любой край или угол."""

    MARGIN = 6

    def __init__(self, win):
        super().__init__(win)
        self.setObjectName("root_bg")
        self._win = win
        # При фиксированном размере окна тянуть за края нечего — тогда рамка
        # работает просто как тонкий контур окна и мышь не перехватывает.
        self._resizable = bool(STYLE.get("resizable", True))
        self._edges = Qt.Edges()
        self._press_edges = None
        self._start_geo = None
        self._start_pos = None
        self.setMouseTracking(True)

    def _edges_at(self, pos):
        edges = Qt.Edges()
        if not self._resizable:
            return edges
        m = self.MARGIN + 2
        r = self.rect()
        if pos.x() <= m:
            edges |= Qt.LeftEdge
        elif pos.x() >= r.width() - m:
            edges |= Qt.RightEdge
        if pos.y() <= m:
            edges |= Qt.TopEdge
        elif pos.y() >= r.height() - m:
            edges |= Qt.BottomEdge
        return edges

    @staticmethod
    def _cursor_for(edges):
        left = bool(edges & Qt.LeftEdge)
        right = bool(edges & Qt.RightEdge)
        top = bool(edges & Qt.TopEdge)
        bottom = bool(edges & Qt.BottomEdge)
        if (left and top) or (right and bottom):
            return Qt.SizeFDiagCursor
        if (right and top) or (left and bottom):
            return Qt.SizeBDiagCursor
        if left or right:
            return Qt.SizeHorCursor
        if top or bottom:
            return Qt.SizeVerCursor
        return Qt.ArrowCursor

    def mouseMoveEvent(self, e):
        if self._press_edges is None:
            if not self._win.isMaximized():
                self.setCursor(self._cursor_for(self._edges_at(e.position().toPoint())))
            super().mouseMoveEvent(e)
            return
        # запасной (ручной) режим изменения размера
        g = QRect(self._start_geo)
        d = e.globalPosition().toPoint() - self._start_pos
        mw = self._win.minimumWidth() or 400
        mh = self._win.minimumHeight() or 300
        if self._press_edges & Qt.LeftEdge:
            g.setLeft(min(g.left() + d.x(), g.right() - mw))
        if self._press_edges & Qt.RightEdge:
            g.setRight(max(g.right() + d.x(), g.left() + mw))
        if self._press_edges & Qt.TopEdge:
            g.setTop(min(g.top() + d.y(), g.bottom() - mh))
        if self._press_edges & Qt.BottomEdge:
            g.setBottom(max(g.bottom() + d.y(), g.top() + mh))
        self._win.setGeometry(g)
        super().mouseMoveEvent(e)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton or self._win.isMaximized():
            super().mousePressEvent(e)
            return
        edges = self._edges_at(e.position().toPoint())
        if not edges:
            super().mousePressEvent(e)
            return
        handle = self._win.windowHandle()
        if handle is not None and hasattr(handle, "startSystemResize"):
            try:
                if handle.startSystemResize(edges):
                    return
            except Exception:
                pass
        self._press_edges = edges
        self._start_geo = self._win.geometry()
        self._start_pos = e.globalPosition().toPoint()

    def mouseReleaseEvent(self, e):
        self._press_edges = None
        super().mouseReleaseEvent(e)


# ──────────────────────────  КОМПАКТНАЯ КНОПКА-ИКОНКА  ──────────────────────────
class IconBtn(QPushButton):
    """Квадратная кнопка-иконка (36×36) с подсветкой и плавным тултипом."""

    def __init__(self, icon, tooltip_text, variant="normal", parent=None, tip_side="above"):
        """
        icon       — строка (emoji/unicode) или QIcon
        variant    — 'normal' | 'danger' | 'green'
        tip_side   — 'above' | 'below' — предпочитаемая сторона показа тултипа
                     (используется, если с этой стороны достаточно места в окне)
        """
        super().__init__(parent)
        if isinstance(icon, str):
            self.setText(icon)
        else:
            self.setIcon(icon)
            self.setIconSize(QSize(18, 18))

        obj_map = {"normal": "icon_btn", "danger": "icon_btn_danger", "green": "icon_btn_green"}
        self.setObjectName(obj_map.get(variant, "icon_btn"))
        self.setToolTip(tooltip_text)
        self.setFixedSize(36, 36)
        self.setCursor(Qt.PointingHandCursor)
        self._tip_side = tip_side

    def enterEvent(self, e):
        BubbleTip.shared().show_for(self, prefer=self._tip_side)
        super().enterEvent(e)

    def leaveEvent(self, e):
        BubbleTip.shared().hide_tip()
        super().leaveEvent(e)

    def event(self, e):
        # Подавляем нативный Qt-тултип (второй, дублирующийся снизу) —
        # текст уже показывает наш плавный BubbleTip сверху.
        if e.type() == QEvent.ToolTip:
            return True
        return super().event(e)


class IconFolderBtn(IconBtn):
    """Иконка-кнопка выбора папки (тот же квадратный формат 36×36, что и в
    тулбаре «Редактирование»). Сама папка не выводится текстом на кнопке —
    только во всплывающей подсказке при наведении показывается полный путь."""

    def __init__(self, tooltip_placeholder="Выбрать папку", parent=None, tip_side="above"):
        super().__init__(_icon(ICONS["folder"], FG, "📁"), tooltip_placeholder, parent=parent, tip_side=tip_side)
        self._path = ""
        self._placeholder = tooltip_placeholder
        self.clicked.connect(self._browse)

    def _browse(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            self.set(path)

    def get(self):
        return self._path

    def set(self, v):
        self._path = v or ""
        self.setToolTip(self._path if self._path else self._placeholder)


# ══════════════════════════════════════════════════════════════════════════════
#                              ГЛАВНОЕ ОКНО
# ══════════════════════════════════════════════════════════════════════════════
class App(QMainWindow):
    # Сигналы для обновления UI из фоновых потоков (доставляются в главный поток)
    sigLog = Signal(str)
    sigPreview = Signal(str, str)
    sigPreviewStats = Signal(str, int, int, int)
    sigPreviewStatsHide = Signal()
    sigProgress = Signal(bool)
    sigInfo = Signal(str, str)
    sigError = Signal(str, str)
    sigCleanDone = Signal(str, str, str)  # title, text (info)
    sigUpdate = Signal(str, object)  # этап обновления, данные

    sigCityStatus = Signal(str, str)
    sigCityFill = Signal(object)
    sigCityCount = Signal(str)
    sigCityFlash = Signal()  # вспышка при смене таблицы городов

    sigGenStatus = Signal(str, str)
    sigGenBusy = Signal(bool)
    sigGenProgress = Signal(int, int)  # (сделано файлов, всего ожидается файлов)

    sigEditStatus = Signal(str, str)
    sigEditCity = Signal(list)
    sigEditReload = Signal(object)  # список кортежей (phone, trans)
    sigFlash = Signal(bool)      # True = дубли найдены (красный), False = чисто (зелёный)
    sigFlashBtn = Signal(bool)  # True = успех (зелёный), False = ошибка (красный)

    # Всплывающее уведомление и счётчики сессии. Оба сигнала нужны затем,
    # чтобы фоновые потоки НЕ трогали виджеты напрямую: поток шлёт сигнал,
    # а рисует уже главный поток (иначе Qt падает или молча теряет отрисовку).
    sigToast = Signal(str, str, str)     # заголовок, текст, вид (ok/err/info)

    TAB_CLEAN = "Очистка"
    TAB_CITY = "Города"
    TAB_EDIT = "Редактирование"
    TAB_GEN = "Генерация"
    TAB_DEPOSITS = "Вклады"

    # Пояснение под крупным заголовком раздела (шапка рабочей области)
    TAB_SUBTITLES = {
        TAB_CLEAN: "Удаление дублей и нарезка базы на готовые файлы",
        TAB_CITY: "Города, телефонные коды и диапазоны номеров",
        TAB_EDIT: "Ручная правка, слияние и выгрузка номеров",
        TAB_GEN: "Генерация номеров по выбранным диапазонам",
        TAB_DEPOSITS: "Сверка двух выгрузок и фильтрация результата",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self._resizable = bool(STYLE.get("resizable", True))

        # состояние
        self.cities = None
        self.city_index = None
        self.ph_ops = None
        self.ph_codes = None
        self.ph_ranges = None
        self._last_rows = []
        self._preview_running = False
        self._edit_undo = []
        self._edit_detected_city = ""
        self._edit_city_counts = {}
        self._gen_ranges_undo = []
        self._trigger_re = None

        # ── каркас окна ──
        # Системная рамка выключена: заголовок рисуем сами (TitleBar), а по
        # периметру корневого виджета остаётся полоска в несколько пикселей,
        # за которую окно тянется мышью (ResizeFrame). Если платформа вдруг
        # не даст поставить флаг — тихо откатываемся на обычное окно.
        self._frameless = bool(STYLE.get("frameless", True))
        if self._frameless:
            try:
                self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
            except Exception:
                self._frameless = False

        central = ResizeFrame(self) if self._frameless else QWidget()
        central.setObjectName("root_bg")
        self.setCentralWidget(central)
        self._nav_on_top = False

        root = QVBoxLayout(central)
        self._root_lay = root
        if not self._frameless:
            _m = 0
        elif self._resizable:
            _m = ResizeFrame.MARGIN      # полоска для захвата мышью
        else:
            _m = 1                       # просто волосяной контур окна
        root.setContentsMargins(_m, _m, _m, _m)
        root.setSpacing(0)

        self.titlebar = None
        if self._frameless:
            self.titlebar = TitleBar(self)
            root.addWidget(self.titlebar)

        # ── тело: боковое меню + рабочая область ──
        body = QWidget()
        body.setObjectName("content_area")
        body_lay = QHBoxLayout(body)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(0)
        root.addWidget(body, 1)

        tab_names = [self.TAB_CLEAN, self.TAB_CITY, self.TAB_EDIT, self.TAB_GEN, self.TAB_DEPOSITS]
        self.header = Sidebar(tab_names, self._switch_tab)
        self.header.badge.clicked.connect(self._open_palette)
        body_lay.addWidget(self.header)

        right_col = QWidget()
        right_col.setObjectName("content_area")
        right_lay = QVBoxLayout(right_col)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.setSpacing(0)

        # ── шапка раздела: крупный заголовок + пояснение + чип с датой ──
        topbar = QFrame()
        topbar.setObjectName("topbar")
        topbar.setFixedHeight(76)
        tb_lay = QHBoxLayout(topbar)
        tb_lay.setContentsMargins(26, 14, 24, 10)
        tb_lay.setSpacing(16)

        self._title_box = QWidget()
        tcol = QVBoxLayout(self._title_box)
        tcol.setContentsMargins(0, 0, 0, 0)
        tcol.setSpacing(2)
        # Заголовок рисуется вручную: по тексту непрерывно течёт светлая волна
        self.topbar_title = GradientLabel(self.TAB_CLEAN, 16.0)
        self.topbar_sub = QLabel(self.TAB_SUBTITLES.get(self.TAB_CLEAN, ""))
        self.topbar_sub.setObjectName("muted")
        tcol.addWidget(self.topbar_title)
        tcol.addWidget(self.topbar_sub)
        tb_lay.addWidget(self._title_box)
        tb_lay.addStretch()

        # Заголовок раздела мягко «проявляется» при переключении вкладки.
        self._title_eff = QGraphicsOpacityEffect(self._title_box)
        self._title_eff.setOpacity(1.0)
        self._title_box.setGraphicsEffect(self._title_eff)
        self._title_anim = QPropertyAnimation(self._title_eff, b"opacity", self)
        self._title_anim.setDuration(260)
        self._title_anim.setEasingCurve(QEasingCurve.OutCubic)

        self.topbar_chip = QLabel(datetime.date.today().strftime("%d.%m.%Y"))
        self.topbar_chip.setObjectName("chip")
        tb_lay.addWidget(self.topbar_chip, 0, Qt.AlignVCenter)

        right_lay.addWidget(topbar)

        # Тонкая линия под шапкой: при смене раздела по ней пробегает блик
        self.sweep = SweepLine()
        sweep_wrap = QWidget()
        sw = QHBoxLayout(sweep_wrap)
        sw.setContentsMargins(24, 0, 24, 0)
        sw.addWidget(self.sweep)
        right_lay.addWidget(sweep_wrap)

        self.stack = QStackedWidget()
        # Анимация fade при переключении вкладок
        self._tab_opacity_effect = QGraphicsOpacityEffect(self.stack)
        self._tab_opacity_effect.setOpacity(1.0)
        self.stack.setGraphicsEffect(self._tab_opacity_effect)
        self._tab_anim = QPropertyAnimation(self._tab_opacity_effect, b"opacity", self)
        self._tab_anim.setDuration(220)
        self._tab_anim.setEasingCurve(QEasingCurve.InOutQuad)
        self._pending_tab_idx = None
        self._tab_anim.finished.connect(self._on_tab_anim_finished)

        # Карточка-панель вокруг контента.
        # Важно: НЕ вешаем сюда QGraphicsDropShadowEffect — он оборачивает
        # self.stack, у которого уже есть свой QGraphicsOpacityEffect для
        # fade-перехода между вкладками. Вложенные друг в друга графические
        # эффекты Qt приводят к серьёзным артефактам перерисовки (дублирование
        # текста, «пропадание» кнопок при наведении) — поэтому оформляем
        # карточку только через QSS (скруглённые углы, рамка), без эффекта.
        card = QFrame()
        card.setObjectName("content_card")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(20, 18, 20, 20)
        card_lay.addWidget(self.stack)

        wrap = QWidget()
        wl = QVBoxLayout(wrap)
        wl.setContentsMargins(24, 14, 24, 16)
        wl.addWidget(card)
        right_lay.addWidget(wrap, 1)

        body_lay.addWidget(right_col, 1)

        # ── нижняя строка состояния ──
        status = QFrame()
        status.setObjectName("statusbar")
        status.setFixedHeight(30)
        sl = QHBoxLayout(status)
        sl.setContentsMargins(20, 0, 6, 0)
        sl.setSpacing(10)
        self.status_hint = QLabel("Готово к работе   ·   Ctrl+K — быстрый переход")
        self.status_hint.setObjectName("muted")
        sl.addWidget(self.status_hint)
        sl.addStretch()
        self.update_btn = QPushButton("  Обновить")
        self.update_btn.setIcon(_icon("mdi6.update|mdi.update|fa5s.sync-alt", FG2, ""))
        self.update_btn.setIconSize(QSize(13, 13))
        self.update_btn.setCursor(Qt.PointingHandCursor)
        self.update_btn.setToolTip("Проверить обновления на GitHub и установить")
        self.update_btn.setFixedHeight(22)
        self._update_btn_style(False)
        self.update_btn.clicked.connect(lambda: self._check_update(manual=True))
        sl.addWidget(self.update_btn)
        ver = QLabel(f"VORTEX  ·  {C.APP_VERSION_FULL.upper()}")
        ver.setObjectName("eyebrow")
        ver.setToolTip("Версия программы и время сборки")
        sl.addWidget(ver)
        if self._frameless and self._resizable:
            grip = QSizeGrip(status)
            grip.setFixedSize(16, 16)
            sl.addWidget(grip, 0, Qt.AlignBottom | Qt.AlignRight)
        root.addWidget(status)

        self._pages = {}
        self._pages[self.TAB_CLEAN] = self._build_clean_tab()
        self._pages[self.TAB_CITY] = self._build_city_tab()
        self._pages[self.TAB_EDIT] = self._build_edit_tab()
        self._pages[self.TAB_GEN] = self._build_gen_tab()
        self._pages[self.TAB_DEPOSITS] = self._build_deposits_tab()
        for nm in tab_names:
            self.stack.addWidget(self._pages[nm])

        # подключаем сигналы
        self.sigLog.connect(self._append_log)
        self.sigPreview.connect(lambda t, c: self._set_label(self.preview_label, t, c))
        self.sigPreviewStats.connect(self._update_preview_stats)
        self.sigPreviewStatsHide.connect(lambda: self.preview_stats.setVisible(False))
        self.sigProgress.connect(self._set_progress)
        self.sigInfo.connect(lambda t, m: QMessageBox.information(self, t, m))
        self.sigError.connect(lambda t, m: QMessageBox.critical(self, t, m))
        # Модальное окно «Готово!» заменено уведомлением: подтверждать нечего,
        # а полный отчёт и так остаётся в логе под кнопкой.
        self.sigCleanDone.connect(self._on_clean_done)
        self.sigUpdate.connect(self._on_update_event)
        self._update_busy = False
        self._update_info = None
        # тихая проверка обновлений через несколько секунд после запуска
        QTimer.singleShot(4000, lambda: self._check_update(manual=False))
        self.sigToast.connect(lambda t, m, k: self.toast(t, m, k))
        self.sigCityStatus.connect(lambda t, c: self._set_label(self.city_status, t, c))
        self.sigCityFill.connect(self._fill_city_table)
        self.sigCityCount.connect(lambda t: self._set_label(self.city_count, t, ACCENT))
        self.sigCityFlash.connect(self._flash_city_table)
        self.sigGenStatus.connect(lambda t, c: self._set_label(self.gen_status, t, c))
        self.sigGenBusy.connect(self._set_gen_busy)
        self.sigGenProgress.connect(self._set_gen_progress)
        self.sigEditStatus.connect(lambda t, c: self._set_label(self.edit_status, t, c))
        self.sigEditCity.connect(self._set_edit_city_badge)
        self.sigEditReload.connect(self._reload_edit_table)
        self.sigFlash.connect(self._flash_preview)
        self.sigFlashBtn.connect(self._flash_clean_btn)
        self._flash_timer = QTimer()
        self._flash_step = 0
        self._flash_btn_timer = QTimer()
        self._flash_btn_step = 0
        self._tab_phase = None

        self._toasts = []

        # ── подсказки вместо пустых таблиц ──
        EmptyState.attach(self.city_table, "map", "Здесь появятся города",
                          "Введите областной центр и нажмите «Найти» — "
                          "или ищите по специализации.")
        EmptyState.attach(self.edit_table, "edit", "Таблица пока пуста",
                          "Откройте файл с номерами кнопкой в панели сверху.")
        EmptyState.attach(self.dep_table, "piggy", "Нет данных для сверки",
                          "Загрузите оба файла — «ФИО» и выгрузку с суммами.")

        self._apply_window_size()

        # Командная палитра — быстрый переход по разделам и действиям
        sc_pal = QShortcut(QKeySequence("Ctrl+K"), self)
        sc_pal.activated.connect(self._open_palette)

        # F11 — полный экран, Esc — выход из него
        sc_full = QShortcut(QKeySequence("F11"), self)
        sc_full.activated.connect(self.toggle_fullscreen)
        sc_esc = QShortcut(QKeySequence(Qt.Key_Escape), self)
        sc_esc.activated.connect(self._exit_fullscreen)

    def _apply_window_size(self):
        """Задаёт окну постоянный размер и ставит его по центру экрана.

        Размер берётся из стилевых токенов (win_w/win_h) и подобран так, чтобы
        самая широкая вкладка помещалась целиком — при фиксированном окне
        вёрстка и шрифты уже никуда не «съезжают». Если экран меньше рабочего
        размера, окно ужимается под доступную область (но не ниже win_min_*),
        иначе на маленьком мониторе часть интерфейса просто оказалась бы за
        краем экрана без возможности растянуть окно обратно."""
        w = int(STYLE.get("win_w", 1280))
        h = int(STYLE.get("win_h", 760))
        min_w = int(STYLE.get("win_min_w", 1215))
        min_h = int(STYLE.get("win_min_h", 640))

        geo = None
        try:
            screen = QApplication.primaryScreen()
            if screen is not None:
                geo = screen.availableGeometry()
        except Exception:
            geo = None

        if geo is not None:
            w = max(min_w, min(w, geo.width() - 20))
            h = max(min_h, min(h, geo.height() - 20))

        if self._resizable:
            self.resize(w, h)
            self.setMinimumSize(min_w, min_h)
        else:
            self.setFixedSize(w, h)

        if geo is not None:
            self.move(geo.left() + (geo.width() - w) // 2,
                      geo.top() + max(0, (geo.height() - h) // 2))

    # ---- вспомогательное ----
    def _maybe_flip_layout(self, lay):
        """Раньше переставляла блоки вкладки местами для «чёрной» темы.
        Тема осталась одна (Vortex), порядок блоков фиксированный — метод
        сохранён пустым, чтобы не трогать вызовы во всех пяти вкладках."""
        return

    def _switch_tab(self, name):
        idx = list(self._pages.keys()).index(name)
        # Уже показываем нужную вкладку и нет активного перехода — ничего не делаем
        if idx == (self._pending_tab_idx if self._tab_phase else self.stack.currentIndex()):
            return
        self._pending_tab_idx = idx
        self.header.select(name)
        self._set_topbar(name)
        if hasattr(self, "sweep"):
            self.sweep.sweep()
        # Если анимация уже идёт — просто перенаправляем назначение
        if self._tab_phase == "out":
            return
        if self._tab_phase == "in":
            self._tab_anim.stop()
        # Fade-out
        self._tab_anim.setStartValue(self._tab_opacity_effect.opacity())
        self._tab_anim.setEndValue(0.0)
        self._tab_anim.setDuration(140)
        self._tab_anim.setEasingCurve(QEasingCurve.OutQuad)
        self._tab_phase = "out"
        self._tab_anim.start()

    def _set_topbar(self, name):
        """Меняет заголовок и подпись раздела с мягким проявлением."""
        if not hasattr(self, "topbar_title"):
            return
        self.topbar_title.setText(name)
        self.topbar_sub.setText(self.TAB_SUBTITLES.get(name, ""))
        self._title_anim.stop()
        self._title_anim.setStartValue(0.0)
        self._title_anim.setEndValue(1.0)
        self._title_anim.start()

    def _on_tab_anim_finished(self):
        if self._tab_phase == "out" and self._pending_tab_idx is not None:
            # Переключаем страницу и запускаем fade-in
            self.stack.setCurrentIndex(self._pending_tab_idx)
            self._tab_anim.setStartValue(0.0)
            self._tab_anim.setEndValue(1.0)
            self._tab_anim.setDuration(230)
            self._tab_anim.setEasingCurve(QEasingCurve.OutCubic)
            self._tab_phase = "in"
            self._tab_anim.start()
        else:
            self._tab_phase = None
            self._pending_tab_idx = None

    def changeEvent(self, e):
        """При разворачивании окна убираем рамку-«поле» по периметру
        (в развёрнутом виде тянуть окно за края не нужно) и переключаем
        иконку кнопки «развернуть/восстановить»."""
        if e.type() == QEvent.WindowStateChange and getattr(self, "_frameless", False):
            m = 0 if self.isMaximized() else ResizeFrame.MARGIN
            if hasattr(self, "_root_lay"):
                self._root_lay.setContentsMargins(m, m, m, m)
            if getattr(self, "titlebar", None) is not None:
                self.titlebar.sync_max_icon()
        super().changeEvent(e)

    def showEvent(self, e):
        """Первое появление окна — мягкое проявление вместо резкого «щелчка»."""
        super().showEvent(e)
        if not getattr(self, "_intro_done", False):
            self._intro_done = True
            self.setWindowOpacity(0.0)
            if self.windowOpacity() > 0.5:
                # Система не умеет прозрачные окна — тогда просто показываем
                # окно как есть, без анимации (и без потока предупреждений).
                self.setWindowOpacity(1.0)
                return
            a = QPropertyAnimation(self, b"windowOpacity", self)
            a.setDuration(380)
            a.setEasingCurve(QEasingCurve.OutCubic)
            a.setStartValue(0.0)
            a.setEndValue(1.0)
            a.start()
            self._intro_anim = a

    def toggle_fullscreen(self):
        """Включает/выключает полноэкранный режим.

        Важно: у окна с фиксированным размером сняты ограничения ДО вызова
        showFullScreen() — иначе оконный менеджер просто не может растянуть
        окно на весь экран, и кнопка выглядела бы нерабочей. При выходе
        размер фиксируется обратно."""
        if self.isFullScreen():
            self._exit_fullscreen()
            return
        self._unlock_size()
        self.showFullScreen()
        if getattr(self, "titlebar", None) is not None:
            self.titlebar.sync_state_icons()
        if hasattr(self, "status_hint"):
            self.status_hint.setText("Полноэкранный режим · F11 или Esc — выйти")

    def _exit_fullscreen(self):
        if not self.isFullScreen():
            return
        self.showNormal()
        self._apply_window_size()
        if getattr(self, "titlebar", None) is not None:
            self.titlebar.sync_state_icons()
        if hasattr(self, "status_hint"):
            self.status_hint.setText("Готово к работе   ·   Ctrl+K — быстрый переход")

    def _unlock_size(self):
        """Снимает фиксацию размера окна (нужно перед разворачиванием)."""
        self.setMinimumSize(0, 0)
        self.setMaximumSize(16777215, 16777215)

    # ═══════════════  ТЕМА ОФОРМЛЕНИЯ  ═══════════════
    def _switch_theme(self):
        """Открывает окно выбора оформления.

        Тема применяется при следующем запуске: цвета зашиты в стили всех
        виджетов на момент их создания, поэтому честнее перезапустить
        программу, чем пересобирать окно на лету и терять открытые данные."""
        cur = getattr(C, "_active_theme_name", C.DEFAULT_THEME)
        dlg = ThemePicker(self, cur)
        if dlg.exec() != QDialog.Accepted or not dlg.chosen:
            return
        nxt = dlg.chosen
        if nxt == cur:
            return
        if not C.save_theme(nxt):
            QMessageBox.warning(self, "Тема",
                                "Не удалось сохранить выбор темы: нет доступа к файлу theme.json.")
            return
        label = C.THEMES[nxt]["label"]
        resp = QMessageBox.question(
            self, "Смена темы",
            f"Тема «{label}» будет применена после перезапуска программы.\n\n"
            "Перезапустить сейчас?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if resp == QMessageBox.Yes:
            self._restart_app()

    def _restart_app(self):
        """Перезапускает программу (после смены темы)."""
        import subprocess, time
        try:
            if getattr(sys, "frozen", False):
                # Собранный .exe: sys.argv[0] — это уже сам exe, повторно
                # добавлять sys.executable перед ним не нужно.
                exe_path = os.path.abspath(sys.executable)
                work_dir = os.path.dirname(exe_path)
                cmd = [exe_path] + sys.argv[1:]
            else:
                # Берём АБСОЛЮТНЫЙ путь к скрипту и его собственную папку:
                # при запуске двойным кликом рабочая папка процесса может быть
                # совсем другой, и относительный путь просто не найдётся.
                script_path = os.path.abspath(sys.argv[0])
                work_dir = os.path.dirname(script_path)
                cmd = [sys.executable, script_path] + sys.argv[1:]
            proc = subprocess.Popen(cmd, cwd=work_dir,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            # Даём новой копии стартовать и проверяем, что она не рухнула сразу:
            # иначе старое окно закроется, новое не откроется, и будет похоже,
            # что «ничего не произошло».
            time.sleep(0.8)
            ret = proc.poll()
            if ret is not None:
                _, err = proc.communicate()
                err_text = err.decode("utf-8", errors="replace").strip() if err else "(нет текста ошибки)"
                QMessageBox.critical(
                    self, "Не удалось перезапустить",
                    f"Новая копия программы запустилась и сразу закрылась (код {ret}).\n\n"
                    f"Команда: {' '.join(cmd)}\nРабочая папка: {work_dir}\n\n{err_text}")
                return
        except Exception as e:
            QMessageBox.critical(
                self, "Ошибка перезапуска",
                f"Не удалось перезапустить программу автоматически:\n{e}\n\n"
                "Тема уже сохранена — закрой программу и запусти её заново вручную.")
            return
        # Жёсткий выход сразу после старта новой копии — без ожидания фоновых
        # потоков, которые иногда мешали процессу завершиться.
        os._exit(0)

    # ═══════════════  ВСПЛЫВАЮЩИЕ УВЕДОМЛЕНИЯ  ═══════════════
    def toast(self, title, text="", kind="info", ms=4200):
        """Показывает карточку-уведомление в правом нижнем углу."""
        host = self.centralWidget()
        if host is None:
            return None
        # Больше четырёх карточек одновременно — уже свалка: самую старую
        # закрываем сами.
        while len(self._toasts) >= 4:
            self._toasts[0].dismiss()
            self._toasts.pop(0)
        t = Toast(host, title, text, kind, ms)
        t.closed.connect(lambda w=t: self._drop_toast(w))
        self._toasts.append(t)
        self._layout_toasts(new=t)
        return t

    def _drop_toast(self, w):
        if w in self._toasts:
            self._toasts.remove(w)
        self._layout_toasts()

    def _layout_toasts(self, new=None):
        """Раскладывает карточки стопкой снизу вверх над строкой состояния."""
        host = self.centralWidget()
        if host is None:
            return
        x = host.width() - Toast.WIDTH - 26
        y = host.height() - 48
        for t in reversed(self._toasts):
            h = t.sizeHint().height()
            y -= h
            pos = QPoint(x, y)
            if t is new:
                t.appear(pos)
            elif t.isVisible():
                t.slide_to(pos)
            else:
                t.move(pos)
            y -= 10

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if getattr(self, "_toasts", None):
            self._layout_toasts()

    def _on_clean_done(self, title, text, _extra=""):
        self.toast(title, text, "ok", ms=9000)

    # ═══════════════  ОБНОВЛЕНИЕ С GITHUB  ═══════════════
    def _update_btn_style(self, highlight):
        col = ACCENT if highlight else FG2
        self.update_btn.setStyleSheet(
            f"QPushButton {{ background:{_rgba(ACCENT, 0.16) if highlight else 'transparent'};"
            f" color:{col}; border:1px solid {_rgba(ACCENT, 0.45) if highlight else 'transparent'};"
            f" border-radius:8px; padding:0 10px; font-size:8.5pt; font-weight:700; }}"
            f"QPushButton:hover {{ background:{_rgba(ACCENT, 0.22)}; color:{ACCENT}; }}"
            f"QPushButton:disabled {{ color:{FG2}; }}")

    def _check_update(self, manual=True):
        if U is None:
            if manual:
                QMessageBox.warning(self, "Обновление", "Модуль обновления (updater.py) не найден.")
            return
        if self._update_busy:
            return
        if not manual and not U.load_config().get("repo"):
            return  # репозиторий ещё не настроен — молча пропускаем
        self._update_busy = True
        if manual:
            self.update_btn.setEnabled(False)
            self.update_btn.setText("  Проверяю…")

        def work():
            try:
                newer, info = U.check(C.APP_VERSION)
                self.sigUpdate.emit("checked", (manual, newer, info))
            except Exception as e:
                self.sigUpdate.emit("check_error", (manual, str(e)))
        threading.Thread(target=work, daemon=True).start()

    def _on_update_event(self, kind, data):
        if kind == "checked":
            manual, newer, info = data
            self._update_busy = False
            self.update_btn.setEnabled(True)
            if not newer:
                self.update_btn.setText("  Обновить")
                self._update_btn_style(False)
                if manual:
                    self.toast("Обновлений нет", f"У тебя последняя версия ({C.APP_VERSION}).", "ok")
                return
            self._update_info = info
            self.update_btn.setText(f"  Обновить до {info['version']}")
            self._update_btn_style(True)
            if not manual:
                self.toast("Есть обновление", f"Доступна версия {info['version']} — "
                           "нажми «Обновить» внизу окна.", "info", ms=7000)
                return
            changes = "\n".join(f"•  {c}" for c in info.get("changes", []))
            text = (f"Доступна версия {info['version']} (сейчас {C.APP_VERSION}).\n\n"
                    + (f"Что нового:\n{changes}\n\n" if changes else "")
                    + "Установить и перезапустить программу?")
            if QMessageBox.question(self, "Обновление", text) != QMessageBox.Yes:
                return
            self._install_update(info)
        elif kind == "check_error":
            manual, msg = data
            self._update_busy = False
            self.update_btn.setEnabled(True)
            self.update_btn.setText("  Обновить")
            if manual:
                QMessageBox.warning(self, "Обновление", f"Не удалось проверить обновления.\n\n{msg}")
        elif kind == "progress":
            self.update_btn.setText(f"  {data}")
            self.status_hint.setText(data)
        elif kind == "installed":
            QMessageBox.information(self, "Обновление",
                                    f"Версия {data} установлена.\nПрограмма сейчас перезапустится.")
            try:
                U.restart()
            except Exception as e:
                QMessageBox.warning(self, "Обновление",
                                    f"Обновление установлено, но перезапустить не удалось ({e}).\n"
                                    "Закрой и открой программу вручную.")
                return
            QApplication.quit()
        elif kind == "install_error":
            self._update_busy = False
            self.update_btn.setEnabled(True)
            self.update_btn.setText("  Обновить")
            self.status_hint.setText("Готово к работе   ·   Ctrl+K — быстрый переход")
            QMessageBox.warning(self, "Обновление", f"Обновление не установлено.\n\n{data}")

    def _install_update(self, info):
        self._update_busy = True
        self.update_btn.setEnabled(False)

        def work():
            try:
                U.install(info, C.APP_VERSION,
                          progress=lambda t: self.sigUpdate.emit("progress", t))
                self.sigUpdate.emit("installed", info["version"])
            except Exception as e:
                self.sigUpdate.emit("install_error", str(e))
        threading.Thread(target=work, daemon=True).start()

    def _open_about(self):
        """Карточка «о программе» — версия, тема и список изменений."""
        AboutCard(self, C.APP_VERSION_FULL, getattr(C, "APP_CHANGES", ())).exec()

    # ═══════════════  КОМАНДНАЯ ПАЛИТРА (Ctrl+K)  ═══════════════
    def _open_palette(self):
        cmds = []
        for name in self._pages:
            icon_name = Sidebar._ICONS.get(name, (ICONS["list"], ""))[0]
            cmds.append((name, "раздел", icon_name, lambda n=name: self._switch_tab(n)))
        cmds += [
            ("Полный экран", "F11", "fullscreen", self.toggle_fullscreen),
            ("Оформление", "выбрать тему", "palette", self._switch_theme),
            ("О программе", "версия и что нового", "info", self._open_about),
            ("Обновить программу", "скачать новую версию с GitHub", "mdi6.update|mdi.update|fa5s.sync-alt",
             lambda: self._check_update(manual=True)),
            ("Триггер-слова", "редактор списка", "list", self._open_trigger_editor),
        ]
        dlg = CommandPalette(self, cmds)
        if dlg.exec() == QDialog.Accepted and dlg.chosen:
            # Действие запускаем после закрытия окна, иначе оно откроется
            # поверх ещё живущей палитры.
            QTimer.singleShot(0, dlg.chosen)

    def _set_edit_city_badge(self, ranked):
        """ranked — список (город, доля_0_100) от самого частого к редкому,
        УЖЕ включает все города (не только топ-3) — каждый становится своим
        маленьким чипом; при нехватке места по ширине FlowLayout сам переносит
        лишние чипы на следующую строку, так что видно все, хоть 20 городов."""
        self.edit_city_flow.clear()
        for city, pct in ranked:
            chip = QLabel(f"{city} {pct}%" if pct is not None else city)
            chip.setStyleSheet(
                f"color:{ACCENT_SOFT}; font-weight:700; font-size:9pt; "
                f"background-color:{_rgba(ACCENT, 0.12)}; border:1px solid {_rgba(ACCENT, 0.28)}; "
                f"border-radius:9px; padding:3px 9px;")
            self.edit_city_flow.addWidget(chip)
        self.edit_city_box.setVisible(bool(ranked))
        self.edit_city_box.updateGeometry()

    def _set_label(self, label, text, color):
        label.setText(text)
        label.setStyleSheet(f"color:{color};")

    def _section(self, text):
        lab = QLabel(text)
        lab.setStyleSheet(f"color:{ACCENT}; font-weight:bold; font-size:10pt;")
        return lab

    def _divider(self):
        line = QFrame()
        line.setFixedHeight(1)
        line.setStyleSheet(f"background-color:{DIVIDER};")
        return line

    def _ensure_data(self):
        if self.cities is None:
            self.cities, self.city_index = C.load_city_db()
        if self.ph_ranges is None:
            self.ph_ops, self.ph_codes, self.ph_ranges = C.load_phone_db()

    # ═══════════════════════════  ВКЛАДКА «ОЧИСТКА»  ═══════════════════════════
    def _build_clean_tab(self):
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(20, 8, 20, 10)
        lay.setSpacing(6)

        self.called_row   = FolderRow("Прозвоненные",     "#818cf8", icon_variant="dupes", allow_files=True,
                                       show_folder_btn=False)
        self.uncalled_row = FolderRow("База для очистки", "#fbbf24", icon_variant="upload")
        self.ready_row    = FolderRow("Результат",        GREEN,     icon_variant="check",
                                       show_folder_btn=False)
        # «Результат» больше не выбирается вручную — папка создаётся сама,
        # внутри «Базы для очистки», с именем «без дублей <сегодняшняя дата>».
        # Поле оставлено только для информации (куда лягут файлы) и потому
        # нередактируемое.
        self.ready_row.edit.setReadOnly(True)
        self.ready_row.edit.setToolTip(
            "Папка создаётся автоматически внутри «Базы для очистки» "
            "и называется «без дублей <сегодняшняя дата>».")
        self.called_row.changed.connect(self._trigger_preview)
        self.uncalled_row.changed.connect(self._trigger_preview)
        self.uncalled_row.changed.connect(self._update_ready_path)
        lay.addWidget(self.called_row)
        lay.addWidget(self.uncalled_row)
        lay.addWidget(self.ready_row)
        self._update_ready_path()

        self.pf = QFrame()
        self.pf.setObjectName("preview_panel")
        self.pf.setStyleSheet(_preview_panel_qss())
        pfl = QHBoxLayout(self.pf)
        self.preview_label = QLabel("")
        self.preview_label.setStyleSheet(f"color:{FG2}; font-style:italic;")
        self.preview_label.setWordWrap(True)
        pfl.addWidget(self.preview_label)

        def _chip(icon_role, color):
            ic_lbl = QLabel()
            qic = _icon(icon_role, color, None)
            if qic:
                ic_lbl.setPixmap(qic.pixmap(15, 15))
            ic_lbl.setFixedWidth(17)
            txt_lbl = QLabel("")
            txt_lbl.setStyleSheet(f"color:{color}; font-weight:600;")
            return ic_lbl, txt_lbl

        self.preview_stats = QWidget()
        stats_lay = QHBoxLayout(self.preview_stats)
        stats_lay.setContentsMargins(0, 0, 0, 0)
        stats_lay.setSpacing(8)

        self._chip_city_icon, self._chip_city_text = _chip(ICONS["map_pin"], ACCENT)
        self._chip_total_icon, self._chip_total_text = _chip(ICONS["hashtag"], FG)
        self._chip_dupes_icon, self._chip_dupes_text = _chip(ICONS["trash"], DUPE_RED)
        self._chip_remain_icon, self._chip_remain_text = _chip(ICONS["check"], GREEN)

        for i, (ic, tx) in enumerate((
            (self._chip_city_icon, self._chip_city_text),
            (self._chip_total_icon, self._chip_total_text),
            (self._chip_dupes_icon, self._chip_dupes_text),
            (self._chip_remain_icon, self._chip_remain_text),
        )):
            if i > 0:
                # Раньше здесь стоял символ «|» — на тёмной подложке он почти
                # не читался, а из-за рамки-наследования выглядел пустой
                # рамкой. Теперь это настоящая тонкая линия.
                sep = QFrame()
                sep.setFixedSize(1, 16)
                sep.setStyleSheet(f"background-color:{_veil(0.14)}; border:none;")
                stats_lay.addWidget(sep, 0, Qt.AlignVCenter)
            stats_lay.addWidget(ic)
            stats_lay.addWidget(tx)
        stats_lay.addStretch()
        self.preview_stats.setVisible(False)
        pfl.addWidget(self.preview_stats)
        lay.addWidget(self.pf)

        lay.addWidget(self._divider())

        btn_container = QWidget()
        btns = QHBoxLayout(btn_container)
        btns.setContentsMargins(0, 0, 0, 0)
        self.clean_btn = QPushButton("Удалить дубли и нарезать")
        self.clean_btn.setObjectName("green")
        self.clean_btn.setMinimumHeight(36)
        self.clean_btn.clicked.connect(self.run_clean)
        clean_btn = self.clean_btn
        btns.addWidget(clean_btn)

        self.shuffle_cb = QCheckBox("  Перемешать номера перед нарезкой")
        self.shuffle_cb.setIcon(_icon(ICONS["shuffle"], FG, ""))
        self.shuffle_cb.setIconSize(QSize(15, 15))
        btns.addSpacing(12)
        btns.addWidget(self.shuffle_cb)

        btns.addStretch()

        self.chunk_seg = DropdownButton(
            [("5 000", 5000), ("10 000", 10000), ("15 000", 15000), ("20 000", 20000)],
            default=5000, allow_custom=True)
        btns.addWidget(self.chunk_seg)

        # Каждая кнопка/чекбокс в этой строке держит свой естественный
        # размер (по тексту/шрифту) и не сжимается — раньше при узком окне
        # QHBoxLayout силой ужимал виджеты меньше их текста, отсюда и
        # "плывущий" шрифт. Если весь ряд не помещается — появляется тонкая
        # горизонтальная прокрутка вместо порчи текста (тот же приём, что и
        # в панели инструментов вкладки "Редактирование").
        for w in (clean_btn, self.shuffle_cb, self.chunk_seg):
            w.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

        btn_scroll = QScrollArea()
        btn_scroll.setWidget(btn_container)
        btn_scroll.setWidgetResizable(True)
        btn_scroll.setFixedHeight(48)
        btn_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        btn_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        btn_scroll.setStyleSheet(
            "QScrollArea { border:none; background:transparent; }"
            "QScrollBar:horizontal { height:4px; margin:0; }"
            f"QScrollBar::handle:horizontal {{ background:{BORDER}; border-radius:2px; min-width:20px; }}"
            "QScrollBar::add-line, QScrollBar::sub-line { width:0; }")
        lay.addWidget(btn_scroll)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(
            f"background-color:{SURFACE}; color:{LOG_FG}; "
            f"font-family:'Consolas'; font-size:11pt; border:1px solid {BORDER}; border-radius:10px;")
        lay.addWidget(self.log, 1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        lay.addWidget(self.progress)
        self._maybe_flip_layout(lay)
        return page

    def _flash_preview(self, has_dupes):
        """Запускает вспышку фона блока предпросмотра: красная — дубли найдены, зелёная — чисто."""
        flash_r, flash_g, flash_b = _hex_rgb(DUPE_RED if has_dupes else GREEN)
        base_r, base_g, base_b = _hex_rgb(BG3)
        total_steps = 18

        if self._flash_timer.isActive():
            self._flash_timer.stop()
        try:
            self._flash_timer.timeout.disconnect()
        except RuntimeError:
            pass

        self._flash_step = 0

        def _step():
            t = self._flash_step / total_steps
            alpha = 1.0 - abs(t * 2.0 - 1.0)
            alpha = min(max(alpha, 0.0), 1.0) * 0.45
            r = int(base_r + (flash_r - base_r) * alpha)
            g = int(base_g + (flash_g - base_g) * alpha)
            b = int(base_b + (flash_b - base_b) * alpha)
            self.pf.setStyleSheet(_preview_panel_qss(f"#{r:02x}{g:02x}{b:02x}"))
            self._flash_step += 1
            if self._flash_step > total_steps:
                self._flash_timer.stop()
                self.pf.setStyleSheet(_preview_panel_qss())

        self._flash_timer.timeout.connect(_step)
        self._flash_timer.start(35)

    def _flash_clean_btn(self, success):
        """Вспышка на кнопке «Удалить дубли и нарезать»: зелёная — успех, красная — ошибка."""
        if success:
            base_r, base_g, base_b = _hex_rgb(GREEN)
            peak_r, peak_g, peak_b = _hex_rgb(_mix(GREEN_HV, "#ffffff", 0.45))
        else:
            base_r, base_g, base_b = _hex_rgb(GRAY_BTN)
            peak_r, peak_g, peak_b = _hex_rgb(DUPE_RED)

        total_steps = 22

        if self._flash_btn_timer.isActive():
            self._flash_btn_timer.stop()
        try:
            self._flash_btn_timer.timeout.disconnect()
        except RuntimeError:
            pass

        self._flash_btn_step = 0
        orig_style = self.clean_btn.styleSheet()

        def _step():
            t = self._flash_btn_step / total_steps
            alpha = 1.0 - abs(t * 2.0 - 1.0)
            alpha = min(max(alpha, 0.0), 1.0)
            if success:
                r = int(base_r + (peak_r - base_r) * alpha)
                g = int(base_g + (peak_g - base_g) * alpha)
                b = int(base_b + (peak_b - base_b) * alpha)
            else:
                r = int(base_r + (peak_r - base_r) * alpha)
                g = int(base_g + (peak_g - base_g) * alpha)
                b = int(base_b + (peak_b - base_b) * alpha)
            self.clean_btn.setStyleSheet(
                f"background-color:#{r:02x}{g:02x}{b:02x}; color:{'white' if success else DUPE_RED_HV};"
                f"border:none; border-radius:10px; padding:9px 14px; font-weight:bold;")
            self._flash_btn_step += 1
            if self._flash_btn_step > total_steps:
                self._flash_btn_timer.stop()
                self.clean_btn.setStyleSheet(orig_style)

        self._flash_btn_timer.timeout.connect(_step)
        self._flash_btn_timer.start(35)

    def _set_progress(self, running):
        if running:
            self.progress.setRange(0, 0)   # неопределённый режим
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)

    def _append_log(self, msg):
        self.log.append(msg)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def log_msg(self, msg):
        self.sigLog.emit(msg)

    def _trigger_preview(self):
        if self._preview_running:
            return
        self._preview_running = True
        threading.Thread(target=self._run_preview, daemon=True).start()

    def _run_preview(self):
        try:
            called = self.called_row.get()
            uncalled = self.uncalled_row.get()
            if not called or not C.source_is_valid(called) or not uncalled or not os.path.isdir(uncalled):
                self.sigPreview.emit("", FG2)
                self.sigPreviewStatsHide.emit()
                return
            self.sigPreviewStatsHide.emit()
            self.sigPreview.emit("🔍 Считаю дубли…", FG2)
            phones = C.load_phones_from_source(called)
            if not phones:
                self.sigPreview.emit("⚠ Не найдено номеров в папке «Прозвоненные»", ACCENT)
                return
            uncalled_nums = C.load_uncalled_numbers(uncalled)
            if not uncalled_nums:
                self.sigPreview.emit("⚠ Не найдено файлов в папке «База для очистки»", ACCENT)
                return
            total = len(uncalled_nums)
            dupes = sum(1 for n in uncalled_nums if n in phones)
            remain = total - dupes
            city = C.get_city_from_source(called)
            self.sigPreview.emit("", FG2)
            self.sigPreviewStats.emit(city, total, dupes, remain)
            self.sigFlash.emit(dupes > 0)
        finally:
            self._preview_running = False

    def _update_preview_stats(self, city, total, dupes, remain):
        self._chip_city_text.setText(city)
        self._chip_total_text.setText(f"Всего: {total:,}".replace(",", " "))
        self._chip_dupes_text.setText(f"Дублей: {dupes:,}".replace(",", " "))
        dupes_color = DUPE_RED if dupes > 0 else FG2
        self._chip_dupes_text.setStyleSheet(f"color:{dupes_color}; font-weight:600;")
        self._chip_remain_text.setText(f"Останется: {remain:,}".replace(",", " "))
        self.preview_stats.setVisible(True)

    def _update_ready_path(self):
        """Пересчитывает путь «Результат» по текущей «Базе для очистки»:
        подпапка «без дублей <сегодняшняя дата>» прямо внутри неё. Саму
        папку на диске не создаём здесь — только показываем путь; реально
        создаётся она уже при сохранении (os.makedirs(..., exist_ok=True))."""
        base = self.uncalled_row.get()
        if base:
            today = datetime.date.today().strftime("%d.%m.%Y")
            path = os.path.join(base, f"без дублей {today}")
        else:
            path = ""
        self.ready_row.edit.setText(path)

    def run_clean(self):
        if not self.called_row.get() or not self.uncalled_row.get() or not self.ready_row.get():
            QMessageBox.critical(self, "Ошибка", "Выбери все три папки!")
            return

        # Читаем «Прозвоненные» сразу здесь (в основном потоке) — быстро, и
        # нужно ДО запуска фонового потока, чтобы успеть спросить у
        # пользователя, какие статусы считать дублями (диалог с чекбоксами
        # нельзя показывать из фонового потока).
        try:
            all_phones, by_status = C.load_phones_by_status_from_source(self.called_row.get())
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", self._file_err(
                "Не удалось прочитать «Прозвоненные»:", e))
            return

        named_statuses = sorted(s for s in by_status if s)
        exclude_phones = all_phones  # по умолчанию — как раньше, дубль есть дубль
        if named_statuses:
            chosen = self._show_clean_dup_status_dialog(by_status, named_statuses)
            if chosen is None:
                return  # отменили в диалоге
            exclude_phones = set(by_status.get("", set()))
            for s in chosen:
                exclude_phones |= by_status.get(s, set())

        threading.Thread(target=self._clean, args=(exclude_phones,), daemon=True).start()

    def _show_clean_dup_status_dialog(self, by_status, named_statuses):
        """Диалог со списком статусов, найденных среди «Прозвоненных»
        (пример: «Недозвон», «Перезвон», «Дозвон» и т.п.), с чекбоксами.
        Отмеченный статус = номера с этим статусом считаются дублями и
        удаляются из «Базы для очистки» и не попадают в результат. Если
        галочку снять — номера с этим статусом дублями не считаются и
        остаются в результате, даже если совпали с «Прозвоненными».

        Возвращает set выбранных (отмеченных) статусов, либо None, если
        пользователь нажал «Отмена».
        """
        dlg = QDialog(self)
        dlg.setWindowTitle("Какие статусы считать дублями?")
        dlg.setMinimumWidth(460)
        lay = QVBoxLayout(dlg)
        lay.setSpacing(10)

        hint = QLabel(
            "В «Прозвоненных» найдены разные статусы обзвона.\n"
            "Отметь те, чьи номера нужно УДАЛИТЬ как дубли — остальные\n"
            "статусы (без галочки) останутся в результате, даже если номер\n"
            "уже встречался в «Прозвоненных».")
        hint.setStyleSheet(f"color:{FG};")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setSpacing(4)
        checks = {}
        for status in named_statuses:
            cnt = len(by_status.get(status, ()))
            cb = QCheckBox(f"{status}   ({cnt:,})".replace(",", " "))
            cb.setChecked(True)  # по умолчанию — как раньше, удаляем всё
            checks[status] = cb
            inner_lay.addWidget(cb)
        inner_lay.addStretch()
        scroll.setWidget(inner)
        scroll.setMaximumHeight(280)
        lay.addWidget(scroll)

        quick_bar = QHBoxLayout()
        select_all = QPushButton("Отметить всё")
        select_all.clicked.connect(lambda: [cb.setChecked(True) for cb in checks.values()])
        select_none = QPushButton("Снять всё")
        select_none.clicked.connect(lambda: [cb.setChecked(False) for cb in checks.values()])
        quick_bar.addWidget(select_all)
        quick_bar.addWidget(select_none)
        quick_bar.addStretch()
        lay.addLayout(quick_bar)

        no_status_cnt = len(by_status.get("", ()))
        if no_status_cnt:
            note = QLabel(
                f"Ещё {no_status_cnt:,} номер(ов) без определённого статуса — они, "
                "как и раньше, считаются дублями всегда.".replace(",", " "))
            note.setStyleSheet(f"color:{FG2};")
            note.setWordWrap(True)
            lay.addWidget(note)

        bar = QHBoxLayout()
        ok_btn = QPushButton("  Продолжить")
        ok_btn.setObjectName("green")
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(dlg.reject)
        ok_btn.clicked.connect(dlg.accept)
        bar.addStretch()
        bar.addWidget(cancel_btn)
        bar.addWidget(ok_btn)
        lay.addLayout(bar)

        if dlg.exec() != QDialog.Accepted:
            return None
        return {s for s, cb in checks.items() if cb.isChecked()}

    def _save_chunks(self, df, city, chunk):
        os.makedirs(self.ready_row.get(), exist_ok=True)
        if self.shuffle_cb.isChecked():
            df = df.sample(frac=1).reset_index(drop=True)
            self.log_msg("🔀 Номера перемешаны.")
        total_files = 0
        for i, start in enumerate(range(0, len(df), chunk), 1):
            part = df.iloc[start:start + chunk]
            out = os.path.join(self.ready_row.get(), f"{city} {i:02d}.xlsx")
            part.to_excel(out, index=False)
            self.log_msg(f"  ✅ {city} {i:02d}.xlsx  ({len(part):,} строк)".replace(",", " "))
            total_files += 1
        return total_files


    def _clean(self, exclude_phones):
        self.sigProgress.emit(True)
        try:
            self.log_msg("=" * 52)
            city = C.get_city_from_source(self.called_row.get())
            self.log_msg(f"🏙  Город: {city}")
            self.log_msg(
                f"📞 Прозвоненных номеров к удалению как дубли: {len(exclude_phones):,}"
                .replace(",", " "))
            self.log_msg("")
            all_dfs = []
            for f in sorted(os.listdir(self.uncalled_row.get())):
                if not f.endswith(".xlsx"):
                    continue
                path = os.path.join(self.uncalled_row.get(), f)
                try:
                    df = pd.read_excel(path, header=None)
                    df.columns = ["Номер"]
                    df["Номер"] = df["Номер"].astype(str).str.lstrip("+").str.strip()
                    self.log_msg(f"📂 {f}  ({len(df):,})".replace(",", " "))
                    all_dfs.append(df)
                except Exception as e:
                    self.log_msg(f"  ⚠ Ошибка: {e}")
            if not all_dfs:
                self.log_msg("⚠ Нет файлов для обработки!")
                return
            combined = pd.concat(all_dfs, ignore_index=True)
            before = len(combined)
            combined = combined[~combined["Номер"].isin(exclude_phones)]
            removed = before - len(combined)

            self.log_msg("")
            self.log_msg(f"📊 Всего номеров  : {before:,}".replace(",", " "))
            self.log_msg(f"🗑  Дублей удалено : {removed:,}".replace(",", " "))
            self.log_msg(f"✅ Осталось      : {len(combined):,}".replace(",", " "))
            self.log_msg("")
            chunk = self.chunk_seg.value()
            total_files = self._save_chunks(combined, city, chunk)
            self.log_msg("")
            self.log_msg("=" * 52)
            self.log_msg(f"✔ ГОТОВО!  Файлов создано: {total_files}")
            self.log_msg("=" * 52)
            done_text = f"Город: {city}\nУдалено дублей: {removed:,}".replace(",", " ")
            done_text += f"\nОсталось: {len(combined):,}\nФайлов создано: {total_files}".replace(",", " ")
            self.sigCleanDone.emit("Готово!", done_text, "")
            self.sigFlashBtn.emit(True)
        except Exception as _exc:
            self.sigToast.emit("Не удалось очистить базу", str(_exc), "err")
            self.sigFlashBtn.emit(False)
            raise
        finally:
            self.sigProgress.emit(False)
            self._trigger_preview()

    # ═══════════════════════════  ВКЛАДКА «ГОРОДА»  ═══════════════════════════
    def _build_city_tab(self):
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(20, 8, 20, 10)
        lay.setSpacing(6)

        row = QHBoxLayout()
        row.addWidget(QLabel("Город:"))
        self.city_edit = QLineEdit()
        self.city_edit.returnPressed.connect(self.run_city_search)
        row.addWidget(self.city_edit, 1)
        find = QPushButton("Найти")
        find.setObjectName("accent")
        find.clicked.connect(self.run_city_search)
        row.addWidget(find)
        lay.addLayout(row)

        rad_row = QHBoxLayout()
        rad_row.addWidget(QLabel("Радиус:"))
        self.radius_seg = Segmented(
            [("100 км", 100), ("200 км", 200), ("300 км", 300),
             ("500 км", 500), ("1000 км", 1000), ("Всё", 99999)],
            default=300)
        self.radius_seg.changed.connect(self._on_radius)
        rad_row.addWidget(self.radius_seg)
        rad_row.addStretch()
        lay.addLayout(rad_row)

        spec_row = QHBoxLayout()
        spec_row.addWidget(QLabel("Специализация:"))
        self.spec_edit = QLineEdit()
        self.spec_edit.setPlaceholderText("напр. ВПК, ракеты, агро, ядерщики, атом…")
        self.spec_edit.returnPressed.connect(self.run_spec_search)
        spec_row.addWidget(self.spec_edit, 1)
        spec_find = QPushButton("Найти по специализации")
        spec_find.setObjectName("accent")
        spec_find.clicked.connect(self.run_spec_search)
        spec_row.addWidget(spec_find)
        lay.addLayout(spec_row)

        self.city_status = QLabel("Введи областной центр и нажми «Найти».")
        self.city_status.setStyleSheet(f"color:{FG2}; font-style:italic;")
        lay.addWidget(self.city_status)

        cols = ["км", "код", "Город", "Область", "Население", "Специализация"]
        self.city_table = QTableWidget(0, len(cols))
        self.city_table.setHorizontalHeaderLabels(cols)
        self.city_table.setSortingEnabled(True)
        self.city_table.verticalHeader().setVisible(False)
        self.city_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.city_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.city_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.city_table.customContextMenuRequested.connect(self._city_context_menu)
        self.city_table.cellDoubleClicked.connect(self._on_city_dbl)
        hh = self.city_table.horizontalHeader()
        for i in range(len(cols)):
            hh.setSectionResizeMode(i, QHeaderView.Interactive)
        hh.setStretchLastSection(True)
        self.city_table.setColumnWidth(0, 60)
        self.city_table.setColumnWidth(1, 76)
        self.city_table.setColumnWidth(2, 150)
        self.city_table.setColumnWidth(3, 150)
        self.city_table.setColumnWidth(4, 100)
        sc = QShortcut(QKeySequence.Copy, self.city_table)
        sc.activated.connect(self._copy_city_list)
        lay.addWidget(self.city_table, 1)

        self.city_count = QLabel("")
        self.city_count.setStyleSheet(f"color:{ACCENT}; font-weight:bold;")
        lay.addWidget(self.city_count)

        # ── Коды городов: ручной поиск «7 + код» и вытягивание из папки ──
        lay.addWidget(self._divider())

        crow = QHBoxLayout()
        self.code_edit = QLineEdit()
        self.code_edit.setPlaceholderText("напр. Ростов → 7863")
        self.code_edit.returnPressed.connect(self._lookup_single_code)
        self.code_edit.setFixedWidth(260)
        crow.addWidget(self.code_edit)
        code_btn = QPushButton("Код")
        code_btn.setToolTip("Найти код города (формат: 7 + код региона)")
        code_btn.clicked.connect(self._lookup_single_code)
        crow.addWidget(code_btn)
        self.code_result = QLineEdit()
        self.code_result.setReadOnly(True)
        self.code_result.setFixedWidth(110)
        self.code_result.setAlignment(Qt.AlignCenter)
        self.code_result.setPlaceholderText("7863")
        crow.addWidget(self.code_result)
        crow.addStretch()
        lay.addLayout(crow)

        prow = QHBoxLayout()
        pull_btn = QPushButton("  Вытащить коды из папки…")
        pull_btn.setIcon(_icon(ICONS["folder"], BG, ""))
        pull_btn.setIconSize(QSize(16, 16))
        pull_btn.setObjectName("accent")
        pull_btn.clicked.connect(self._pull_codes_from_folder)
        prow.addWidget(pull_btn)
        prow.addStretch()
        lay.addLayout(prow)
        self._maybe_flip_layout(lay)
        return page

    def _on_radius(self, _value):
        if self.city_edit.text().strip():
            QTimer.singleShot(60, self.run_city_search)

    def run_city_search(self):
        threading.Thread(target=self._safe_city_search, daemon=True).start()

    def _safe_city_search(self):
        try:
            self._city_search()
        except Exception as e:
            self.sigCityStatus.emit(f"⚠ Сбой поиска: {e}", DUPE_RED)
            self.sigToast.emit("Сбой поиска", str(e), "err")

    def _city_search(self):
        try:
            self.sigCityStatus.emit("⏳ Загружаю базы…", FG2)
            self._ensure_data()
        except Exception as e:
            self.sigCityStatus.emit(f"⚠ Ошибка загрузки данных: {e}", ACCENT)
            return
        q = C._norm(self.city_edit.text())
        if not q:
            self.sigCityStatus.emit("⚠ Введи название города.", ACCENT)
            return
        matches = self.city_index.get(q)
        if not matches:
            cand = [c for c in self.cities if C._norm(c["n"]).startswith(q)] \
                   or [c for c in self.cities if q in C._norm(c["n"])]
            if not cand:
                self.sigCityStatus.emit(f"⚠ Город «{self.city_edit.text().strip()}» не найден.", ACCENT)
                return
            matches = cand
        center = max(matches, key=lambda c: c["pop"] or 0)
        rad = self.radius_seg.value()
        rows = []
        for c in self.cities:
            km = 0 if c is center else C.haversine(center["lat"], center["lon"], c["lat"], c["lon"])
            if c is center or km <= rad:
                key = f"{c['n']}|{c['sub']}"
                code = self.ph_codes.get(key, "")
                spec = C.INDUSTRY.get(c["n"], "")
                rows.append((round(km), code, c["n"], c["sub"], c["pop"] or 0, spec, key))
        rows.sort(key=lambda r: r[0])
        self._last_rows = rows
        rlabel = "без ограничения" if rad >= 99999 else f"{rad} км"
        self.sigCityFill.emit(rows)
        self.sigCityFlash.emit()   # вспышка при смене списка городов
        self.sigCityStatus.emit(
            f"Центр: {center['n']} ({center['sub']}) • радиус {rlabel} • двойной клик → диапазоны", GREEN)
        self.sigCityCount.emit(f"Найдено городов: {len(rows)}")

    def run_spec_search(self):
        threading.Thread(target=self._safe_spec_search, daemon=True).start()

    def _safe_spec_search(self):
        try:
            self._spec_search()
        except Exception as e:
            self.sigCityStatus.emit(f"⚠ Сбой поиска по специализации: {e}", DUPE_RED)
            self.sigToast.emit("Сбой поиска по специализации", str(e), "err")

    def _spec_search(self):
        try:
            self.sigCityStatus.emit("⏳ Загружаю базы…", FG2)
            self._ensure_data()
        except Exception as e:
            self.sigCityStatus.emit(f"⚠ Ошибка загрузки данных: {e}", ACCENT)
            return
        query = self.spec_edit.text().strip()
        if not query:
            self.sigCityStatus.emit("⚠ Введи слово специализации (например «ВПК», «ракеты», «агро»).", ACCENT)
            return
        found = C.search_by_specialization(query, self.cities)
        if not found:
            self.sigCityStatus.emit(f"⚠ По запросу «{query}» ничего не найдено.", ACCENT)
            self.sigCityFill.emit([])
            self.sigCityCount.emit("Найдено городов: 0")
            return
        rows = []
        for c in found:
            key = f"{c['n']}|{c['sub']}"
            code = self.ph_codes.get(key, "")
            spec = C.INDUSTRY.get(c["n"], "")
            rows.append((0, code, c["n"], c["sub"], c["pop"] or 0, spec, key))
        self._last_rows = rows
        self.sigCityFill.emit(rows)
        self.sigCityFlash.emit()
        self.sigCityStatus.emit(
            f"Специализация: «{query}» • по всей базе (без учёта центра/радиуса) • двойной клик → диапазоны", GREEN)
        self.sigCityCount.emit(f"Найдено городов: {len(rows)}")

    def _flash_city_table(self):
        """Синяя подсветка таблицы городов с плавным затуханием за ~1 сек.
        Реализован через QWidget-оверлей поверх таблицы — stylesheet не трогается
        вообще, поэтому «серой двойной вспышки» быть не может."""
        tbl = self.city_table

        # Ленивая инициализация оверлея
        if not hasattr(self, "_city_glow"):
            self._city_glow = FlashOverlay(tbl, ACCENT_HV, radius=STYLE["radius_lg"])
            self._city_glow.show()

        self._city_glow.resize(tbl.size())
        self._city_glow.raise_()
        self._city_glow.flash(0.26, 900)

    def _fill_city_table(self, rows):
        t = self.city_table
        t.setSortingEnabled(False)
        t.setRowCount(len(rows))
        self._row_keys = {}
        for i, r in enumerate(rows):
            km, code, name, sub, pop, spec, key = r
            t.setItem(i, 0, NumItem(str(km), km))
            t.setItem(i, 1, txt_item(code or "—"))
            nm = txt_item(name)
            t.setItem(i, 2, nm)
            t.setItem(i, 3, txt_item(sub))
            t.setItem(i, 4, NumItem(f"{pop:,}".replace(",", " "), pop))
            t.setItem(i, 5, txt_item(spec or "—"))
            self._row_keys[i] = (key, name, sub, code)
        t.setSortingEnabled(True)

    def _row_meta(self, visual_row):
        # получить (key, name, sub, code) по текущей строке (учитывая сортировку)
        name = self.city_table.item(visual_row, 2).text()
        sub = self.city_table.item(visual_row, 3).text()
        code = self.city_table.item(visual_row, 1).text()
        code = "" if code == "—" else code
        key = f"{name}|{sub}"
        return key, name, sub, code

    def _on_city_dbl(self, r, _c):
        key, name, sub, code = self._row_meta(r)
        self.open_ranges_popup(key, name, sub, code)

    def _city_context_menu(self, pos):
        idx = self.city_table.indexAt(pos)
        if not idx.isValid():
            return
        r = idx.row()
        key, name, sub, code = self._row_meta(r)
        menu = QMenu(self)
        a1 = menu.addAction(f"📋  Скопировать «{name}»")
        a2 = menu.addAction("➤  Вставить в поле поиска")
        a3 = menu.addAction("☎  Открыть диапазоны номеров")
        act = menu.exec(self.city_table.viewport().mapToGlobal(pos))
        if act == a1:
            QApplication.clipboard().setText(name)
            self._set_label(self.city_status, f"📋 «{name}» скопировано.", GREEN)
        elif act == a2:
            self.city_edit.setText(name)
            self._set_label(self.city_status, f"➤ «{name}» вставлен в поиск. Нажми «Найти».", ACCENT)
        elif act == a3:
            self.open_ranges_popup(key, name, sub, code)

    def _copy_city_list(self):
        t = self.city_table
        rows = sorted({i.row() for i in t.selectedIndexes()}) or range(t.rowCount())
        lines = []
        for r in rows:
            km = t.item(r, 0).text()
            code = t.item(r, 1).text()
            name = t.item(r, 2).text()
            sub = t.item(r, 3).text()
            pop = t.item(r, 4).text()
            spec = t.item(r, 5).text()
            base = f"{km} км — {('☎'+code+' — ') if code != '—' else ''}{pop} — {name} ({sub})"
            lines.append(base + (f" — {spec}" if spec != '—' else ""))
        if lines:
            QApplication.clipboard().setText("\n".join(lines))
            self._set_label(self.city_count, f"Скопировано строк: {len(lines)} ✅", GREEN)

    # ── Коды городов «7 + код» ──
    def _code_for_city(self, raw):
        """Подобрать город по названию и вернуть (код, город_dict) либо ('', None).
        Среди точных и «начинается на» совпадений берём самый крупный по населению —
        чтобы «Ростов» дал Ростов-на-Дону, «Киров» — Киров (Кировская обл.) и т.п."""
        q = C._norm(raw or "")
        if not q:
            return "", None
        exact = list(self.city_index.get(q) or
                     [c for c in self.cities if C._norm(c["n"]) == q])
        seen = {id(c) for c in exact}
        pool = exact + [c for c in self.cities
                        if id(c) not in seen and C._norm(c["n"]).startswith(q)]
        if not pool:
            pool = [c for c in self.cities if q in C._norm(c["n"])]
        if not pool:
            return "", None
        center = max(pool, key=lambda c: c["pop"] or 0)
        code = self.ph_codes.get(f"{center['n']}|{center['sub']}", "")
        return code, center

    def _lookup_single_code(self):
        try:
            self._ensure_data()
        except Exception as e:
            self._set_label(self.city_status, f"⚠ Ошибка загрузки данных: {e}", ACCENT)
            return
        raw = self.code_edit.text().strip()
        if not raw:
            self.code_result.setText("")
            return
        code, center = self._code_for_city(raw)
        if center is None:
            self.code_result.setText("—")
            self._set_label(self.city_status, f"⚠ Город «{raw}» не найден.", ACCENT)
            return
        if not code:
            self.code_result.setText("—")
            self._set_label(self.city_status,
                            f"⚠ Для «{center['n']} ({center['sub']})» кода нет в базе.", ACCENT)
            return
        out = "7" + code
        self.code_result.setText(out)
        QApplication.clipboard().setText(out)
        self._set_label(self.city_status,
                        f"☎ {center['n']} ({center['sub']}) → {out}  (скопировано)", GREEN)

    def _extract_region_code_from_file(self, filepath):
        """Читает номера из файла (первый столбец) и возвращает наиболее часто
        встречающийся 4-значный префикс — это и есть код региона
        (например «7863», «7842», «7342» — первая цифра «7» и код города).
        Возвращает '' если валидных номеров не нашлось."""
        try:
            if filepath.lower().endswith(".csv"):
                df = pd.read_csv(filepath, header=None)
            else:
                df = pd.read_excel(filepath, header=None)
        except Exception:
            return ""
        from collections import Counter
        cnt = Counter()
        for val in df.iloc[:, 0].dropna():
            s = str(val).lstrip("+").strip()
            if s.isdigit() and len(s) >= 4:
                cnt[s[:4]] += 1
        if not cnt:
            return ""
        return cnt.most_common(1)[0][0]

    @staticmethod
    def _smart_read_table(path):
        """Читает xlsx/csv, сам решая, есть ли строка заголовков: если самая
        первая ячейка уже похожа на настоящее ФИО (два/три слова с заглавной
        буквы) — считаем, что шапки нет, и подписываем столбцы «Столбец A/B/C…».
        (Проверять всю строку на числа ненадёжно — заголовок вполне может
        называться просто «1».)"""
        def _read(header):
            if str(path).lower().endswith(".csv"):
                try:
                    return pd.read_csv(path, header=header, dtype=str, sep=None, engine="python").fillna("")
                except Exception:
                    return pd.read_csv(path, header=header, dtype=str, sep=";").fillna("")
            return pd.read_excel(path, header=header, dtype=str).fillna("")

        probe = _read(None)
        looks_like_data = False
        if len(probe) > 0:
            first_cell = str(probe.iloc[0, 0]).strip()
            if re.match(r"^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+)?$", first_cell):
                looks_like_data = True
        if looks_like_data:
            df = probe
            df.columns = [f"Столбец {chr(65 + i)}" for i in range(df.shape[1])]
            return df
        df = _read(0)
        return df

    def _pull_codes_from_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Папка с файлами городов")
        if not folder:
            return
        try:
            files = sorted(os.listdir(folder))
        except Exception as e:
            self._set_label(self.city_status, f"⚠ Не удалось прочитать папку: {e}", ACCENT)
            return
        results = []       # (город_из_файла, код_региона «7ХХХ»)
        seen = set()
        for f in files:
            if not f.lower().endswith((".xlsx", ".xls", ".csv")):
                continue
            stem = os.path.splitext(f)[0]
            # убрать хвостовые индексы/номера: «Ростов 01», «Ростов_2», «Ростов-3» → «Ростов»
            city = re.sub(r'[\s_\-]*\d+\s*$', '', stem).strip() or stem.strip()
            nkey = C._norm(city)
            if not nkey or nkey in seen:
                continue
            seen.add(nkey)
            path = os.path.join(folder, f)
            code = self._extract_region_code_from_file(path)
            results.append((city, code))
        if not results:
            self._set_label(self.city_status,
                            "⚠ В папке нет файлов с номерами (.xlsx/.xls/.csv).", ACCENT)
            return
        self._show_codes_dialog(folder, results)

    def _show_codes_dialog(self, folder, results):
        dlg = QDialog(self)
        dlg.setWindowTitle("Коды регионов из папки")
        dlg.resize(560, 560)
        dlg.setStyleSheet(build_qss())
        lay = QVBoxLayout(dlg)

        base = os.path.basename(folder.rstrip("/\\")) or folder
        title = QLabel(f"📂  {base}")
        title.setStyleSheet(f"color:{FG}; font-weight:bold; font-size:13pt;")
        lay.addWidget(title)

        resolved = [(c, code) for (c, code) in results if code]
        # По одному коду в строке (сверху вниз), а не через пробел в одну строку
        codes_str = "\n".join(code for _c, code in resolved)

        sub = QLabel(f"Файлов: {len(results)}   •   с кодом: {len(resolved)}")
        sub.setStyleSheet(f"color:{FG2};")
        lay.addWidget(sub)

        lay.addWidget(self._section("КОДЫ РЕГИОНОВ (по одному в строке)"))
        box = QPlainTextEdit()
        box.setPlainText(codes_str)
        box.setReadOnly(False)
        box.setMinimumHeight(160)
        lay.addWidget(box, 1)

        tbl = QTableWidget(0, 2)
        tbl.setHorizontalHeaderLabels(["Город (файл)", "Код региона"])
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.setColumnWidth(0, 260)
        tbl.setColumnWidth(1, 120)
        tbl.setRowCount(len(results))
        for i, (city, code) in enumerate(results):
            tbl.setItem(i, 0, txt_item(city))
            tbl.setItem(i, 1, txt_item(code or "не найден"))
        lay.addWidget(tbl, 1)

        bar = QHBoxLayout()
        copy_btn = QPushButton("  Копировать коды")
        copy_btn.setIcon(_icon(ICONS["copy"], BG, ""))
        copy_btn.setIconSize(QSize(15, 15))
        copy_btn.setObjectName("green")
        def _copy():
            QApplication.clipboard().setText(box.toPlainText())
            sub.setText(f"Скопировано ✅")
            sub.setStyleSheet(f"color:{GREEN}; font-weight:bold;")
        copy_btn.clicked.connect(_copy)
        bar.addWidget(copy_btn)
        bar.addStretch()
        close = QPushButton("Закрыть")
        close.setObjectName("accent")
        close.clicked.connect(dlg.accept)
        bar.addWidget(close)
        lay.addLayout(bar)
        dlg.exec()

    def open_ranges_popup(self, key, name, sub, code):
        try:
            self._ensure_data()
        except Exception:
            pass
        ranges = (self.ph_ranges or {}).get(key, [])
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{name} — телефонные диапазоны")
        dlg.resize(800, 580)
        dlg.setStyleSheet(build_qss())
        lay = QVBoxLayout(dlg)

        title = QLabel(f"☎  {name}")
        title.setStyleSheet(f"color:{FG}; font-weight:bold; font-size:13pt;")
        lay.addWidget(title)
        codetxt = code if code else "—"
        sub_l = QLabel(f"{sub}   •   Код города: +7 {codetxt}   •   двойной клик по строке → callzone.info")
        sub_l.setStyleSheet(f"color:{FG2};")
        lay.addWidget(sub_l)

        if not ranges:
            empty = QLabel("Для этого города в базе Россвязи нет городских диапазонов\n"
                           "(либо все они принадлежат мобильным операторам).")
            empty.setStyleSheet(f"color:{FG2}; font-style:italic;")
            lay.addWidget(empty)
            lay.addStretch()
            close = QPushButton("Закрыть")
            close.setObjectName("accent")
            close.clicked.connect(dlg.accept)
            lay.addWidget(close, alignment=Qt.AlignRight)
            dlg.exec()
            return

        info = QLabel(f"ВСЕГО ДИАПАЗОНОВ: {len(ranges)}")
        info.setStyleSheet(f"color:{ACCENT}; font-weight:bold; font-size:10pt;")
        lay.addWidget(info)

        cols = ["Диапазон", "Номеров", "Оператор", "Тип"]
        tbl = QTableWidget(0, 4)
        tbl.setHorizontalHeaderLabels(cols)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.setColumnWidth(0, 260)
        tbl.setColumnWidth(1, 90)
        tbl.setColumnWidth(2, 160)
        tbl.setColumnWidth(3, 70)
        tbl.setSelectionBehavior(QAbstractItemView.SelectItems)
        tbl.setSelectionMode(QAbstractItemView.ExtendedSelection)

        rows_sorted = sorted(ranges, key=lambda r: r[0])
        row_map = {}
        tbl.setRowCount(len(rows_sorted))
        for i, r in enumerate(rows_sorted):
            frm, span, cap, opidx, kind = r
            to = frm + span
            op = self.ph_ops[opidx] if 0 <= opidx < len(self.ph_ops) else "?"
            typ = "город" if kind == 0 else "район"
            tbl.setItem(i, 0, txt_item(f"{C.fmt_num(frm)}  →  {C.fmt_num(to)}"))
            tbl.setItem(i, 1, txt_item(f"{cap:,}".replace(",", " ")))
            tbl.setItem(i, 2, txt_item(op))
            tbl.setItem(i, 3, txt_item(typ))
            row_map[i] = frm
        tbl.cellDoubleClicked.connect(
            lambda rr, _c: webbrowser.open(f"https://callzone.info/7{row_map[rr]}"))
        lay.addWidget(tbl, 1)

        hint = QLabel("Ctrl+C — скопировать только выделенные ячейки")
        hint.setStyleSheet(f"color:{FG2}; font-size:9pt;")
        lay.addWidget(hint)

        def _range_line(r):
            frm, span, cap, opidx, kind = r
            op = self.ph_ops[opidx] if 0 <= opidx < len(self.ph_ops) else "?"
            typ = "город" if kind == 0 else "район"
            return f"{C.fmt_num(frm)} до {C.fmt_num(frm+span)}  {op} ({cap} номеров) [{typ}]"

        def copy_selected():
            idxs = tbl.selectedIndexes()
            if not idxs:
                return
            by_row = {}
            for idx in idxs:
                by_row.setdefault(idx.row(), {})[idx.column()] = idx
            lines = []
            for r in sorted(by_row):
                cols = by_row[r]
                cells = [tbl.item(r, c).text() if tbl.item(r, c) else ""
                         for c in sorted(cols)]
                lines.append("\t".join(cells))
            QApplication.clipboard().setText("\n".join(lines))
            info.setText(f"Скопировано ячеек: {len(idxs)} ✅")
            info.setStyleSheet(f"color:{GREEN}; font-weight:bold;")

        copy_sc = QShortcut(QKeySequence.Copy, tbl)
        copy_sc.setContext(Qt.WidgetShortcut)
        copy_sc.activated.connect(copy_selected)

        bar = QHBoxLayout()
        copy_btn = QPushButton("  Копировать все диапазоны")
        copy_btn.setIcon(_icon(ICONS["copy"], FG, ""))
        copy_btn.setIconSize(QSize(15, 15))
        def copy_ranges():
            lines = [_range_line(r) for r in rows_sorted]
            QApplication.clipboard().setText("\n".join(lines))
            info.setText(f"Скопировано диапазонов: {len(lines)} ✅")
            info.setStyleSheet(f"color:{GREEN}; font-weight:bold;")
        copy_btn.clicked.connect(copy_ranges)
        close = QPushButton("Закрыть")
        close.setObjectName("accent")
        close.clicked.connect(dlg.accept)
        bar.addWidget(copy_btn)
        bar.addStretch()
        bar.addWidget(close)
        lay.addLayout(bar)
        dlg.exec()

    # ═══════════════════════════  ВКЛАДКА «РЕДАКТИРОВАНИЕ»  ═══════════════════════════
    def _build_edit_tab(self):
        page = QWidget()
        page.setAcceptDrops(True)
        page.dragEnterEvent = self._edit_drag_enter
        page.dropEvent = self._edit_drop
        lay = QVBoxLayout(page)
        lay.setContentsMargins(20, 10, 20, 10)
        lay.setSpacing(6)

        # Тулбар: компактные кнопки-иконки с glow-эффектом и тултипами
        _btn_container = QWidget()
        _btn_container.setStyleSheet("background:transparent;")
        top = QHBoxLayout(_btn_container)
        top.setContentsMargins(0, 2, 0, 2)
        top.setSpacing(6)

        # Иконки: qtawesome если доступен, иначе unicode-запасной вариант
        # Через _icon() — а не напрямую через qta.icon() — на случай, если
        # конкретное имя иконки вдруг окажется недоступно в установленной
        # версии qtawesome: тогда вместо падения всего приложения при
        # запуске мы просто получим безопасную запасную иконку.
        _csv_icon   = _icon(ICONS["excel"], "#1D6F42", "📊")
        _merge_icon = _icon(ICONS["shuffle"], GREEN, "🔀")

        b_open   = IconBtn(_icon(ICONS["folder"], FG, "📁"), "Открыть файл")
        b_merge  = IconBtn(_merge_icon, "Объединить файлы (можно сразу несколько городов, до 40)", variant="green")
        b_undo   = IconBtn(_undo_icon(22, FG), "Отменить")
        b_trig   = IconBtn(_icon(ICONS["warning"], DUPE_RED, "⚠"), "Триггер-слова",    variant="danger")
        b_words  = IconBtn(_icon(ICONS["list"], FG, "✎"), "Список слов")
        b_macro  = IconBtn(_icon(ICONS["filter"], FG, "⚙"), "Макрос: убрать организации из Excel (можно сразу до 20 файлов)")
        b_status = IconBtn(_icon(ICONS["status"], ACCENT, "📋"), "Статусы: разобрать номера по статусу обзвона")
        b_calls  = IconBtn(_icon(ICONS["phone"], GREEN, "☎"),
                           "Дозвон: убрать недозвон и автоответчик, оставить дозвон "
                           "(выгрузки со столбцами Phone / Result / Result+)", variant="green")
        b_export = IconBtn(_csv_icon,   "Выгрузить CSV",    variant="green")
        b_clear  = IconBtn(_icon(ICONS["eraser"], DUPE_RED, "✕"), "Очистить таблицу", variant="danger")

        b_open.clicked.connect(self._load_edit_file)
        b_merge.clicked.connect(self._load_merge_files)
        b_undo.clicked.connect(self._undo_edit)
        b_trig.clicked.connect(self._open_trigger_scan)
        b_words.clicked.connect(self._open_trigger_editor)
        b_macro.clicked.connect(self._run_company_filter_macro)
        b_status.clicked.connect(self._run_statuses_tool)
        b_calls.clicked.connect(self._run_calls_tool)
        b_export.clicked.connect(self._export_edit_csv)
        b_clear.clicked.connect(self._clear_edit_table)

        for b in (b_open, b_merge, b_undo, b_trig, b_words, b_macro, b_status, b_calls):
            top.addWidget(b)
        top.addStretch()
        top.addWidget(b_export)
        top.addWidget(b_clear)

        _btn_scroll = QScrollArea()
        _btn_scroll.setWidget(_btn_container)
        _btn_scroll.setWidgetResizable(True)
        _btn_scroll.setFixedHeight(46)
        _btn_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        _btn_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        _btn_scroll.setStyleSheet(
            "QScrollArea { border:none; background:transparent; }"
            "QScrollBar:horizontal { height:4px; margin:0; }"
            f"QScrollBar::handle:horizontal {{ background:{BORDER}; border-radius:2px; min-width:20px; }}"
            "QScrollBar::add-line, QScrollBar::sub-line { width:0; }")
        lay.addWidget(_btn_scroll)

        info_row = QHBoxLayout()
        self.edit_status = QLabel("")
        self.edit_status.setStyleSheet(f"color:{FG2};")
        self.edit_count = QLabel("")
        self.edit_count.setStyleSheet(f"color:{ACCENT}; font-weight:bold;")
        info_row.addWidget(self.edit_status)
        info_row.addWidget(self.edit_count)
        info_row.addStretch()
        lay.addLayout(info_row)

        # Городá-«чипы» — отдельная строка НАД таблицей, с переносом на
        # несколько строк, если городов много (даже 20): раньше это был один
        # QLabel в один ряд с текстом вроде «топ-3 + …», из-за чего остальные
        # города были просто не видны. Теперь каждый город — свой маленький
        # компактный виджет, и когда места по ширине не хватает, лишние чипы
        # просто уходят на следующую строку, а не обрезаются.
        self.edit_city_box = QWidget()
        self.edit_city_flow = FlowLayout(self.edit_city_box, margin=0, h_spacing=6, v_spacing=6)
        self.edit_city_box.setVisible(False)
        lay.addWidget(self.edit_city_box)

        self.edit_table = QTableWidget(0, 4)
        self.edit_table.setHorizontalHeaderLabels(["Номер телефона", "Транскрибация клиента", "Оператор", "Город"])
        self.edit_table.verticalHeader().setVisible(False)
        self.edit_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.edit_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.edit_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.edit_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.edit_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.edit_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.edit_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self.edit_table.horizontalHeader().setStretchLastSection(True)
        self.edit_table.setColumnWidth(0, 170)
        self.edit_table.setColumnWidth(1, 340)
        self.edit_table.setColumnWidth(2, 130)
        self.edit_table.setColumnWidth(3, 130)
        self.edit_table.verticalHeader().setDefaultSectionSize(34)
        self.edit_table.cellDoubleClicked.connect(self._on_edit_dbl)
        self.edit_table.horizontalHeader().sectionDoubleClicked.connect(self._on_edit_header_dbl)
        del_sc = QShortcut(QKeySequence(Qt.Key_Delete), self.edit_table)
        del_sc.activated.connect(self._del_edit_rows)
        copy_sc = QShortcut(QKeySequence.Copy, self.edit_table)
        copy_sc.activated.connect(self._copy_edit_numbers)
        lay.addWidget(self.edit_table, 1)
        self._maybe_flip_layout(lay)
        return page

    def _edit_drag_enter(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def _edit_drop(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls()]
        if paths:
            self._process_edit_file(paths[0])

    def _set_edit_rows(self, pairs):
        t = self.edit_table
        try:
            self._ensure_data()
        except Exception:
            pass
        t.setRowCount(0)
        t.setRowCount(len(pairs))
        city_items = []
        for i, (phone, trans) in enumerate(pairs):
            t.setItem(i, 0, txt_item(str(phone).strip()))
            t.setItem(i, 1, txt_item(trans))
            try:
                city, op = self._lookup_phone_city_operator(phone)
            except Exception:
                city, op = "", ""
            t.setItem(i, 2, txt_item(op))
            city_item = txt_item(city)
            t.setItem(i, 3, city_item)
            city_items.append((city, city_item))
        self._colorize_edit_cities(city_items)
        self._update_edit_count()

    def _colorize_edit_cities(self, city_items):
        """Если после объединения файлов в таблице оказалось НЕСКОЛЬКО разных
        городов — красим название каждого города в столбце «Город» своим
        цветом (один и тот же город — всегда один и тот же цвет), чтобы при
        прокрутке сразу было видно, где какой город, не вчитываясь в текст.
        Если город один (или не определился ни у кого) — обычный цвет,
        подсветка тут не нужна."""
        distinct = sorted({c for c, _ in city_items if c})
        if len(distinct) <= 1:
            return
        colors = self._distinct_city_colors(len(distinct))
        color_map = dict(zip(distinct, colors))
        for city, item in city_items:
            color = color_map.get(city)
            if color is None:
                continue
            item.setForeground(color)
            f = item.font()
            f.setBold(True)
            item.setFont(f)

    @staticmethod
    def _distinct_city_colors(n):
        """N визуально различимых цветов по кругу оттенков (HSL) — с фиксированной
        насыщенностью и яркостью, подобранными так, чтобы читаться и на тёмном,
        и на светлом фоне (не сливаться ни с белым, ни с чёрным текстом рядом)."""
        colors = []
        for i in range(n):
            hue = int(360 * i / n)
            colors.append(QColor.fromHsl(hue, 190, 145))
        return colors

    def _reload_edit_table(self, pairs):
        self._set_edit_rows(pairs)

    def _all_edit_pairs(self):
        t = self.edit_table
        out = []
        for r in range(t.rowCount()):
            p = t.item(r, 0).text() if t.item(r, 0) else ""
            tr = t.item(r, 1).text() if t.item(r, 1) else ""
            out.append((p, tr))
        return out

    def _update_edit_count(self):
        n = self.edit_table.rowCount()
        self.edit_count.setText("" if n == 0 else f"Осталось: {n:,}".replace(",", " "))

    def _load_edit_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть файл", "", "Excel/CSV (*.xlsx *.xlsm *.xls *.csv)")
        if path:
            self._process_edit_file(path)

    def _load_merge_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите файлы для объединения (можно сразу несколько городов, до 40)", "",
            "Excel/CSV (*.xlsx *.xlsm *.xls *.csv)")
        if not paths:
            return
        # Предел на число файлов за раз: столько выгрузок ещё нормально
        # читается и помещается в таблицу, не подвешивая интерфейс.
        MAX_MERGE_FILES = 40
        if len(paths) > MAX_MERGE_FILES:
            QMessageBox.warning(
                self, "Предупреждение",
                f"Выбрано более {MAX_MERGE_FILES} файлов. "
                f"Будут использованы первые {MAX_MERGE_FILES}.")
            paths = paths[:MAX_MERGE_FILES]
        threading.Thread(target=self._merge_worker, args=(paths,), daemon=True).start()

    # Номер телефона: 9-14 цифр, допускается «+» и код страны 7/8 впереди
    _PHONE_RE = re.compile(r"^\+?[78]?\d{9,14}$")

    @classmethod
    def _phone_ratio(cls, series):
        """Доля значений в столбце, похожих на телефонный номер (0..1)."""
        vals = [str(v).strip() for v in series if str(v).strip()]
        if not vals:
            return 0.0
        ok = 0
        for v in vals:
            if cls._PHONE_RE.match(re.sub(r"[ ()\-]", "", v)):
                ok += 1
        return ok / len(vals)

    # Слова, по которым узнаётся служебный столбец со статусом обзвона
    _STATUS_WORDS = ("дозвон", "ндз", "автоответчик", "абонент", "недоступен",
                     "занят", "сбросил", "ответил", "недозвон")

    @classmethod
    def _status_ratio(cls, series):
        """Доля значений, похожих на статус обзвона (0..1).

        Ограничение по длине важно: реплики клиента тоже иногда содержат слово
        «ответил», но они длинные, а статусы — короткие."""
        vals = [str(v).strip().lower() for v in series if str(v).strip()]
        if not vals:
            return 0.0
        ok = sum(1 for v in vals if len(v) <= 40 and any(w in v for w in cls._STATUS_WORDS))
        return ok / len(vals)

    @staticmethod
    def _is_dozvon(values):
        """Дозвон ли это, судя по всем статусным столбцам строки.

        Берём строку, если хоть один статус говорит «Дозвон» / «Абонент
        ответил», и при этом ни один не говорит «автоответчик». Недозвон
        («Ндз», «Телефон недоступен», «Не дозвонились», «Абонент занят»)
        под это правило не подходит и отсеивается."""
        positive = False
        for v in values:
            low = str(v).strip().lower()
            if not low:
                continue
            if "автоответчик" in low:
                return False
            if low == "дозвон" or "абонент ответил" in low:
                positive = True
        return positive

    @staticmethod
    def _text_score(series):
        """Насколько столбец похож на транскрибацию диалога.

        По средней длине ориентироваться нельзя: в выгрузке, где дозвонов мало,
        транскрибация заполнена у считанных строк и в среднем короче служебного
        «Телефон недоступен». Зато служебные столбцы всегда состоят из
        двух-трёх повторяющихся значений, а реплики клиента почти все разные —
        поэтому решает количество РАЗНЫХ значений. Явная метка «Клиент:»
        перевешивает всё остальное."""
        vals = [str(v).strip() for v in series if str(v).strip()]
        if not vals:
            return (0, 0, 0)
        marked = sum(1 for v in vals if "клиент" in v.lower())
        has_marker = 1 if marked >= max(1, len(vals) // 5) else 0
        return (has_marker, len(set(vals)), max(len(v) for v in vals))

    @classmethod
    def _pick_cols_by_content(cls, data):
        """Определяет по содержимому, где номер, где статусы, где транскрибация.

        Порядок важен: сначала находим номер, затем служебные столбцы со
        статусами обзвона, и только из оставшихся выбираем транскрибацию —
        иначе на роль текста легко попадает «Телефон недоступен» или
        «Неизвестно»."""
        cols = list(data.columns)
        if not cols:
            return None, None, []
        scores = {c: cls._phone_ratio(data[c]) for c in cols}
        phone_col = max(scores, key=lambda c: scores[c])
        if scores[phone_col] < 0.5:
            phone_col = cols[0]

        status_cols = [c for c in cols
                       if c != phone_col and cls._status_ratio(data[c]) >= 0.6]

        rest = [c for c in cols if c != phone_col and c not in status_cols]
        if not rest:                      # кроме номера и статусов ничего нет
            return phone_col, None, status_cols
        text_col = max(rest, key=lambda c: cls._text_score(data[c]))
        return phone_col, text_col, status_cols

    def _read_two_cols(self, path, filter_dozvon=True):
        """Читает файл и оставляет ровно два столбца: номер и что сказал клиент.

        Понимает и прежние выгрузки (с шапкой «Номер телефона» /
        «Транскрибация клиента»), и файлы без шапки, где между номером и
        диалогом стоят служебные столбцы — например:

            78423522059 | Дозвон | Активен | Абонент ответил | … | Клиент: Алло

        Наличие шапки определяется просто: если в первой строке нет ни одного
        значения, похожего на телефонный номер, — значит это заголовки.
        Проверять по словам нельзя: в файлах без шапки первая же строка
        содержит «Клиент: …», и по ключевому слову она ошибочно принималась бы
        за заголовок, а первый номер терялся."""
        if str(path).lower().endswith(".csv"):
            raw, last_err = None, None
            for enc in ("utf-8-sig", "utf-8", "cp1251"):
                try:
                    raw = pd.read_csv(path, header=None, dtype=str, encoding=enc,
                                      keep_default_na=False, sep=None, engine="python")
                    break
                except Exception as e:
                    last_err = e
            if raw is None:
                raise ValueError(f"не удалось прочитать CSV ({last_err})")
        else:
            raw = pd.read_excel(path, header=None, dtype=str)
        raw = raw.fillna("")

        if raw.empty:
            return pd.DataFrame(columns=["Номер телефона", "Транскрибация клиента"])

        first_row = [str(x).strip() for x in raw.iloc[0].tolist()]
        has_header = not any(self._PHONE_RE.match(re.sub(r"[ ()\-]", "", v))
                             for v in first_row if v)

        phone_col = trans_col = None
        if has_header:
            data = raw.iloc[1:].reset_index(drop=True)
            data.columns = [str(c).strip() for c in first_row]
            # Сначала пробуем по названиям столбцов — так надёжнее всего
            for col in data.columns:
                c = str(col).lower()
                if phone_col is None and (("номер" in c and "телефон" in c)
                                          or c in ("phone", "phone_number", "телефон")):
                    phone_col = col
                elif trans_col is None and ("транскрибация" in c or "клиент" in c
                                            or "ответ" in c or "диалог" in c
                                            or "расшифров" in c):
                    trans_col = col
        else:
            data = raw

        auto_phone, auto_text, auto_status = self._pick_cols_by_content(data)
        if phone_col is None:
            phone_col = auto_phone
        if trans_col is None:
            trans_col = auto_text

        if phone_col is None:
            return pd.DataFrame(columns=["Номер телефона", "Транскрибация клиента"])

        out = pd.DataFrame({
            "Номер телефона": data[phone_col].astype(str),
            "Транскрибация клиента": (data[trans_col].astype(str)
                                      if trans_col is not None else ""),
        })
        # Пустые строки в конце выгрузки (частый хвост у Excel) не нужны
        keep_mask = out["Номер телефона"].str.strip() != ""
        out = out[keep_mask]
        data = data[keep_mask]

        # ── оставляем только дозвон ──
        # Статусные столбцы («Дозвон/Ндз», «Абонент ответил / Ответил
        # автоответчик / Телефон недоступен») в таблицу не попадают. Если
        # filter_dozvon включён — по ним же отбираются строки: недозвон и
        # автоответчик убираются. При объединении файлов («Объединить
        # файлы») этот фильтр отключается — там статус (в т.ч. «Недозвон»)
        # должен оставаться нетронутым, ничего не удаляется по нему.
        status_cols = [c for c in auto_status if c not in (phone_col, trans_col)]
        total = len(out)
        dropped, fallback = 0, False
        if filter_dozvon and status_cols and total:
            mask = data[status_cols].apply(
                lambda row: self._is_dozvon(row.tolist()), axis=1)
            kept = int(mask.sum())
            if kept:
                out = out[mask.values]
                dropped = total - kept
            else:
                # Ни одной строки с дозвоном — показываем файл целиком, иначе
                # выглядело бы как «программа ничего не открыла».
                fallback = True

        out = out.reset_index(drop=True).fillna("")
        out.attrs["dropped_by_status"] = dropped
        out.attrs["status_fallback"] = fallback
        out.attrs["had_status"] = bool(status_cols)
        return out

    def _merge_worker(self, paths):
        try:
            self.sigEditStatus.emit(f"⏳ Объединяю {len(paths)} файла(ов)…", ACCENT)
            # filter_dozvon=False — при объединении файлов статус (в т.ч.
            # «Недозвон») не должен удаляться, оставляем строки как есть.
            all_dfs = [self._read_two_cols(p, filter_dozvon=False) for p in paths]
            dropped = sum(int(d.attrs.get("dropped_by_status", 0)) for d in all_dfs)
            fallback = sum(1 for d in all_dfs if d.attrs.get("status_fallback"))
            combined = pd.concat(all_dfs, ignore_index=True)
            pairs = [(str(row["Номер телефона"]).strip(),
                      C_clean_transcript(row["Транскрибация клиента"]))
                     for _, row in combined.iterrows()]
            self.sigEditReload.emit(pairs)
            note = ""
            if dropped:
                note += f" Отброшено без дозвона: {dropped}."
            if fallback:
                note += f" В {fallback} файле(ах) дозвона нет — показаны целиком."
            self.sigEditStatus.emit(
                f"✅ Объединено {len(paths)} файла(ов): {len(pairs)} строк." + note, GREEN)
            self._detect_edit_city(pairs)
        except Exception as e:
            self.sigError.emit("Ошибка", self._file_err("Не удалось объединить файлы:", e))
            self.sigEditStatus.emit("⚠ Ошибка объединения файлов", DUPE_RED)

    @staticmethod
    def _status_note(df):
        """Приписка к строке состояния: сколько строк отсеял фильтр дозвона."""
        dropped = int(df.attrs.get("dropped_by_status", 0))
        if dropped:
            return f" Отброшено без дозвона: {dropped}."
        if df.attrs.get("status_fallback"):
            return " Строк со статусом «Дозвон» нет — показан весь файл."
        return ""

    def _process_edit_file(self, path):
        threading.Thread(target=self._process_edit_worker, args=(path,), daemon=True).start()

    def _process_edit_worker(self, path):
        try:
            self.sigEditStatus.emit("⏳ Чтение и умная очистка файла…", ACCENT)
            df = self._read_two_cols(path)
            pairs = [(str(row["Номер телефона"]).strip(),
                      C_clean_transcript(row["Транскрибация клиента"]))
                     for _, row in df.iterrows()]
            self.sigEditReload.emit(pairs)
            self.sigEditStatus.emit(
                f"✅ Загружено {len(pairs)} строк. Лишние столбцы отсечены."
                + self._status_note(df), GREEN)
            self._detect_edit_city(pairs)
        except Exception as e:
            self.sigError.emit("Ошибка", self._file_err("Не удалось открыть файл:", e))
            self.sigEditStatus.emit("⚠ Ошибка загрузки файла", DUPE_RED)

    def _build_rev_codes(self):
        """code -> [(город, key), ...]. Один и тот же код иногда числится
        за несколькими городами (например, у более мелкого города в области
        код совпадает с префиксом кода более крупного) — поэтому здесь
        список кандидатов, а не один город."""
        if getattr(self, "_rev_codes", None):
            return self._rev_codes
        rev = {}
        for key, code in (self.ph_codes or {}).items():
            if not code:
                continue
            city = key.split("|")[0]
            rev.setdefault(str(code), []).append((city, key))
        # ВАЖНО: кешируем только непустой результат. Если self.ph_codes ещё
        # не загружен (кто-то вызвал определение города раньше, чем
        # отработал _ensure_data), rev получится пустым — раньше это пустое
        # значение всё равно кешировалось навсегда, и после этого определение
        # города/оператора переставало работать во ВСЕЙ программе до
        # перезапуска, даже после того как база всё-таки загружалась.
        if rev:
            self._rev_codes = rev
        return rev

    def _lookup_phone_city_operator(self, raw_phone):
        """Определяет (город, оператор) для номера по нашей же базе диапазонов
        (та же, что использует вкладка «Города») — без обращения к сторонним
        сайтам. При совпадении кода сразу у нескольких городов сверяемся с
        настоящими диапазонами номеров, чтобы не перепутать город (например,
        Ростов и Морозовск, у которых код может пересекаться по префиксу)."""
        # Подстраховка: гарантируем, что база городов/номеров уже загружена,
        # независимо от того, откуда вызвана эта функция — раньше это было
        # обязанностью каждого вызывающего места по отдельности, и один
        # пропущенный вызов ломал определение города везде до перезапуска.
        try:
            self._ensure_data()
        except Exception:
            pass
        digits = re.sub(r'\D', '', str(raw_phone))
        if digits.startswith('7') or digits.startswith('8'):
            digits = digits[1:]
        if len(digits) < 6:
            return "", ""
        try:
            num = int(digits)
        except ValueError:
            return "", ""

        rev = self._build_rev_codes()
        for ln in (5, 4, 3):
            candidates = rev.get(digits[:ln])
            if not candidates:
                continue
            if len(candidates) > 1:
                for city, key in candidates:
                    for rng in (self.ph_ranges or {}).get(key, []):
                        frm, span, cap, opidx, kind = rng
                        if frm <= num <= frm + span:
                            op = self.ph_ops[opidx] if 0 <= opidx < len(self.ph_ops) else ""
                            return city, op
                # ни один диапазон не подошёл точно — берём первого кандидата
                # как запасной вариант (город без гарантии оператора)
                return candidates[0][0], ""
            city, key = candidates[0]
            for rng in (self.ph_ranges or {}).get(key, []):
                frm, span, cap, opidx, kind = rng
                if frm <= num <= frm + span:
                    op = self.ph_ops[opidx] if 0 <= opidx < len(self.ph_ops) else ""
                    return city, op
            return city, ""
        return "", ""

    def _detect_edit_city(self, pairs):
        try:
            self._ensure_data()
            city_counts = {}
            for phone, _ in pairs:
                city, _op = self._lookup_phone_city_operator(phone)
                if city:
                    city_counts[city] = city_counts.get(city, 0) + 1
            self._edit_city_counts = city_counts
            if city_counts:
                total = sum(city_counts.values())
                ranked = sorted(city_counts.items(), key=lambda kv: -kv[1])
                top_city = ranked[0][0]
                self._edit_detected_city = top_city
                top_pct = int(ranked[0][1] / total * 100) if total else 0
                if len(ranked) == 1 or top_pct >= 90:
                    # один город (или он явно преобладает) — один чип, без %,
                    # если он практически единственный
                    chips = [(top_city, None if top_pct >= 90 else top_pct)]
                else:
                    # смесь городов — чип на КАЖДЫЙ (не только топ-3, чипы сами
                    # переносятся на следующую строку при нехватке места)
                    chips = [(c, int(cnt / total * 100)) for c, cnt in ranked]
            else:
                self._edit_detected_city = ""
                self._edit_city_counts = {}
                chips = []
            self.sigEditCity.emit(chips)
        except Exception:
            pass

    def _push_undo(self):
        self._edit_undo.append(self._all_edit_pairs())
        if len(self._edit_undo) > 50:
            self._edit_undo.pop(0)

    def _del_edit_rows(self):
        rows = sorted({i.row() for i in self.edit_table.selectedIndexes()}, reverse=True)
        if not rows:
            self._set_label(self.edit_status, "Выделите строки, затем нажмите Delete.", FG2)
            return
        self._push_undo()
        for r in rows:
            self.edit_table.removeRow(r)
        self._set_label(self.edit_status, f"🗑 Удалено строк: {len(rows)}", DUPE_RED)
        self._update_edit_count()

    def _clear_edit_table(self):
        n = self.edit_table.rowCount()
        if n == 0:
            return
        self._push_undo()
        self.edit_table.setRowCount(0)
        self._set_label(self.edit_status, f"🗑 Список очищен ({n} строк).", FG2)
        self._update_edit_count()
        self._set_edit_city_badge([])
        self._edit_detected_city = ""
        self._edit_city_counts = {}

    # Список слов для поиска в столбце A — тот же, что и в исходном Excel-макросе.
    _MACRO_KEYWORDS = ("2GIS", "Yell.ru", "1cont.ru", "CDEK", "Miltor", "Suppliers",
                       "RusProfile.ru", "Bureau van Dijk", "RuZakaz", "Checko", "MegaMarket")

    def _run_company_filter_macro(self):
        """Портированный Excel-макрос RemoveRowsByKeywords + дополнительная
        пост-обработка, применяется к каждому выбранному файлу (от 1 до 20)
        по отдельности: убирает строки, где в столбце A встречается одно из
        ключевых слов (названия организаций/агрегаторов), а также все прочие
        строки, где в столбце B встречается тот же номер, что и в найденной
        строке. Затем из оставшихся строк берёт только столбец «request»
        (номер телефона), убирает заголовок, схлопывает дубликаты номеров и
        записывает в B1 итоговое количество уникальных номеров."""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Открыть Excel-файлы (можно несколько, до 20)", "",
            "Excel файлы (*.xlsx *.xls);;Все файлы (*.*)")
        if not paths:
            return
        if len(paths) > 20:
            QMessageBox.warning(
                self, "Инфо",
                f"Выбрано {len(paths)} файлов — за один раз можно обработать не больше 20. "
                "Возьми первые 20 из выбранных.")
            paths = paths[:20]

        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        dated_dir = os.path.join(
            desktop_dir, f"готово к загрузке {datetime.date.today().strftime('%d.%m.%Y')}")
        try:
            os.makedirs(dated_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать папку:\n{dated_dir}\n{e}")
            return

        results = []  # (имя_файла, ok, детали_или_ошибка)
        for path in paths:
            try:
                stats, save_path = self._apply_company_filter_macro_to_file(path, dated_dir)
                results.append((os.path.basename(path), True, stats, save_path))
            except Exception as e:
                results.append((os.path.basename(path), False, str(e), None))

        ok_count = sum(1 for r in results if r[1])
        lines = [f"Обработано файлов: {ok_count} из {len(results)}", f"Папка: {dated_dir}", ""]
        for name, ok, info, save_path in results:
            if ok:
                lines.append(
                    f"✅ {name}: было {info['initial']}, удалено {info['deleted_kw']}+{info['deleted_num']}, "
                    f"уникальных номеров {info['unique']} → {os.path.basename(save_path)}")
            else:
                lines.append(f"⚠ {name}: ошибка — {info}")

        QMessageBox.information(self, "Готово", "\n".join(lines))

    def _apply_company_filter_macro_to_file(self, path, dated_dir):
        """Обрабатывает один файл макросом и сохраняет результат в dated_dir.
        Возвращает (словарь со статистикой, путь к сохранённому файлу)."""
        df = pd.read_excel(path, header=None, dtype=str).fillna("")

        if df.shape[0] == 0:
            raise ValueError("файл пуст")
        if df.shape[1] < 2:
            raise ValueError("меньше двух столбцов — макросу нужны столбцы A и B")

        initial_count = len(df)
        col_a = df.iloc[:, 0].astype(str)
        col_b = df.iloc[:, 1].astype(str)

        # Первый проход: строки с ключевыми словами в столбце A —
        # запоминаем их номера (столбец B) и сразу помечаем на удаление.
        kw_mask = col_a.apply(
            lambda v: any(kw.lower() in v.lower() for kw in self._MACRO_KEYWORDS))
        flagged_numbers = set(n for n in col_b[kw_mask] if n.strip())

        # Второй проход: любая строка, где столбец B совпадает с одним
        # из найденных номеров, — тоже удаляется.
        dup_mask = col_b.isin(flagged_numbers)

        drop_mask = kw_mask | dup_mask
        deleted_by_keyword = int(kw_mask.sum())
        deleted_by_number = int((dup_mask & ~kw_mask).sum())

        result_df = df[~drop_mask]

        # Оставляем только столбец «request» (номер телефона, столбец B),
        # убираем всё остальное, убираем сам заголовок «request» и
        # схлопываем дубликаты номеров, оставляя одно вхождение каждого.
        numbers = result_df.iloc[:, 1].astype(str).str.strip()
        numbers = numbers[numbers.str.match(r"^\+?\d+$")]  # только цифры — убирает и шапку «request»
        unique_numbers = list(dict.fromkeys(numbers))  # дедуп с сохранением порядка

        col_b_out = [""] * len(unique_numbers)
        if unique_numbers:
            col_b_out[0] = len(unique_numbers)
        out_df = pd.DataFrame({0: unique_numbers, 1: col_b_out})

        base_name = os.path.splitext(os.path.basename(path))[0] + "_очищено.xlsx"
        save_path = os.path.join(dated_dir, base_name)
        # Если файл с таким именем уже есть — не перезаписываем, а добавляем номер.
        n = 2
        stem, ext = os.path.splitext(save_path)
        while os.path.exists(save_path):
            save_path = f"{stem} ({n}){ext}"
            n += 1

        out_df.to_excel(save_path, header=False, index=False)

        stats = {
            "initial": initial_count,
            "deleted_kw": deleted_by_keyword,
            "deleted_num": deleted_by_number,
            "remaining": len(result_df),
            "unique": len(unique_numbers),
        }
        return stats, save_path

    # ───────────────  Кнопка «Статусы»: разбор номеров по статусу обзвона  ───────────────
    def _run_statuses_tool(self):
        """Загружает одну или сразу несколько CSV/Excel-выгрузок со статусами
        обзвона (номер, статус — остальные столбцы вроде имени, региона,
        паспорта и т.п. не нужны и отбрасываются), определяет оператора по
        номеру нашей же базой диапазонов и открывает ОДИН общий диалог для
        всех файлов сразу, где можно отметить, какие статусы оставить.
        Сохраняется при этом каждый файл отдельно — в свой файл."""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Открыть файл(ы) со статусами обзвона (можно несколько)", "",
            "Таблицы (*.csv *.xlsx *.xls);;CSV (*.csv);;Excel (*.xlsx *.xls);;Все файлы (*.*)")
        if not paths:
            return

        frames = []
        errors = []
        for path in paths:
            try:
                df = self._load_statuses_file(path)
                if df.empty:
                    errors.append((os.path.basename(path), "нет строк с номером и статусом"))
                    continue
                df["__source"] = path
                frames.append(df)
            except Exception as e:
                errors.append((os.path.basename(path), str(e)))

        if errors:
            lines = [f"⚠ {name}: {err}" for name, err in errors]
            QMessageBox.warning(
                self, "Часть файлов не прочиталась",
                "Эти файлы пропущены:\n\n" + "\n".join(lines) +
                ("\n\nОстальные файлы обработаются нормально." if frames else ""))
        if not frames:
            return

        combined = pd.concat(frames, ignore_index=True)
        self._show_statuses_dialog(combined, paths)

    def _load_statuses_file(self, path):
        """Читает CSV или Excel-выгрузку, оставляет только номер, статус и
        (вычисленного) оператора — всё остальное (имя, регион, паспорт,
        telegram и т.д.) для этой задачи не нужно и отбрасывается."""
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            raw = None
            last_err = None
            # Экспорт может быть не строго в UTF-8 — пробуем по очереди
            # самые частые варианты кодировки, пока одна не сработает.
            for enc in ("utf-8-sig", "utf-8", "cp1251"):
                try:
                    raw = pd.read_csv(path, encoding=enc, dtype=str)
                    break
                except Exception as e:
                    last_err = e
                    continue
            if raw is None:
                raise ValueError(f"не удалось определить кодировку файла ({last_err})")
        else:
            raw = pd.read_excel(path, dtype=str)

        raw.columns = [str(c).strip().lower() for c in raw.columns]
        if "phone_number" not in raw.columns or "status" not in raw.columns:
            raise ValueError(
                "в файле должны быть столбцы «phone_number» и «status» "
                "(это формат выгрузки статусов обзвона)")

        out = raw[["phone_number", "status"]].copy()
        out.columns = ["number", "status"]
        out["number"] = out["number"].astype(str).str.strip()
        out["status"] = out["status"].astype(str).str.strip()
        out = out[out["number"].str.match(r"^\+?\d{6,15}$")]
        out["number"] = out["number"].str.lstrip("+")  # сохраняем без «+» впереди
        out["operator"] = out["number"].apply(lambda n: self._lookup_phone_city_operator(n)[1] or "—")
        return out.reset_index(drop=True)

    def _show_statuses_dialog(self, df, src_paths):
        dlg = QDialog(self)
        dlg.setWindowTitle("Статусы обзвона")
        dlg.setMinimumWidth(460)
        lay = QVBoxLayout(dlg)
        lay.setSpacing(10)

        total = len(df)
        n_files = df["__source"].nunique() if "__source" in df.columns else len(src_paths)
        op_counts = df["operator"].value_counts()
        op_line = ", ".join(f"{op}: {cnt:,}".replace(",", " ") for op, cnt in op_counts.items())
        info = QLabel(
            f"Загружено файлов: {n_files}, номеров: {total:,}\nПо операторам: {op_line}".replace(",", " "))
        info.setStyleSheet(f"color:{FG2};")
        info.setWordWrap(True)
        lay.addWidget(info)

        hint = QLabel("Отметь статусы, которые нужно ОСТАВИТЬ — остальные будут удалены:")
        hint.setStyleSheet(f"color:{FG}; font-weight:600;")
        lay.addWidget(hint)

        status_counts = df["status"].value_counts()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setSpacing(4)
        checks = {}
        for status, cnt in status_counts.items():
            cb = QCheckBox(f"{status}   ({cnt:,})".replace(",", " "))
            cb.setChecked(True)  # по умолчанию оставляем все статусы — пользователь сам снимает лишнее
            checks[status] = cb
            inner_lay.addWidget(cb)
        inner_lay.addStretch()
        scroll.setWidget(inner)
        scroll.setMaximumHeight(280)
        lay.addWidget(scroll)

        quick_bar = QHBoxLayout()
        select_all = QPushButton("Выбрать всё")
        select_all.clicked.connect(lambda: [cb.setChecked(True) for cb in checks.values()])
        select_none = QPushButton("Снять всё")
        select_none.clicked.connect(lambda: [cb.setChecked(False) for cb in checks.values()])
        quick_bar.addWidget(select_all)
        quick_bar.addWidget(select_none)
        quick_bar.addStretch()
        lay.addLayout(quick_bar)

        # Маленькие файлы (по 40-60 номеров на файл, если статусов после
        # фильтра осталось мало) можно объединить в один пул и нарезать
        # заново по нужному количеству номеров на файл — как в «Генерации».
        # Маленькие файлы (по 40-60 номеров на файл, если статусов после
        # фильтра осталось мало) можно объединить в один пул и нарезать
        # заново по нужному количеству номеров на файл — как в «Генерации».
        merge_row = QHBoxLayout()
        merge_cb = QCheckBox("  Объединить все файлы в один пул и нарезать по:")
        merge_cb.setIcon(_icon(ICONS["shuffle"], FG, ""))
        merge_cb.setIconSize(QSize(15, 15))
        chunk_edit = QLineEdit("5000")
        chunk_edit.setFixedWidth(90)
        chunk_edit.setEnabled(False)
        merge_cb.toggled.connect(chunk_edit.setEnabled)
        merge_row.addWidget(merge_cb)
        merge_row.addWidget(chunk_edit)
        merge_row.addWidget(QLabel("номеров в файле"))
        merge_row.addStretch()
        lay.addLayout(merge_row)

        preview = QLabel("")
        preview.setStyleSheet(f"color:{FG2};")
        lay.addWidget(preview)

        def update_preview():
            keep = {s for s, cb in checks.items() if cb.isChecked()}
            n = df[df["status"].isin(keep)]["number"].nunique()
            if merge_cb.isChecked():
                try:
                    chunk = max(1, int(chunk_edit.text().strip()))
                except ValueError:
                    chunk = 0
                files_n = (n + chunk - 1) // chunk if chunk else 0
                preview.setText(
                    f"Будет сохранено уникальных номеров: {n:,} → файлов: {files_n:,}".replace(",", " "))
            else:
                preview.setText(
                    f"Будет сохранено уникальных номеров: {n:,} (по одному файлу на исходный файл)".replace(",", " "))

        for cb in checks.values():
            cb.toggled.connect(update_preview)
        merge_cb.toggled.connect(update_preview)
        chunk_edit.textChanged.connect(update_preview)
        update_preview()

        bar = QHBoxLayout()
        save_btn = QPushButton("  Сохранить выбранные")
        save_btn.setIcon(_icon(ICONS["save"], BG, ""))
        save_btn.setObjectName("green")
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(dlg.reject)
        bar.addStretch()
        bar.addWidget(cancel_btn)
        bar.addWidget(save_btn)
        lay.addLayout(bar)

        def do_save():
            keep_statuses = {s for s, cb in checks.items() if cb.isChecked()}
            if not keep_statuses:
                QMessageBox.warning(dlg, "Инфо", "Не выбрано ни одного статуса.")
                return
            filtered = df[df["status"].isin(keep_statuses)]
            if filtered.empty:
                QMessageBox.warning(dlg, "Инфо", "После фильтра не осталось ни одного номера.")
                return
            if merge_cb.isChecked():
                try:
                    chunk = int(chunk_edit.text().strip())
                    if chunk <= 0:
                        raise ValueError
                except ValueError:
                    QMessageBox.warning(dlg, "Ошибка", "«Номеров в файле» должно быть целым числом больше нуля.")
                    return
                try:
                    results = self._save_statuses_merged(filtered, chunk)
                except Exception as e:
                    QMessageBox.critical(dlg, "Ошибка", f"Не удалось сохранить:\n{e}")
                    return
            else:
                try:
                    results = self._save_statuses_per_file(filtered)
                except Exception as e:
                    QMessageBox.critical(dlg, "Ошибка", f"Не удалось сохранить:\n{e}")
                    return
            dlg.accept()
            total_numbers = sum(r[1] for r in results)
            lines = [f"Сохранено файлов: {len(results)}", f"Всего номеров: {total_numbers:,}".replace(",", " "), ""]
            for name, cnt, save_path in results:
                lines.append(f"✅ {name}: {cnt:,} номеров → {os.path.basename(save_path)}".replace(",", " "))
            QMessageBox.information(self, "Готово!", "\n".join(lines))

        save_btn.clicked.connect(do_save)
        dlg.exec()

    @staticmethod
    def _write_numbers_csv(numbers, save_path):
        """Пишет «голый» столбик номеров: одна строка — один номер.

        Без заголовка, без статуса, без счётчика и без BOM — макросу нужны
        только цифры. Пишем напрямую, а не через pandas: так гарантированно
        не появятся кавычки и лишние разделители."""
        with open(save_path, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(str(n).strip() for n in numbers))
            f.write("\n")

    def _save_statuses_merged(self, filtered_df, chunk_size):
        """Объединяет номера из ВСЕХ загруженных файлов в один пул (глобальный
        дедуп — не по каждому файлу отдельно), затем нарезает по chunk_size
        номеров на файл — так же, как это делает вкладка «Генерация».
        Файлы кладутся в подпапку с названием города — так же, как при
        раздельном сохранении (см. _save_statuses_per_file). Город берётся
        из имени файла; если объединённые файлы оказались из РАЗНЫХ городов,
        подпапку не создаём — кладём прямо в датированную папку, чтобы не
        приписывать чужой город чужим номерам."""
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        dated_dir = os.path.join(
            desktop_dir, f"чистые статусы {datetime.date.today().strftime('%d.%m.%Y')}")

        cities = {C.get_city_from_source(p) for p in filtered_df["__source"].unique()}
        out_dir = os.path.join(dated_dir, next(iter(cities))) if len(cities) == 1 else dated_dir
        os.makedirs(out_dir, exist_ok=True)

        numbers = list(dict.fromkeys(filtered_df["number"]))  # дедуп по всем файлам сразу
        if not numbers:
            return []

        results = []
        for i in range(0, len(numbers), chunk_size):
            part = numbers[i:i + chunk_size]
            if not part:
                break
            idx = i // chunk_size + 1
            col_b = [""] * len(part)
            col_b[0] = len(part)
            out_df = pd.DataFrame({0: part, 1: col_b})

            save_path = os.path.join(out_dir, f"статусы_{idx:03d}.xlsx")
            n = 2
            stem, ext = os.path.splitext(save_path)
            while os.path.exists(save_path):
                save_path = f"{stem} ({n}){ext}"
                n += 1

            out_df.to_excel(save_path, header=False, index=False)
            results.append((os.path.basename(save_path), len(part), save_path))
        return results

    def _save_statuses_per_file(self, filtered_df):
        """Сохраняет каждый исходный файл ОТДЕЛЬНО (не объединяя в один) —
        группируем уже отфильтрованные строки обратно по исходному файлу
        (столбец __source, проставленный при загрузке) и для каждого
        вызываем тот же _save_statuses_result. Дубли схлопываются в
        пределах каждого файла отдельно, как и у кнопки «Макрос»."""
        results = []
        for src_path, group in filtered_df.groupby("__source", sort=False):
            numbers = list(dict.fromkeys(group["number"]))  # дедуп с сохранением порядка
            if not numbers:
                continue
            save_path = self._save_statuses_result(numbers, src_path)
            results.append((os.path.basename(src_path), len(numbers), save_path))
        return results

    def _save_statuses_result(self, numbers, src_path):
        """Сохраняет отфильтрованные номера тем же способом, что и кнопка
        «Макрос»: столбец A — номера в столбик (без «+» впереди), столбец B —
        только в первой строке итоговое количество уникальных номеров.
        Папка создаётся на рабочем столе с сегодняшней датой в названии, а
        внутри неё — подпапка с названием города, определённым по имени
        исходного файла (та же логика, что и в остальной программе, см.
        get_city_from_source — например, «09_06_Мурманск_дом_export.csv»
        даёт город «Мурманск»)."""
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        dated_dir = os.path.join(
            desktop_dir, f"чистые статусы {datetime.date.today().strftime('%d.%m.%Y')}")
        city = C.get_city_from_source(src_path)
        dated_dir = os.path.join(dated_dir, city)
        os.makedirs(dated_dir, exist_ok=True)

        col_b = [""] * len(numbers)
        col_b[0] = len(numbers)
        out_df = pd.DataFrame({0: numbers, 1: col_b})

        base_name = os.path.splitext(os.path.basename(src_path))[0] + "_статусы.xlsx"
        save_path = os.path.join(dated_dir, base_name)
        n = 2
        stem, ext = os.path.splitext(save_path)
        while os.path.exists(save_path):
            save_path = f"{stem} ({n}){ext}"
            n += 1

        out_df.to_excel(save_path, header=False, index=False)
        return save_path

    # ═══════════════════  ДОЗВОН (выгрузки прозвона)  ═══════════════════
    # Отдельный инструмент, не связанный с кнопкой «Статусы»: там свой формат
    # файлов (phone_number/status) и свой сценарий, и смешивать их в одной
    # кнопке оказалось неудобно.
    def _run_calls_tool(self):
        """Загружает выгрузки прозвона (Phone + Result/Result+), даёт убрать
        недозвон и автоответчик и сохраняет оставшиеся номера — по одному
        файлу на город или по файлу на каждую загруженную выгрузку."""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Открыть выгрузки прозвона (можно сразу несколько городов)", "",
            "Таблицы (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv);;Все файлы (*.*)")
        if not paths:
            return

        frames, errors = [], []
        for path in paths:
            try:
                df = self._load_calls_file(path)
                if df.empty:
                    errors.append((os.path.basename(path), "нет строк с номером и статусом"))
                    continue
                df["__source"] = path
                frames.append(df)
            except Exception as e:
                errors.append((os.path.basename(path), str(e)))

        if errors:
            lines = [f"⚠ {name}: {err}" for name, err in errors]
            QMessageBox.warning(
                self, "Часть файлов не прочиталась",
                "Эти файлы пропущены:\n\n" + "\n".join(lines) +
                ("\n\nОстальные файлы обработаются нормально." if frames else ""))
        if not frames:
            return

        self._show_calls_dialog(pd.concat(frames, ignore_index=True))

    def _load_calls_file(self, path):
        """Читает выгрузку прозвона: номер + статус, остальное отбрасывается.

        Статус берём из Result+, если он есть: там разбивка подробнее
        («Недоступен/Автоответчик», «Абонент не ответил», «Дозвон»), тогда как
        в Result всё сведено к паре «Дозвон / Недозвон». «Дозвон» в обоих
        столбцах называется одинаково, поэтому фильтр работает в любом случае.
        """
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            raw, last_err = None, None
            for enc in ("utf-8-sig", "utf-8", "cp1251"):
                try:
                    raw = pd.read_csv(path, encoding=enc, dtype=str)
                    break
                except Exception as e:
                    last_err = e
            if raw is None:
                raise ValueError(f"не удалось определить кодировку файла ({last_err})")
        else:
            raw = pd.read_excel(path, dtype=str)

        raw.columns = [str(c).strip().lower() for c in raw.columns]
        cols = set(raw.columns)
        if "phone" in cols and ("result+" in cols or "result" in cols):
            num_col = "phone"
            st_col = "result+" if "result+" in cols else "result"
        elif "phone_number" in cols and "status" in cols:
            num_col, st_col = "phone_number", "status"
        else:
            raise ValueError(
                "не найдены нужные столбцы — нужны «Phone» и «Result» (или «Result+»)")

        out = raw[[num_col, st_col]].copy()
        out.columns = ["number", "status"]
        out["number"] = out["number"].astype(str).str.strip()
        out["status"] = out["status"].astype(str).str.strip()
        out = out[out["number"].str.match(r"^\+?\d{6,15}$")]
        out["number"] = out["number"].str.lstrip("+")   # номера без «+» впереди
        return out.reset_index(drop=True)

    def _show_calls_dialog(self, df):
        dlg = QDialog(self)
        dlg.setWindowTitle("Дозвон")
        dlg.setMinimumWidth(520)
        lay = QVBoxLayout(dlg)
        lay.setSpacing(10)

        sources = list(dict.fromkeys(df["__source"].tolist()))
        cities = [C.get_city_from_source(p) for p in sources]
        info = QLabel(
            f"Загружено файлов: {len(sources)}, номеров: {len(df):,}\n"
            f"Города: {', '.join(sorted(set(cities)))}".replace(",", " ", 1))
        info.setStyleSheet(f"color:{FG2};")
        info.setWordWrap(True)
        lay.addWidget(info)

        hint = QLabel("Отметь статусы, которые нужно ОСТАВИТЬ — остальные будут удалены:")
        hint.setStyleSheet(f"color:{FG}; font-weight:600;")
        lay.addWidget(hint)

        status_counts = df["status"].value_counts()
        # По умолчанию отмечен только «Дозвон»: недозвон и автоответчик в этом
        # сценарии всегда идут под снос. Если такого статуса в файле нет —
        # отмечаем всё, чтобы ничего молча не потерялось.
        has_dozvon = any(str(st).strip().lower() == "дозвон" for st in status_counts.index)
        checks = {}
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setSpacing(4)
        for status, cnt in status_counts.items():
            cb = QCheckBox(f"{status}   ({cnt:,})".replace(",", " "))
            cb.setChecked(str(status).strip().lower() == "дозвон" if has_dozvon else True)
            checks[status] = cb
            inner_lay.addWidget(cb)
        inner_lay.addStretch()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        scroll.setMaximumHeight(200)
        lay.addWidget(scroll)

        cap = QLabel("Как сохранять:")
        cap.setStyleSheet(f"color:{FG}; font-weight:600;")
        lay.addWidget(cap)
        rb_city = QRadioButton("  Один файл на город (куски одного города склеиваются)")
        rb_city.setIcon(_icon(ICONS["map"], FG, ""))
        rb_city.setIconSize(QSize(15, 15))
        rb_each = QRadioButton("  Каждый файл отдельно")
        rb_each.setIcon(_icon(ICONS["layers"], FG, ""))
        rb_each.setIconSize(QSize(15, 15))
        grp = QButtonGroup(dlg)
        grp.addButton(rb_city)
        grp.addButton(rb_each)
        (rb_each if getattr(self, "_calls_group_mode", "city") == "each" else rb_city).setChecked(True)
        lay.addWidget(rb_city)
        lay.addWidget(rb_each)

        fmt_row = QHBoxLayout()
        fmt_row.addWidget(QLabel("Формат:"))
        rb_csv = QRadioButton("CSV — только номера (под макрос)")
        rb_xlsx = QRadioButton("Excel — номера + количество")
        fmt_grp = QButtonGroup(dlg)
        fmt_grp.addButton(rb_csv)
        fmt_grp.addButton(rb_xlsx)
        (rb_csv if getattr(self, "_calls_csv_mode", True) else rb_xlsx).setChecked(True)
        fmt_row.addWidget(rb_csv)
        fmt_row.addWidget(rb_xlsx)
        fmt_row.addStretch()
        lay.addLayout(fmt_row)

        preview = QLabel("")
        preview.setStyleSheet(f"color:{FG2};")
        preview.setWordWrap(True)
        lay.addWidget(preview)

        def update_preview():
            keep = {st for st, cb in checks.items() if cb.isChecked()}
            sub = df[df["status"].isin(keep)]
            n = sub["number"].nunique()
            ext = "csv" if rb_csv.isChecked() else "xlsx"
            head = f"Будет сохранено уникальных номеров: {n:,}".replace(",", " ")
            if rb_city.isChecked():
                cities_n = len({C.get_city_from_source(p) for p in sub["__source"].unique()}) if not sub.empty else 0
                preview.setText(f"{head} → городов: {cities_n}, по одному .{ext} на город")
            else:
                preview.setText(f"{head} → файлов: {sub['__source'].nunique()} (по одному .{ext} на выгрузку)")

        for cb in checks.values():
            cb.toggled.connect(update_preview)
        for rb in (rb_city, rb_each, rb_csv, rb_xlsx):
            rb.toggled.connect(update_preview)
        update_preview()

        bar = QHBoxLayout()
        save_btn = QPushButton("  Сохранить")
        save_btn.setIcon(_icon(ICONS["save"], BG, ""))
        save_btn.setObjectName("green")
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(dlg.reject)
        bar.addStretch()
        bar.addWidget(cancel_btn)
        bar.addWidget(save_btn)
        lay.addLayout(bar)

        def do_save():
            keep = {st for st, cb in checks.items() if cb.isChecked()}
            if not keep:
                QMessageBox.warning(dlg, "Инфо", "Не выбрано ни одного статуса.")
                return
            filtered = df[df["status"].isin(keep)]
            if filtered.empty:
                QMessageBox.warning(dlg, "Инфо", "После фильтра не осталось ни одного номера.")
                return
            as_csv = rb_csv.isChecked()
            by_city = rb_city.isChecked()
            self._calls_csv_mode = as_csv
            self._calls_group_mode = "city" if by_city else "each"
            try:
                results = (self._save_calls_by_city(filtered, as_csv) if by_city
                           else self._save_calls_per_file(filtered, as_csv))
            except Exception as e:
                QMessageBox.critical(dlg, "Ошибка", self._file_err("Не удалось сохранить:", e))
                return
            dlg.accept()
            total = sum(r[1] for r in results)
            lines = [f"Сохранено файлов: {len(results)}",
                     f"Всего номеров: {total:,}".replace(",", " "), ""]
            for name, cnt, path in results:
                lines.append(f"✅ {name}: {cnt:,} номеров → {os.path.basename(path)}".replace(",", " "))
            if results:
                lines += ["", f"Папка: {os.path.dirname(results[0][2])}"]
            QMessageBox.information(self, "Готово!", "\n".join(lines))
            self.toast("Дозвон готов", f"Файлов: {len(results)}, номеров: {total}", "ok")

        save_btn.clicked.connect(do_save)
        dlg.exec()

    def _calls_out_dir(self):
        out = os.path.join(os.path.expanduser("~"), "Desktop",
                           f"дозвон {datetime.date.today().strftime('%d.%m.%Y')}")
        os.makedirs(out, exist_ok=True)
        return out

    @staticmethod
    def _safe_file_name(name):
        return "".join(c for c in str(name) if c.isalnum() or c in " _-").strip() or "город"

    def _write_calls_file(self, numbers, save_path, as_csv):
        """CSV — голый столбик номеров (без заголовка, статуса и счётчика),
        Excel — как в «Статусах»: номера в столбце A, количество в B1."""
        n = 2
        stem, ext = os.path.splitext(save_path)
        while os.path.exists(save_path):
            save_path = f"{stem} ({n}){ext}"
            n += 1
        if as_csv:
            self._write_numbers_csv(numbers, save_path)
        else:
            col_b = [""] * len(numbers)
            col_b[0] = len(numbers)
            pd.DataFrame({0: numbers, 1: col_b}).to_excel(save_path, header=False, index=False)
        return save_path

    def _save_calls_by_city(self, filtered_df, as_csv=False):
        """Склеивает все выгрузки одного города в ОДИН файл.

        Город берётся из имени исходного файла, поэтому «Обнинск 02 3к»,
        «Обнинск 02 7к» и «Обнинск 03 3к» сходятся в один «Обнинск». Дубли
        схлопываются в пределах города, города между собой не смешиваются."""
        out_dir = self._calls_out_dir()
        by_city = {}
        for src_path, group in filtered_df.groupby("__source", sort=False):
            city = C.get_city_from_source(src_path) or "без города"
            by_city.setdefault(city, []).extend(group["number"].tolist())

        results = []
        for city, numbers in by_city.items():
            numbers = list(dict.fromkeys(numbers))
            if not numbers:
                continue
            name = f"{self._safe_file_name(city)}_дозвон." + ("csv" if as_csv else "xlsx")
            path = self._write_calls_file(numbers, os.path.join(out_dir, name), as_csv)
            results.append((city, len(numbers), path))
        return results

    def _save_calls_per_file(self, filtered_df, as_csv=False):
        """Каждая загруженная выгрузка сохраняется в свой файл, в подпапку
        своего города."""
        out_dir = self._calls_out_dir()
        results = []
        for src_path, group in filtered_df.groupby("__source", sort=False):
            numbers = list(dict.fromkeys(group["number"].tolist()))
            if not numbers:
                continue
            city = C.get_city_from_source(src_path) or "без города"
            city_dir = os.path.join(out_dir, self._safe_file_name(city))
            os.makedirs(city_dir, exist_ok=True)
            base = os.path.splitext(os.path.basename(src_path))[0] + "_дозвон." + ("csv" if as_csv else "xlsx")
            path = self._write_calls_file(numbers, os.path.join(city_dir, base), as_csv)
            results.append((os.path.basename(src_path), len(numbers), path))
        return results

    def _undo_edit(self):
        if not self._edit_undo:
            self._set_label(self.edit_status, "↩ Нечего отменять.", FG2)
            return
        pairs = self._edit_undo.pop()
        self._set_edit_rows(pairs)
        self._set_label(self.edit_status, f"↩ Восстановлено строк: {len(pairs)}", ACCENT)

    def _copy_edit_numbers(self):
        rows = sorted({i.row() for i in self.edit_table.selectedIndexes()})
        if not rows:
            rows = range(self.edit_table.rowCount())
        nums = [self.edit_table.item(r, 0).text() for r in rows if self.edit_table.item(r, 0)]
        if not nums:
            return
        QApplication.clipboard().setText("\n".join(nums))
        self._set_label(self.edit_status, f"📋 Скопировано номеров: {len(nums)}", ACCENT)

    def _on_edit_header_dbl(self, col):
        if col == 0:
            self._copy_edit_numbers_all()

    def _copy_edit_numbers_all(self):
        nums = [self.edit_table.item(r, 0).text()
                for r in range(self.edit_table.rowCount()) if self.edit_table.item(r, 0)]
        if nums:
            QApplication.clipboard().setText("\n".join(nums))
            self._set_label(self.edit_status, f"📋 Скопированы все номера: {len(nums)}", GREEN)

    def _on_edit_dbl(self, r, c):
        if c == 0:
            num = self.edit_table.item(r, 0).text()
            QApplication.clipboard().setText(num)
            self._set_label(self.edit_status, f"📋 Номер скопирован: {num}", ACCENT)
        else:
            phone = self.edit_table.item(r, 0).text()
            trans = self.edit_table.item(r, 1).text() if self.edit_table.item(r, 1) else ""
            self._show_edit_detail(phone, trans)

    def _show_edit_detail(self, phone, trans):
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{phone} — диалог")
        dlg.resize(640, 460)
        dlg.setStyleSheet(build_qss())
        lay = QVBoxLayout(dlg)
        head = QLabel(f"☎  {phone}")
        head.setStyleSheet(f"color:{ACCENT}; font-weight:bold; font-size:13pt;")
        lay.addWidget(head)
        box = QTextEdit()
        box.setReadOnly(True)
        box.setPlainText(trans if trans else "— пусто —")
        lay.addWidget(box, 1)
        bar = QHBoxLayout()
        cp = QPushButton(" Копировать диалог")
        cp.setIcon(_icon(ICONS["copy"], FG, ""))
        cp.setIconSize(QSize(15, 15))
        cp.clicked.connect(lambda: (QApplication.clipboard().setText(trans),
                                     cp.setText("✓ Скопировано")))
        cl = QPushButton("Закрыть")
        cl.setObjectName("accent")
        cl.clicked.connect(dlg.accept)
        bar.addWidget(cp)
        bar.addStretch()
        bar.addWidget(cl)
        lay.addLayout(bar)
        dlg.exec()

    def _export_edit_csv(self):
        n = self.edit_table.rowCount()
        if n == 0:
            QMessageBox.information(self, "Инфо", "Нет данных для выгрузки.")
            return

        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        dated_dir = os.path.join(desktop_dir, f"Выгрузка {datetime.date.today().strftime('%d.%m.%Y')}")
        try:
            os.makedirs(dated_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать папку:\n{dated_dir}\n{e}")
            return

        def unique_path(path):
            i = 2
            stem, ext = os.path.splitext(path)
            while os.path.exists(path):
                path = f"{stem} ({i}){ext}"
                i += 1
            return path

        def safe_name(city):
            return "".join(c for c in city if c.isalnum() or c in " _-").strip()

        # Разбиваем номера по городам (та же база, что и в столбце «Оператор»/«Город»).
        try:
            self._ensure_data()
        except Exception:
            pass
        buckets = {}   # город ("" — не определён) -> [номера]
        for r in range(n):
            item = self.edit_table.item(r, 0)
            if not item:
                continue
            num = item.text().strip()
            if not num:
                continue
            try:
                city, _op = self._lookup_phone_city_operator(num)
            except Exception:
                city = ""
            buckets.setdefault(city, []).append(num)

        # Убираем дубли номеров — отдельно внутри каждого города, чтобы не
        # задеть номера из другого города, даже если они случайно совпали.
        for city in buckets:
            buckets[city] = list(dict.fromkeys(buckets[city]))

        distinct_cities = [c for c in buckets if c]
        split_mode = True
        if len(distinct_cities) > 1:
            box = QMessageBox(self)
            box.setWindowTitle("Несколько городов")
            box.setIcon(QMessageBox.Question)
            box.setText(
                f"В файле номера {len(distinct_cities)} городов: {', '.join(distinct_cities)}.\n"
                "Сохранить отдельными файлами по городу или свести всё в один файл?")
            split_btn = box.addButton("Разделить по городам", QMessageBox.AcceptRole)
            combine_btn = box.addButton("Свести в один файл", QMessageBox.DestructiveRole)
            cancel_btn = box.addButton("Отмена", QMessageBox.RejectRole)
            box.setDefaultButton(split_btn)
            box.exec()
            clicked = box.clickedButton()
            if clicked is cancel_btn:
                return
            split_mode = clicked is split_btn

        try:
            saved = []
            if len(distinct_cities) <= 1:
                # один город (или он вообще не определился) — один файл, как раньше
                nums = list(dict.fromkeys(num for nums in buckets.values() for num in nums))
                city = next((c for c in buckets if c), "")
                file_name = f"{len(nums)}_{safe_name(city)}.csv" if city else f"{len(nums)}.csv"
                path = unique_path(os.path.join(dated_dir, file_name))
                pd.DataFrame({"Номер телефона": nums}).to_csv(path, index=False, encoding="utf-8-sig")
                saved.append((city or "—", len(nums), path))
            elif split_mode:
                # несколько городов — раскладываем по отдельным файлам
                for city, nums in buckets.items():
                    file_name = f"{len(nums)}_{safe_name(city)}.csv" if city else f"{len(nums)}_без_города.csv"
                    path = unique_path(os.path.join(dated_dir, file_name))
                    pd.DataFrame({"Номер телефона": nums}).to_csv(path, index=False, encoding="utf-8-sig")
                    saved.append((city or "без города", len(nums), path))
            else:
                # свести все города в один файл
                nums = list(dict.fromkeys(num for nums in buckets.values() for num in nums))
                if len(distinct_cities) <= 3:
                    combo_name = safe_name("+".join(distinct_cities))
                else:
                    combo_name = "несколько_городов"
                file_name = f"{len(nums)}_{combo_name}.csv"
                path = unique_path(os.path.join(dated_dir, file_name))
                pd.DataFrame({"Номер телефона": nums}).to_csv(path, index=False, encoding="utf-8-sig")
                saved.append((", ".join(distinct_cities), len(nums), path))

            total = sum(s[1] for s in saved)
            if len(saved) == 1:
                self._set_label(self.edit_status, f"Сохранено {total:,} номеров в «{dated_dir}»!".replace(",", " "), GREEN)
                QMessageBox.information(self, "Готово!", f"Сохранено {total:,} номеров.\n\n{saved[0][2]}".replace(",", " "))
            else:
                lines = [f"Разложено по {len(saved)} городам в «{dated_dir}»:", ""]
                for city, cnt, path in saved:
                    lines.append(f"• {city}: {cnt:,} номеров → {os.path.basename(path)}".replace(",", " "))
                self._set_label(self.edit_status, f"Сохранено {total:,} номеров, {len(saved)} файл(ов).".replace(",", " "), GREEN)
                QMessageBox.information(self, "Готово!", "\n".join(lines))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", self._file_err("Не удалось сохранить:", e))

    def _file_err(self, prefix, e):
        msg = str(e)
        if "openpyxl" in msg.lower():
            return (f"{prefix}\n\nНе найден пакет openpyxl (нужен для .xlsx).\n"
                    "Установите: pip install openpyxl")
        return f"{prefix}\n{msg}"

    # ---- триггер-слова ----
    def _get_trigger_re(self):
        if self._trigger_re is None:
            words = sorted({w for w in C.ACTIVE_TRIGGER_WORDS if w}, key=len, reverse=True)
            pattern = r"(?<!\w)(" + "|".join(re.escape(w) for w in words) + r")(?!\w)"
            self._trigger_re = re.compile(pattern, re.IGNORECASE)
        return self._trigger_re

    def _open_trigger_editor(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Список триггер-слов")
        dlg.resize(520, 640)
        dlg.setStyleSheet(build_qss())
        lay = QVBoxLayout(dlg)
        head = QLabel("✎  Триггер-слова")
        head.setStyleSheet(f"color:{FG}; font-weight:bold; font-size:13pt;")
        lay.addWidget(head)
        hint = QLabel("По одному слову/фразе на строку. Изменения применяются сразу "
                      "и сохраняются в trigger_words.txt рядом с программой.")
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{FG2};")
        lay.addWidget(hint)
        box = QPlainTextEdit()
        box.setPlainText("\n".join(sorted(set(w for w in C.ACTIVE_TRIGGER_WORDS if w), key=str.lower)))
        lay.addWidget(box, 1)
        status = QLabel("")
        status.setStyleSheet(f"color:{FG2};")
        lay.addWidget(status)

        def apply(save):
            seen, new = set(), []
            for line in box.toPlainText().splitlines():
                w = line.strip()
                if w and w.lower() not in seen:
                    seen.add(w.lower())
                    new.append(w)
            if not new:
                QMessageBox.warning(dlg, "Внимание", "Список не может быть пустым.")
                return
            C.ACTIVE_TRIGGER_WORDS = new
            self._trigger_re = None
            if save:
                try:
                    C.save_active_trigger_words(new)
                    status.setText(f"✅ Применено и сохранено ({len(new)} слов).")
                    status.setStyleSheet(f"color:{GREEN};")
                except Exception as e:
                    status.setText(f"⚠ Применено, но не сохранено: {e}")
                    status.setStyleSheet(f"color:{DUPE_RED};")
            else:
                status.setText(f"✅ Применено ({len(new)} слов), только для сеанса.")
                status.setStyleSheet(f"color:{GREEN};")

        bar = QHBoxLayout()
        b1 = QPushButton("Применить и сохранить")
        b1.setObjectName("green")
        b1.clicked.connect(lambda: apply(True))
        b2 = QPushButton("Применить (не сохранять)")
        b2.clicked.connect(lambda: apply(False))
        b3 = QPushButton("Закрыть")
        b3.clicked.connect(dlg.accept)
        bar.addWidget(b1)
        bar.addWidget(b2)
        bar.addStretch()
        bar.addWidget(b3)
        lay.addLayout(bar)
        dlg.exec()

    def _delete_city_from_edit(self, city, dlg=None):
        parent = dlg or self
        resp = QMessageBox.question(
            parent, "Подтверждение",
            f"Удалить все номера города «{city}»?",
            QMessageBox.Yes | QMessageBox.No)
        if resp != QMessageBox.Yes:
            return
        try:
            self._ensure_data()
        except Exception:
            pass
        t = self.edit_table
        rows_to_del = []
        for r in range(t.rowCount()):
            item = t.item(r, 0)
            if not item:
                continue
            try:
                c, _op = self._lookup_phone_city_operator(item.text())
            except Exception:
                c = ""
            if c == city:
                rows_to_del.append(r)
        if not rows_to_del:
            QMessageBox.information(parent, "Инфо", f"Номера города «{city}» не найдены.")
            return
        self._push_undo()
        for r in sorted(rows_to_del, reverse=True):
            t.removeRow(r)
        self._set_label(self.edit_status, f"🗑 Удалены номера города «{city}»: {len(rows_to_del)}", DUPE_RED)
        self._update_edit_count()
        # обновляем бейдж/разбивку по городам с учётом удаления
        self._detect_edit_city(self._all_edit_pairs())
        if dlg is not None:
            dlg.accept()

    def _open_trigger_scan(self):
        if self.edit_table.rowCount() == 0:
            QMessageBox.information(self, "Инфо", "Сначала загрузите данные для проверки.")
            return
        pattern = self._get_trigger_re()
        matches = []
        for r in range(self.edit_table.rowCount()):
            phone = self.edit_table.item(r, 0).text() if self.edit_table.item(r, 0) else ""
            trans = self.edit_table.item(r, 1).text() if self.edit_table.item(r, 1) else ""
            if not trans:
                continue
            found, seen = [], set()
            for m in pattern.finditer(trans):
                w = m.group(1).lower()
                if w not in seen:
                    seen.add(w)
                    found.append(w)
            if found:
                matches.append((r, phone, trans, found))
        multi_city = len(self._edit_city_counts) > 1
        if not matches and not multi_city:
            self._set_label(self.edit_status, "✅ Триггер-слова не найдены.", GREEN)
            return
        self._show_trigger_scan(matches)

    def _show_trigger_scan(self, matches):
        dlg = QDialog(self)
        dlg.setWindowTitle("Триггер-слова — найденные совпадения")
        dlg.resize(960, 580)
        dlg.setStyleSheet(build_qss())
        lay = QVBoxLayout(dlg)
        if matches:
            head = QLabel(f"⚠  Найдено строк с триггер-словами: {len(matches)}")
        else:
            head = QLabel("✅  Триггер-слова не найдены")
        head.setStyleSheet(f"color:{DUPE_RED if matches else GREEN}; font-weight:bold; font-size:13pt;")
        lay.addWidget(head)

        if len(self._edit_city_counts) > 1:
            city_box = QFrame()
            city_box.setObjectName("city_box")
            city_box.setStyleSheet(
                f"QWidget#city_box {{ background-color:{PANEL_BG};"
                f" border:1px solid {PANEL_BORDER}; border-radius:12px; }}")
            city_lay = QVBoxLayout(city_box)
            city_lay.setContentsMargins(10, 8, 10, 8)
            city_lay.setSpacing(6)
            city_head = QLabel("🏙 Обнаружено несколько городов — можно удалить номера одного из них:")
            city_head.setStyleSheet(f"color:{ACCENT}; font-weight:bold;")
            city_lay.addWidget(city_head)
            city_btns_row = QHBoxLayout()
            city_btns_row.setSpacing(8)
            ranked = sorted(self._edit_city_counts.items(), key=lambda kv: -kv[1])
            for city, cnt in ranked:
                cb = QPushButton(f"  Удалить «{city}» ({cnt})")
                cb.setIcon(_icon(ICONS["trash"], DANGER_FG, ""))
                cb.setIconSize(QSize(14, 14))
                cb.setObjectName("danger")
                cb.clicked.connect(lambda _, c=city: self._delete_city_from_edit(c, dlg))
                city_btns_row.addWidget(cb)
            city_btns_row.addStretch()
            city_lay.addLayout(city_btns_row)
            lay.addWidget(city_box)

        if not matches:
            close_only = QPushButton("Закрыть")
            close_only.setObjectName("accent")
            close_only.clicked.connect(dlg.reject)
            lay.addWidget(close_only)
            dlg.exec()
            return

        hint = QLabel("Отмеченные галочкой строки будут удалены. Снимите галочку, чтобы оставить строку.")
        hint.setStyleSheet(f"color:{FG2};")
        lay.addWidget(hint)

        tbl = QTableWidget(len(matches), 4)
        tbl.setHorizontalHeaderLabels(["✓", "Номер телефона", "Найденные слова", "Фрагмент"])
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.setColumnWidth(0, 36)
        tbl.setColumnWidth(1, 150)
        tbl.setColumnWidth(2, 220)
        tbl.setColumnWidth(3, 260)
        for i, (src_row, phone, trans, words) in enumerate(matches):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked)
            tbl.setItem(i, 0, chk)
            tbl.setItem(i, 1, txt_item(phone))
            tbl.setItem(i, 2, txt_item(", ".join(words)))
            snippet = trans if len(trans) <= 140 else trans[:137] + "…"
            tbl.setItem(i, 3, txt_item(snippet))
        lay.addWidget(tbl, 1)

        count_lbl = QLabel("")
        count_lbl.setStyleSheet(f"color:{FG2};")
        def upd_count():
            n = sum(1 for i in range(tbl.rowCount()) if tbl.item(i, 0).checkState() == Qt.Checked)
            count_lbl.setText(f"Отмечено на удаление: {n} из {len(matches)}")
        tbl.itemChanged.connect(lambda _: upd_count())
        upd_count()

        bar1 = QHBoxLayout()
        def set_all(state):
            for i in range(tbl.rowCount()):
                tbl.item(i, 0).setCheckState(state)
        ba = QPushButton("Выбрать все")
        ba.clicked.connect(lambda: set_all(Qt.Checked))
        bn = QPushButton("Снять все")
        bn.clicked.connect(lambda: set_all(Qt.Unchecked))
        bar1.addWidget(ba)
        bar1.addWidget(bn)
        bar1.addWidget(count_lbl)
        bar1.addStretch()
        lay.addLayout(bar1)

        bar2 = QHBoxLayout()
        def do_delete():
            to_del = [matches[i][0] for i in range(tbl.rowCount())
                      if tbl.item(i, 0).checkState() == Qt.Checked]
            if not to_del:
                QMessageBox.information(dlg, "Инфо", "Ничего не отмечено.")
                return
            self._push_undo()
            for r in sorted(to_del, reverse=True):
                self.edit_table.removeRow(r)
            self._set_label(self.edit_status, f"🗑 Удалено по триггер-словам: {len(to_del)}", DUPE_RED)
            self._update_edit_count()
            dlg.accept()
        db = QPushButton(" Удалить отмеченные")
        db.setIcon(_icon(ICONS["trash"], DANGER_FG, ""))
        db.setIconSize(QSize(15, 15))
        db.setObjectName("danger")
        db.clicked.connect(do_delete)
        cb = QPushButton("Отмена")
        cb.clicked.connect(dlg.reject)
        bar2.addStretch()
        bar2.addWidget(cb)
        bar2.addWidget(db)
        lay.addLayout(bar2)
        dlg.exec()

    # ═══════════════════════════  ВКЛАДКА «ГЕНЕРАЦИЯ»  ═══════════════════════════
    def _build_gen_tab(self):
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(20, 12, 20, 10)
        lay.setSpacing(6)

        self.gen_city = QLineEdit()
        self.gen_city.setPlaceholderText("Название города, например: Ставрополь")
        lay.addWidget(self.gen_city)

        self.gen_ranges = HintPlainTextEdit(
            "Диапазоны номеров (по одному в строке, формат НАЧАЛО-КОНЕЦ)")
        self.gen_ranges.setStyleSheet(
            f"background-color:{SURFACE}; color:{FG}; font-family:'Consolas'; font-size:14pt;"
            f"border:1px solid {BORDER}; border-radius:8px;")
        self.gen_ranges.setMinimumHeight(150)
        lay.addWidget(self.gen_ranges, 1)

        clean_ranges_btn = IconBtn(_icon(ICONS["broom"], FG, "🧹"), "Очистить диапазоны", tip_side="below")
        clean_ranges_btn.clicked.connect(self._clean_gen_ranges)
        gen_undo_btn = IconBtn(_undo_icon(22, FG), "Вернуть назад", tip_side="below")
        gen_undo_btn.clicked.connect(self._undo_gen_ranges)

        self.gen_chunk = QLineEdit("5000")
        self.gen_chunk.setFixedWidth(160)
        self.gen_chunk.setMinimumHeight(36)
        self.gen_btn = IconBtn(_icon(ICONS["bolt"], GREEN, "⚡"), "Сгенерировать номера", variant="green", tip_side="below")
        self.gen_btn.clicked.connect(self.run_generate)
        self.gen_btn.setFixedSize(46, 46)
        self.gen_btn.setIconSize(QSize(22, 22))

        row2 = QHBoxLayout()
        row2.setSpacing(10)
        row2.addWidget(clean_ranges_btn)
        row2.addWidget(gen_undo_btn)
        row2.addWidget(self.gen_chunk)
        row2.addWidget(self.gen_btn)
        row2.addStretch()
        lay.addLayout(row2)

        self.gen_progress = QProgressBar()
        self.gen_progress.setRange(0, 100)
        self.gen_progress.setValue(0)
        self.gen_progress.setTextVisible(False)
        self.gen_progress.setFixedHeight(6)
        lay.addWidget(self.gen_progress)

        self.gen_status = QLabel("")
        self.gen_status.setStyleSheet(f"color:{FG2};")
        self.gen_status.setWordWrap(True)
        lay.addWidget(self.gen_status)
        self._maybe_flip_layout(lay)
        return page

    def _set_gen_busy(self, busy):
        """Блокирует кнопку «Сгенерировать» на время работы в фоновом потоке,
        чтобы нельзя было запустить генерацию повторно поверх уже идущей,
        и чтобы было явно видно, что процесс действительно запущен."""
        self.gen_btn.setEnabled(not busy)
        self.gen_btn.setToolTip("Идёт генерация…" if busy else "Сгенерировать номера")
        if busy:
            self.gen_progress.setRange(0, 0)   # неопределённый режим — сразу видно, что процесс идёт
        else:
            self.gen_progress.setRange(0, 100)
            self.gen_progress.setValue(0)

    def _set_gen_progress(self, done, total):
        if total <= 0:
            return
        self.gen_progress.setRange(0, total)
        self.gen_progress.setValue(done)

    # ═══════════════════════════  ВКЛАДКА «ВКЛАДЫ»  ═══════════════════════════
    # Фиксированный формат столбцов (без ручного выбора):
    #   файл «Города»: A — ФИО, B — дата, C — адрес
    #   файл «ФИО»:    A — ФИО, B — дата, C — доход, D..G — вклады 2022-2025
    # Города подтягиваются автоматически, как только загружены оба файла.
    def _build_deposits_tab(self):
        page = QWidget()
        lay = QVBoxLayout(page)
        lay.setContentsMargins(20, 10, 20, 10)
        lay.setSpacing(8)

        load_row = QHBoxLayout()
        load_row.setSpacing(10)
        self.dep_f1_btn = IconBtn(_icon(ICONS["excel"], "#1D6F42", "📊"), "Загрузить файл «ФИО»", variant="green")
        self.dep_f1_label = QLabel("Файл не выбран")
        self.dep_f1_label.setStyleSheet(f"color:{FG2}; font-size:9.5pt;")
        load_row.addWidget(self.dep_f1_btn)
        load_row.addWidget(self.dep_f1_label, 1)
        lay.addLayout(load_row)

        load_row2 = QHBoxLayout()
        load_row2.setSpacing(10)
        self.dep_f2_btn = IconBtn(_icon(ICONS["map"], FG, ""), "Загрузить файл «Города»")
        self.dep_f2_label = QLabel("Файл не выбран")
        self.dep_f2_label.setStyleSheet(f"color:{FG2}; font-size:9.5pt;")
        load_row2.addWidget(self.dep_f2_btn)
        load_row2.addWidget(self.dep_f2_label, 1)
        lay.addLayout(load_row2)

        top_btns = QHBoxLayout()
        top_btns.setSpacing(10)
        clear_btn = QPushButton("  Очистить")
        clear_btn.setIcon(_icon(ICONS["eraser"], DUPE_RED, ""))
        clear_btn.setIconSize(QSize(15, 15))
        clear_btn.setObjectName("danger")
        top_btns.addWidget(clear_btn)
        save_btn = QPushButton("  Сохранить результат")
        save_btn.setIcon(_icon(ICONS["save"], FG, ""))
        save_btn.setIconSize(QSize(15, 15))
        top_btns.addWidget(save_btn)
        top_btns.addStretch()
        lay.addLayout(top_btns)

        self.dep_status = QLabel("")
        self.dep_status.setWordWrap(True)
        self.dep_status.setStyleSheet(f"color:{FG2}; font-size:9.5pt;")
        lay.addWidget(self.dep_status)

        filt_row = QHBoxLayout()
        filt_row.setSpacing(8)
        filt_row.addWidget(QLabel("Город:"))
        self.dep_city_filter = QComboBox()
        self.dep_city_filter.setEnabled(False)
        filt_row.addWidget(self.dep_city_filter, 1)
        filt_row.addWidget(QLabel("Столбец:"))
        self.dep_num_col = QComboBox()
        self.dep_num_col.setEnabled(False)
        filt_row.addWidget(self.dep_num_col, 1)
        filt_row.addWidget(QLabel("от:"))
        self.dep_min_edit = QLineEdit()
        self.dep_min_edit.setFixedWidth(90)
        self.dep_min_edit.setEnabled(False)
        filt_row.addWidget(self.dep_min_edit)
        filt_row.addWidget(QLabel("до:"))
        self.dep_max_edit = QLineEdit()
        self.dep_max_edit.setFixedWidth(90)
        self.dep_max_edit.setEnabled(False)
        filt_row.addWidget(self.dep_max_edit)
        apply_filter_btn = QPushButton("Применить")
        apply_filter_btn.setEnabled(False)
        filt_row.addWidget(apply_filter_btn)
        reset_filter_btn = QPushButton("Сбросить")
        reset_filter_btn.setEnabled(False)
        filt_row.addWidget(reset_filter_btn)
        lay.addLayout(filt_row)

        self.dep_table = QTableWidget(0, 0)
        self.dep_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.dep_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.dep_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.dep_table.setSortingEnabled(True)
        self.dep_table.verticalHeader().setDefaultSectionSize(26)
        self.dep_table.setStyleSheet("font-size:9.5pt;")
        copy_sc = QShortcut(QKeySequence.Copy, self.dep_table)
        copy_sc.activated.connect(self._dep_copy_selected)
        lay.addWidget(self.dep_table, 1)

        self._dep_apply_filter_btn = apply_filter_btn
        self._dep_reset_filter_btn = reset_filter_btn
        self._dep_state = {"df1": None, "df2": None, "result_cols": None, "city_col_idx": None}

        self.dep_f1_btn.clicked.connect(self._dep_open_f1)
        self.dep_f2_btn.clicked.connect(self._dep_open_f2)
        clear_btn.clicked.connect(self._dep_clear)
        apply_filter_btn.clicked.connect(self._dep_apply_filter)
        reset_filter_btn.clicked.connect(self._dep_reset_filter)
        save_btn.clicked.connect(self._dep_save)
        self._maybe_flip_layout(lay)
        return page

    @staticmethod
    def _dep_to_number(s):
        s = str(s).strip().replace(" ", "").replace(",", ".")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    def _dep_open_f1(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть файл «ФИО»", "", "Excel/CSV (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            df = self._smart_read_table(path)
        except Exception as e:
            self._set_label(self.dep_status, f"⚠ Не удалось открыть файл «ФИО»: {e}", DUPE_RED)
            return
        self._dep_state["df1"] = df
        self.dep_f1_label.setText(f"{os.path.basename(path)}  ({len(df)} строк)")
        self._dep_maybe_run()

    def _dep_open_f2(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть файл «Города»", "", "Excel/CSV (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            df = self._smart_read_table(path)
        except Exception as e:
            self._set_label(self.dep_status, f"⚠ Не удалось открыть файл «Города»: {e}", DUPE_RED)
            return
        self._dep_state["df2"] = df
        self.dep_f2_label.setText(f"{os.path.basename(path)}  ({len(df)} строк)")
        self._dep_maybe_run()

    def _dep_clear(self):
        self._dep_state = {"df1": None, "df2": None, "result_cols": None, "city_col_idx": None}
        self.dep_f1_label.setText("Файл не выбран")
        self.dep_f2_label.setText("Файл не выбран")
        self.dep_table.setSortingEnabled(False)
        self.dep_table.setRowCount(0)
        self.dep_table.setColumnCount(0)
        self.dep_table.setSortingEnabled(True)
        self.dep_city_filter.clear()
        self.dep_num_col.clear()
        self.dep_min_edit.clear()
        self.dep_max_edit.clear()
        for w in (self.dep_city_filter, self.dep_num_col, self.dep_min_edit,
                  self.dep_max_edit, self._dep_apply_filter_btn, self._dep_reset_filter_btn):
            w.setEnabled(False)
        self._set_label(self.dep_status, "", FG2)

    def _dep_maybe_run(self):
        if self._dep_state["df1"] is None or self._dep_state["df2"] is None:
            return
        df1, df2 = self._dep_state["df1"], self._dep_state["df2"]
        if df1.shape[1] < 1 or df2.shape[1] < 3:
            self._set_label(self.dep_status, "⚠ В одном из файлов не хватает столбцов (нужны A/B/C).", DUPE_RED)
            return

        name1_col = df1.columns[0]
        name2_col = df2.columns[0]
        addr_col = df2.columns[2]

        city_by_name = {}
        for _, row in df2.iterrows():
            nm = str(row.get(name2_col, "")).strip()
            if not nm:
                continue
            city = C.normalize_city_case(C.extract_city_from_address(row.get(addr_col, "")))
            if city:
                city_by_name[nm] = city

        orig_cols = [str(c) for c in df1.columns]
        t = self.dep_table
        t.setSortingEnabled(False)
        t.setColumnCount(len(orig_cols) + 1)
        t.setHorizontalHeaderLabels(orig_cols + ["Город"])
        t.setRowCount(len(df1))

        num_cols = []
        for c in df1.columns:
            vals = [v for v in df1[c] if str(v).strip()]
            if not vals:
                continue
            ok = sum(1 for v in vals if self._dep_to_number(v) is not None)
            if ok / len(vals) >= 0.7:
                num_cols.append(str(c))

        matched = 0
        cities_seen = set()
        for i, (_, row) in enumerate(df1.iterrows()):
            for j, c in enumerate(df1.columns):
                val = row.get(c, "")
                if str(c) in num_cols:
                    num = self._dep_to_number(val)
                    if num is None:
                        text = ""
                    elif num == int(num):
                        text = f"{num:,.0f}".replace(",", " ")
                    else:
                        text = f"{num:,.2f}".replace(",", " ")
                    t.setItem(i, j, NumItem(text, num if num is not None else -1e18))
                else:
                    t.setItem(i, j, txt_item(str(val)))
            nm = str(row.get(name1_col, "")).strip()
            city = city_by_name.get(nm, "")
            if city:
                matched += 1
                cities_seen.add(city)
            t.setItem(i, len(orig_cols), txt_item(city))
        t.setSortingEnabled(True)

        header = t.horizontalHeader()
        for j in range(len(orig_cols)):
            header.setSectionResizeMode(j, QHeaderView.Interactive)
            t.setColumnWidth(j, 140)
        header.setSectionResizeMode(len(orig_cols), QHeaderView.Interactive)
        t.setColumnWidth(len(orig_cols), 150)
        header.setStretchLastSection(True)

        self._dep_state["result_cols"] = orig_cols + ["Город"]
        self._dep_state["city_col_idx"] = len(orig_cols)
        total = len(df1)
        color = GREEN if matched == total else ACCENT
        self._set_label(self.dep_status, f"✅ Города подставлены для {matched} из {total}.", color)

        self.dep_city_filter.clear()
        self.dep_city_filter.addItem("Все города")
        self.dep_city_filter.addItems(sorted(cities_seen))
        self.dep_num_col.clear()
        self.dep_num_col.addItem("— не фильтровать —")
        self.dep_num_col.addItems(num_cols)
        for w in (self.dep_city_filter, self.dep_num_col, self.dep_min_edit, self.dep_max_edit,
                  self._dep_apply_filter_btn, self._dep_reset_filter_btn):
            w.setEnabled(True)

    def _dep_col_index(self, name):
        cols = self._dep_state.get("result_cols") or []
        return cols.index(name) if name in cols else None

    def _dep_apply_filter(self):
        t = self.dep_table
        city = self.dep_city_filter.currentText()
        ncol = self.dep_num_col.currentText()
        ncol_idx = self._dep_col_index(ncol) if ncol != "— не фильтровать —" else None
        lo = self._dep_to_number(self.dep_min_edit.text()) if self.dep_min_edit.text().strip() else None
        hi = self._dep_to_number(self.dep_max_edit.text()) if self.dep_max_edit.text().strip() else None
        city_col = self._dep_state.get("city_col_idx")
        for r in range(t.rowCount()):
            show = True
            if city != "Все города" and city_col is not None:
                it = t.item(r, city_col)
                if not it or it.text() != city:
                    show = False
            if show and ncol_idx is not None:
                it = t.item(r, ncol_idx)
                val = self._dep_to_number(it.text()) if it else None
                if lo is not None and (val is None or val < lo):
                    show = False
                if hi is not None and (val is None or val > hi):
                    show = False
            t.setRowHidden(r, not show)

    def _dep_reset_filter(self):
        self.dep_city_filter.setCurrentIndex(0)
        self.dep_num_col.setCurrentIndex(0)
        self.dep_min_edit.clear()
        self.dep_max_edit.clear()
        for r in range(self.dep_table.rowCount()):
            self.dep_table.setRowHidden(r, False)

    def _dep_copy_selected(self):
        idxs = self.dep_table.selectedIndexes()
        if not idxs:
            return
        by_row = {}
        for idx in idxs:
            by_row.setdefault(idx.row(), {})[idx.column()] = idx
        lines = []
        for r in sorted(by_row):
            cols = by_row[r]
            cells = [self.dep_table.item(r, c).text() if self.dep_table.item(r, c) else ""
                     for c in sorted(cols)]
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))

    def _dep_save(self):
        t = self.dep_table
        n = t.rowCount()
        if n == 0:
            QMessageBox.information(self, "Инфо", "Нет данных для сохранения — загрузи оба файла.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результат", "с_городами.xlsx", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        try:
            ncols = t.columnCount()
            headers = [t.horizontalHeaderItem(c).text() if t.horizontalHeaderItem(c) else f"col{c}"
                       for c in range(ncols)]
            rows = []
            for r in range(n):
                if t.isRowHidden(r):
                    continue
                rows.append([t.item(r, c).text() if t.item(r, c) else "" for c in range(ncols)])
            df_out = pd.DataFrame(rows, columns=headers)
            if path.lower().endswith(".csv"):
                df_out.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                df_out.to_excel(path, index=False)
            self._set_label(self.dep_status, f"💾 Сохранено ({len(rows)} строк): {path}", GREEN)
            QMessageBox.information(self, "Готово!", f"Сохранено:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", self._file_err("Не удалось сохранить:", e))

    def _clean_gen_ranges(self):
        """Приводит 'грязные' диапазоны к чистому виду НАЧАЛО-КОНЕЦ.

        Пример:
            +7-343-764-3200 до +7-343-764-3599   ->   3437643200-3437643599
            +7-343-763-3260 до +7-343-763-3399   ->   3437633260-3437633399
        """
        raw = self.gen_ranges.toPlainText()
        cleaned_lines = []

        def to_clean_number(s):
            # убираем всё, кроме цифр (плюсы, тире, пробелы)
            digits = re.sub(r"\D", "", s)
            # убираем код страны "7" у 11-значного номера
            if len(digits) == 11 and digits.startswith("7"):
                digits = digits[1:]
            return digits

        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            # Сначала делим по слову "до" — это граница между началом и концом
            # диапазона. Важно делать это ДО удаления тире, иначе тире внутри
            # самого номера (343-764-3200) перемешаются с разделителем диапазона.
            parts = re.split(r"\s+(?:до|→|->|-->)\s+", line, maxsplit=1, flags=re.IGNORECASE)
            if len(parts) == 2:
                start_raw, end_raw = parts
            else:
                # "до" нет — либо уже чистый формат "начало-конец",
                # либо просто одно число.
                stripped = line.replace("+", "").replace(" ", "")
                digits_only = stripped.replace("-", "")
                if not digits_only.isdigit():
                    continue
                dash_parts = stripped.split("-")
                if len(dash_parts) == 2:
                    start_raw, end_raw = dash_parts
                else:
                    cleaned_lines.append(to_clean_number(stripped))
                    continue

            start_num = to_clean_number(start_raw)
            end_num = to_clean_number(end_raw)
            if start_num and end_num:
                cleaned_lines.append(f"{start_num}-{end_num}")

        self._gen_ranges_undo.append(raw)
        if len(self._gen_ranges_undo) > 50:
            self._gen_ranges_undo.pop(0)

        self.gen_ranges.setPlainText("\n".join(cleaned_lines))
        self._set_label(self.gen_status, f"🧹 Диапазоны очищены: {len(cleaned_lines)} шт.", GREEN)

    def _undo_gen_ranges(self):
        if not self._gen_ranges_undo:
            self._set_label(self.gen_status, "↩ Нечего отменять.", FG2)
            return
        prev_text = self._gen_ranges_undo.pop()
        self.gen_ranges.setPlainText(prev_text)
        self._set_label(self.gen_status, "↩ Восстановлено предыдущее состояние диапазонов.", ACCENT)

    def _parse_gen_ranges(self, raw):
        ranges = []
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            if "-" not in line:
                raise ValueError(f"Строка «{line}» не в формате НАЧАЛО-КОНЕЦ.")
            s, e = [p.strip() for p in line.split("-", 1)]
            if not (s.isdigit() and e.isdigit()):
                raise ValueError(f"Строка «{line}» содержит нецифровые символы.")
            width = len(s)
            si, ei = int(s), int(e)
            if ei < si:
                raise ValueError(f"В строке «{line}» конец меньше начала.")
            ranges.append((si, ei, width))
        if not ranges:
            raise ValueError("Не указано ни одного диапазона.")
        return ranges

    def run_generate(self):
        # Полный «страховочный» перехват: что бы ни пошло не так на главном
        # потоке (до запуска фоновой генерации) — пользователь ГАРАНТИРОВАННО
        # увидит окно с текстом ошибки, а не молчаливо «ничего не произойдёт».
        try:
            self._run_generate_impl()
        except Exception as e:
            import traceback
            QMessageBox.critical(
                self, "Ошибка",
                f"Не удалось запустить генерацию:\n{e}\n\n{traceback.format_exc()}")

    def _run_generate_impl(self):
        # Читаем все поля на ГЛАВНОМ потоке (виджеты Qt небезопасно читать
        # из фонового потока) и сразу проверяем их — до того, как вообще
        # спрашивать папку для сохранения.
        try:
            ranges = self._parse_gen_ranges(self.gen_ranges.toPlainText())
        except ValueError as e:
            QMessageBox.warning(self, "Ошибка в диапазонах", str(e))
            return
        try:
            chunk = int(self.gen_chunk.text().strip())
            if chunk <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Ошибка",
                                 "«Номеров в одном файле» должно быть целым числом больше нуля.")
            return
        city = self.gen_city.text().strip() or "numbers"

        # Считаем, сколько номеров и файлов получится, ДО запуска потока —
        # если диапазоны заданы с ошибкой (например, перепутаны местами
        # начало/конец разных диапазонов) и результат окажется гигантским,
        # предупреждаем пользователя, а не «зависаем» без единого признака
        # того, что вообще что-то происходит.
        total_preview = sum(end - start + 1 for start, end, _ in ranges)
        if total_preview > 2_000_000:
            files_preview = (total_preview + chunk - 1) // chunk
            resp = QMessageBox.question(
                self, "Подтверждение",
                f"Диапазоны дают {total_preview:,} номеров ({files_preview:,} файл(ов)).\n"
                "Это может занять продолжительное время. Продолжить?".replace(",", " "),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if resp != QMessageBox.Yes:
                return

        out_dir = QFileDialog.getExistingDirectory(self, "Выбери папку для сохранения")
        if not out_dir:
            return

        self.sigGenBusy.emit(True)
        self._set_label(self.gen_status, "⏳ Запускаю генерацию…", FG2)
        threading.Thread(target=self._generate, args=(ranges, chunk, city, out_dir), daemon=True).start()

    def _generate(self, ranges, chunk, city, out_dir):
        safe = "".join(c for c in city if c.isalnum() or c in " _-").strip() or "numbers"
        now = datetime.datetime.now()
        date_str, time_str = now.strftime("%Y%m%d"), now.strftime("%H%M%S")
        try:
            os.makedirs(out_dir, exist_ok=True)
            self.sigGenStatus.emit("⏳ Собираю номера…", FG2)
            all_numbers = []
            for start, end, width in ranges:
                for n in range(start, end + 1):
                    all_numbers.append("7" + str(n).zfill(width))
            total = len(all_numbers)
            if total == 0:
                raise ValueError("Диапазоны не содержат ни одного номера.")
            self.sigGenStatus.emit(f"⏳ Перемешиваю {total:,} номеров…".replace(",", " "), FG2)
            random.shuffle(all_numbers)
            files_made = 0
            total_files_expected = (total + chunk - 1) // chunk
            self.sigGenProgress.emit(0, total_files_expected)
            for i in range(0, total, chunk):
                part = all_numbers[i:i + chunk]
                if not part:
                    break
                files_made += 1
                fname = os.path.join(out_dir, f"{safe}_{date_str}_{time_str}_{files_made:03d}.xlsx")
                pd.DataFrame(part).to_excel(fname, index=False, header=False)
                self.sigGenStatus.emit(
                    f"⏳ Сохраняю файл {files_made} из {total_files_expected}…", FG2)
                self.sigGenProgress.emit(files_made, total_files_expected)
            done_msg = f"✅ Готово: {total:,} номеров, {files_made} файл(ов) в «{out_dir}».".replace(",", " ")
            self.sigGenStatus.emit(done_msg, GREEN)
            self.sigToast.emit("Генерация завершена", done_msg, "ok")
            self.sigInfo.emit(
                "Готово!",
                f"Сгенерировано номеров: {total:,}\nФайлов создано: {files_made}\nПапка: {out_dir}".replace(",", " "))
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.sigGenStatus.emit(f"⚠ Ошибка: {e}", ACCENT)
            self.sigError.emit("Ошибка генерации", f"Не удалось сгенерировать номера:\n{e}\n\n{tb}")
        finally:
            self.sigGenBusy.emit(False)


def _make_app_icon():
    """Иконка приложения (окно/панель задач): тёмная плитка с янтарным
    знаком-«воронкой» — тот же знак, что и в боковой панели. Рисуется
    векторно под каждый размер, поэтому остаётся резкой и в трее, и в меню
    «Пуск» (без мыла от масштабирования одной картинки)."""
    icon = QIcon()
    # Иконку рисуем в тёмном варианте независимо от выбранной темы: она живёт
    # вне окна программы (панель задач, «Пуск»), где светлая плитка потерялась
    # бы на любом фоне.
    _dk = C.THEMES["vortex_dark"]
    tile_a, tile_b = _dk["BG3"], _dk["BG2"]
    arc1, arc2, arc3 = _dk["ACCENT_HV"], _dk["ACCENT"], _dk["ACCENT_SOFT"]

    for size in (16, 24, 32, 48, 64, 128, 256):
        px = QPixmap(size, size)
        px.fill(QColor(0, 0, 0, 0))
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        corner = size * 0.28

        tile = QPainterPath()
        tile.addRoundedRect(QRectF(0, 0, size, size), corner, corner)
        grad = QLinearGradient(0, 0, size, size)
        grad.setColorAt(0.0, QColor(_mix(tile_a, arc2, 0.18)))
        grad.setColorAt(1.0, QColor(_mix(tile_b, "#000000", 0.25)))
        p.fillPath(tile, QBrush(grad))

        bw = max(1.0, size * 0.02)
        pen = QPen(QColor(255, 255, 255, 30))
        pen.setWidthF(bw)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawPath(tile)

        # три пластины — тот же знак, что и в боковой панели
        p.setPen(Qt.NoPen)
        plate_h = size * 0.125
        gap = size * 0.085
        widths = (0.56, 0.42, 0.28)
        total = len(widths) * plate_h + (len(widths) - 1) * gap
        y = (size - total) / 2.0
        for i, wf in enumerate(widths):
            pw = size * wf
            x = (size - pw) / 2.0
            g = QLinearGradient(x, 0, x + pw, 0)
            g.setColorAt(0.0, QColor(arc2))
            g.setColorAt(0.55, QColor(arc1))
            g.setColorAt(1.0, QColor(arc3))
            p.setBrush(QBrush(g))
            p.drawRoundedRect(QRectF(x, y, pw, plate_h), plate_h / 2.0, plate_h / 2.0)
            y += plate_h + gap

        p.end()
        icon.addPixmap(px)

    return icon


def _load_fonts():
    """Подключает шрифт Manrope из папки программы (или из подпапки fonts/).

    Возвращает имя семейства, которое реально доступно: если файла шрифта
    рядом нет, тихо откатываемся на системный запасной — программа не должна
    падать из-за не скопированного .ttf."""
    try:
        base = C._app_dir()
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))

    for folder in (os.path.join(base, "fonts"), base):
        if not os.path.isdir(folder):
            continue
        try:
            names = os.listdir(folder)
        except Exception:
            continue
        for fn in names:
            if fn.lower().endswith((".ttf", ".otf")) and "manrope" in fn.lower():
                try:
                    QFontDatabase.addApplicationFont(os.path.join(folder, fn))
                except Exception:
                    pass

    # Manrope нужен именно набором начертаний (Regular/Medium/SemiBold/Bold/
    # ExtraBold). Если положить только одно, Qt дорисует жирность сам —
    # синтетический bold как раз и выглядит грязным и рваным.
    want = STYLE.get("font", "Segoe UI")
    try:
        available = set(QFontDatabase.families())
    except Exception:
        available = set()
    if want in available:
        return want
    return STYLE.get("font_fallback", "Segoe UI")


def _apply_palette(app):
    """Тёмная палитра уровня Qt — чтобы «родные» части интерфейса
    (диалоги выбора файлов, системные меню, подсказки) не выпадали белыми
    пятнами из тёмного оформления."""
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(BG))
    pal.setColor(QPalette.WindowText, QColor(FG))
    pal.setColor(QPalette.Base, QColor(SURFACE))
    pal.setColor(QPalette.AlternateBase, QColor(ROW_ODD))
    pal.setColor(QPalette.Text, QColor(FG))
    pal.setColor(QPalette.PlaceholderText, QColor(FG3))
    pal.setColor(QPalette.Button, QColor(GRAY_BTN))
    pal.setColor(QPalette.ButtonText, QColor(FG))
    pal.setColor(QPalette.BrightText, QColor(DUPE_RED_HV))
    pal.setColor(QPalette.Highlight, QColor(ACCENT))
    pal.setColor(QPalette.HighlightedText, QColor(ON_ACCENT))
    pal.setColor(QPalette.ToolTipBase, QColor(BG3))
    pal.setColor(QPalette.ToolTipText, QColor(FG))
    pal.setColor(QPalette.Link, QColor(ACCENT_HV))
    pal.setColor(QPalette.Disabled, QPalette.Text, QColor(FG3))
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(FG3))
    app.setPalette(pal)


class ShimmerFilter(QObject):
    """Перелив на главных (акцентных) кнопках при наведении.

    Цвета кнопки не меняются — вдоль диагонального градиента едет светлая
    зона, из-за чего синь выглядит «переливающейся», а не просто закрашенной.
    Стиль ставится inline и только на одну кнопку под курсором, поэтому
    пересчёт стилей не задевает остальной интерфейс.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._anims = {}

    _SKINS = {
        "accent": None,   # заполняется на лету (цвета зависят от темы)
        "green": None,
    }

    def eventFilter(self, obj, event):
        if not isinstance(obj, QAbstractButton) or obj.objectName() not in ("accent", "green"):
            return False
        if obj.property("no_shimmer"):
            return False
        if event.type() == QEvent.Enter and obj.isEnabled():
            self._start(obj)
        elif event.type() == QEvent.Leave:
            self._stop(obj)
        return False

    def _start(self, btn):
        self._stop(btn)
        anim = QVariantAnimation(btn)
        anim.setStartValue(-0.35)
        anim.setEndValue(0.35)
        anim.setDuration(1500)
        anim.setEasingCurve(QEasingCurve.InOutSine)
        anim.setLoopCount(-1)

        name = btn.objectName()
        if name == "green":
            light, mid, deep = _mix(GREEN_HV, "#ffffff", 0.35), GREEN_HV, GREEN
            fg = _best_text_color(GREEN, dark="#06231a", light="#f0fff8")
        else:
            light, mid, deep = ACCENT_HV, ACCENT, ACCENT_GRAD_DEEP
            fg = ON_ACCENT

        def _apply(v):
            try:
                btn.setStyleSheet(
                    f"QPushButton#{name} {{"
                    f" background: {_shift_grad(light, mid, deep, float(v))};"
                    f" color: {fg};"
                    f" border: 1px solid {_rgba('#ffffff', 0.42)};"
                    f" border-radius: {STYLE['radius_md']}px; font-weight: 800;"
                    " padding: 9px 18px; }")
            except RuntimeError:
                pass

        anim.valueChanged.connect(_apply)
        self._anims[btn] = anim
        anim.start()

    def _stop(self, btn):
        anim = self._anims.pop(btn, None)
        if anim is not None:
            anim.stop()
            anim.deleteLater()
            try:
                btn.setStyleSheet("")
            except RuntimeError:
                pass


class _RepaintOnPressFilter(QObject):
    """Общий фильтр событий на всё приложение: принудительно перерисовывает
    любую кнопку сразу после нажатия/отпускания мыши.

    Баг, который это чинит: кнопки со сложным QSS-фоном (полупрозрачные/
    градиентные заливки — свечение при наведении, кольцевое свечение у
    кнопок-иконок и т.п.) на некоторых версиях/платформах Qt иногда не
    перерисовываются сами после клика — кнопка визуально «пропадает» и
    появляется обратно только когда придёт следующее событие (например,
    наведение мыши форсирует repaint). Сама логика кнопок тут ни при чём —
    это кэш стилей Qt; лечится принудительным update() без каких-либо
    графических эффектов (в отличие от QGraphicsDropShadowEffect, который
    в этой программе уже вызывал похожие баги в сочетании с fade-анимациями
    — см. предупреждение у _glow), поэтому фикс полностью безопасен."""

    def eventFilter(self, obj, event):
        if isinstance(obj, QAbstractButton) and event.type() in (
                QEvent.MouseButtonPress, QEvent.MouseButtonRelease):
            obj.update()
        return False


def main():
    # Масштаб экрана передаём в Qt как есть (125%, 150% и т.д.), не округляя
    # до целого. Именно округление раньше и давало «пиксельный» текст: Qt
    # рисовал интерфейс в 100% и растягивал картинку до 125%, размывая буквы.
    # Ставится строго ДО создания QApplication, поэтому ручной вызов
    # SetProcessDpiAwareness больше не нужен — Qt6 выставляет режим сам.
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception:
        pass
    app = QApplication(sys.argv)
    # Fusion — единственный стиль Qt, который одинаково выглядит на всех
    # системах и полностью подчиняется нашему QSS (нативный Windows-стиль
    # частично игнорирует скругления и цвета).
    try:
        app.setStyle("Fusion")
    except Exception:
        pass
    family = _load_fonts()
    app.setFont(_ui_font(10.5, family=family))
    _apply_palette(app)
    app.setStyleSheet(build_qss())
    repaint_filter = _RepaintOnPressFilter(app)
    app.installEventFilter(repaint_filter)
    shimmer_filter = ShimmerFilter(app)
    app.installEventFilter(shimmer_filter)
    ripple_filter = RippleFilter(app)
    app.installEventFilter(ripple_filter)
    icon = _make_app_icon()
    app.setWindowIcon(icon)
    win = App()
    win.setWindowIcon(icon)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
