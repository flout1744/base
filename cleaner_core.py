"""Ядро приложения Base: данные и бизнес-логика (без GUI).
Автоматически извлечено из исходного cleaner.pyw — большие вшитые базы
(города, телефонные диапазоны, специализации, триггер-слова) скопированы
без изменений."""

import os
import re
import json
import gzip
import base64
import math
import sys
import shutil
import subprocess
import tempfile
import datetime
from urllib.parse import quote as _url_quote

try:
    import openpyxl  # noqa: F401  — явный импорт, чтобы PyInstaller включил пакет
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

import pandas as pd

def _app_dir():
    """Папка, где лежит .pyw / .exe — сюда же сохраняются файл с триггер-словами
    и выбранная тема оформления."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


THEME_FILE = os.path.join(_app_dir(), "theme.json")

# ─────────────────────────────  ВЕРСИЯ ПРОГРАММЫ  ─────────────────────────────
# Правило простое: любое изменение в программе — новая строка версии.
# Мелкая правка/починка → +0.1, крупная переделка → следующее целое число.
# APP_BUILD — время и дата сборки, показывается в окне «о программе».
APP_VERSION = "3.3"
APP_BUILD = "16:30 01.11.2026"
APP_VERSION_FULL = f"версия {APP_VERSION} от {APP_BUILD}"

# Короткий список того, что изменилось в этой версии (окно «о программе»).
APP_CHANGES = (
    "база расширена: +17 000 сёл, посёлков, ПГТ, станиц и хуторов (от 500 жителей)",
    "у новых пунктов — телефонный код и диапазоны их района по свежему реестру Минцифры",
    "в столбце специализации у новых пунктов показан тип: село, ПГТ, станица…",
)

# ─────────────────────────────  ОФОРМЛЕНИЕ: VORTEX  ─────────────────────────────
# В программе осталась ровно одна тема — Vortex (обсидиан + янтарь).
# Словари THEMES / ICON_SETS / STYLE_SETS сохранены (а не выпилены совсем),
# потому что на них завязаны загрузка/сохранение выбора и хелперы
# get_icon_set()/get_style_set(), которыми пользуется весь интерфейс.
# Просто теперь в каждом из них — один-единственный набор.
THEMES = {
    "vortex_dark": {
        "label": "Обсидиан",
        "swatch": "#3d74e4",
        "icon": "mdi6.hexagon-multiple-outline",

        # ── подложки ──
        "BG":      "#0f1115",   # общий холст окна
        "BG2":     "#08090c",   # «хром»: заголовок окна, рейл, статус-строка
        "BG3":     "#15181e",   # карточка контента
        "NAV_BG":  "#0b0d11",   # боковое меню

        # ── поверхности и линии ──
        "SURFACE": "#191d24",   # поля ввода, таблицы
        "BORDER":  "#252a33",   # рамки
        "DIVIDER": "#1c2128",   # разделители

        # ── акцент: глубокая синь с переливом в светлую лазурь ──
        # ACCENT_DEEP — нижний (тёмный) стоп градиента, ACCENT — основной цвет
        # (он же используется как цвет текста-акцента, поэтому не слишком тёмный),
        # ACCENT_HV — светлый верх перелива и все свечения при наведении,
        # ACCENT_SOFT — почти белая лазурь для мелких бликов и подписей.
        "ACCENT":     "#3d74e4",
        "ACCENT_HV":  "#6fa2ff",
        "ACCENT_DEEP": "#122a63",
        "ACCENT_SOFT": "#c3d8ff",

        # ── семантика ──
        "DUPE_RED": "#f2555a", "DUPE_RED_HV": "#ff8b90",
        "DANGER_BG": "#2a151a", "DANGER_FG": "#ffb4bb",
        "GREEN": "#34d399", "GREEN_HV": "#6ee7b7",
        "BLUE": "#60a5fa",

        # ── кнопки-«призраки» ──
        "GRAY_BTN": "#1b1f26", "GRAY_HV": "#252b34",

        # ── текст ──
        "FG": "#f3f5f8", "FG2": "#868e9b", "FG3": "#5b636f",
        "LOG_FG": "#b6c2d8",

        # ── таблицы ──
        "HOVER_ROW": "#1e232b", "ROW_ODD": "#13161c",
    },
    "vortex_light": {
        "label": "Дневная",
        "swatch": "#2a5fd0",
        "icon": "mdi6.white-balance-sunny|mdi.white-balance-sunny|fa5s.sun",

        # ── подложки: холодная бумага, карточка — чистый белый ──
        "BG":      "#e9edf4",   # общий холст окна
        "BG2":     "#dde3ec",   # «хром»: заголовок окна, статус-строка, рамка
        "BG3":     "#ffffff",   # карточка контента
        "NAV_BG":  "#f4f7fb",   # боковое меню

        "SURFACE": "#ffffff",   # поля ввода, таблицы
        "BORDER":  "#c9d2e0",
        "DIVIDER": "#dde3ec",

        # ── тот же синий акцент, но подобранный под светлый фон ──
        # ACCENT_SOFT здесь ТЁМНЫЙ: это цвет акцентного текста и подписей,
        # который на светлой подложке обязан быть контрастным (в тёмной теме
        # ту же роль играет почти белая лазурь).
        "ACCENT":     "#2a5fd0",
        "ACCENT_HV":  "#4a80f0",
        "ACCENT_DEEP": "#173a86",
        "ACCENT_SOFT": "#1b3f8f",

        "DUPE_RED": "#d92d20", "DUPE_RED_HV": "#b42318",
        "DANGER_BG": "#fee4e2", "DANGER_FG": "#912018",
        "GREEN": "#12a35d", "GREEN_HV": "#0e8b4e",
        "BLUE": "#2a5fd0",

        "GRAY_BTN": "#ffffff", "GRAY_HV": "#f2f5fa",

        "FG": "#111827", "FG2": "#5b6577", "FG3": "#6f7a8d",
        "LOG_FG": "#475467",

        "HOVER_ROW": "#eef4ff", "ROW_ODD": "#f7f9fc",
    },
    "vortex_graphite": {
        "label": "Графит",
        "swatch": "#8fa3c0",
        "icon": "mdi6.circle-slice-8|mdi.circle-slice-8|fa5s.circle",

        # Полностью нейтральные подложки, без цветного отлива. Всю
        # индивидуальность держит холодный стальной акцент.
        "BG":      "#131417",
        "BG2":     "#0c0d0f",
        "BG3":     "#191b1f",
        "NAV_BG":  "#0e0f12",

        "SURFACE": "#1d1f24",
        "BORDER":  "#2b2e34",
        "DIVIDER": "#212429",

        "ACCENT":     "#8fa3c0",
        "ACCENT_HV":  "#b9c8de",
        "ACCENT_DEEP": "#3a4658",
        "ACCENT_SOFT": "#dde5f0",

        "DUPE_RED": "#e2606a", "DUPE_RED_HV": "#ff8f97",
        "DANGER_BG": "#291619", "DANGER_FG": "#ffbcc1",
        "GREEN": "#4cc38a", "GREEN_HV": "#7fd9ad",
        "BLUE": "#7aa2d8",

        "GRAY_BTN": "#1e2126", "GRAY_HV": "#282c33",

        "FG": "#f0f2f5", "FG2": "#8b9199", "FG3": "#5f656d",
        "LOG_FG": "#b4bac2",

        "HOVER_ROW": "#22262c", "ROW_ODD": "#16181c",
    },

    "vortex_cherry": {
        "label": "Вишня",
        "swatch": "#d4485f",
        "icon": "mdi6.fruit-cherries|mdi.fruit-cherries|fa5s.wine-glass",

        # Подложки почти нейтральные, лишь с еле уловимым тёплым отливом.
        "BG":      "#141013",
        "BG2":     "#0c090b",
        "BG3":     "#1a1518",
        "NAV_BG":  "#100c0e",

        "SURFACE": "#1e191c",
        "BORDER":  "#2f272b",
        "DIVIDER": "#241e21",

        "ACCENT":     "#d4485f",
        "ACCENT_HV":  "#f5788c",
        "ACCENT_DEEP": "#6e1626",
        "ACCENT_SOFT": "#ffd0d8",

        # Раз акцент сам красный, «дубли» и опасные действия помечаем тёплым
        # янтарём — иначе два красных спорили бы друг с другом.
        "DUPE_RED": "#ef8b3c", "DUPE_RED_HV": "#ffb066",
        "DANGER_BG": "#2a1b12", "DANGER_FG": "#ffcfa3",
        "GREEN": "#4cc38a", "GREEN_HV": "#7fd9ad",
        "BLUE": "#7aa2d8",

        "GRAY_BTN": "#211b1e", "GRAY_HV": "#2c2529",

        "FG": "#f7f0f2", "FG2": "#9a8f93", "FG3": "#6b6165",
        "LOG_FG": "#c9bcc0",

        "HOVER_ROW": "#271f23", "ROW_ODD": "#181316",
    },

    "vortex_ocean": {
        "label": "Океан",
        "swatch": "#35b6c9",
        "icon": "mdi6.waves|mdi.waves|fa5s.water",

        "BG":      "#0e1417",
        "BG2":     "#070c0f",
        "BG3":     "#141b1f",
        "NAV_BG":  "#0a1013",

        "SURFACE": "#182126",
        "BORDER":  "#24313a",
        "DIVIDER": "#1b262c",

        "ACCENT":     "#35b6c9",
        "ACCENT_HV":  "#6fdcea",
        "ACCENT_DEEP": "#0d4a58",
        "ACCENT_SOFT": "#c2f0f7",

        "DUPE_RED": "#f2555a", "DUPE_RED_HV": "#ff8b90",
        "DANGER_BG": "#2a151a", "DANGER_FG": "#ffb4bb",
        # Успех делаем травянисто-зелёным, чтобы он не сливался с бирюзой
        "GREEN": "#63cc63", "GREEN_HV": "#93e093",
        "BLUE": "#5aa9e6",

        "GRAY_BTN": "#1a232a", "GRAY_HV": "#243039",

        "FG": "#eef4f6", "FG2": "#85949c", "FG3": "#5b6970",
        "LOG_FG": "#aebdc4",

        "HOVER_ROW": "#1d272e", "ROW_ODD": "#121a1e",
    },
}
THEME_ORDER = ["vortex_dark", "vortex_light", "vortex_graphite",
               "vortex_cherry", "vortex_ocean"]

# Единый набор иконок. Значение — список кандидатов через «|»: берётся первый,
# который есть в установленной версии qtawesome. Так набор mdi6 (скруглённые,
# «дорогие» контурные значки) не ломается на старых сборках библиотеки —
# там подхватится mdi или fa5s, а не безликий кружок-заглушка.
ICON_SETS = {
    "vortex_dark": {
        "folder":   "mdi6.folder-open-outline|mdi.folder-open-outline|fa5s.folder-open",
        "broom":    "mdi6.broom|mdi.broom|fa5s.broom",
        "map":      "mdi6.map-marker-radius|mdi.map-marker-radius|fa5s.map-marked-alt",
        "edit":     "mdi6.pencil-outline|mdi.pencil-outline|fa5s.edit",
        "bolt":     "mdi6.lightning-bolt|mdi.flash|fa5s.bolt",
        "warning":  "mdi6.alert-outline|mdi.alert-outline|fa5s.exclamation-triangle",
        "list":     "mdi6.format-list-bulleted|mdi.format-list-bulleted|fa5s.list-ul",
        "filter":   "mdi6.filter-variant|mdi.filter-variant|fa5s.filter",
        "eraser":   "mdi6.eraser|mdi.eraser|fa5s.eraser",
        "copy":     "mdi6.content-copy|mdi.content-copy|fa5s.copy",
        "trash":    "mdi6.trash-can-outline|mdi.trash-can-outline|fa5s.trash-alt",
        "shuffle":  "mdi6.shuffle-variant|mdi.shuffle-variant|fa5s.random",
        "excel":    "mdi6.microsoft-excel|mdi.microsoft-excel|fa5s.file-excel",
        "layers":   "mdi6.layers-triple-outline|mdi.layers-triple-outline|fa5s.layer-group",
        "phone":    "mdi6.phone-in-talk-outline|mdi.phone-in-talk|fa5s.phone-volume",
        "database": "mdi6.database-outline|mdi.database|fa5s.database",
        "check":    "mdi6.check-decagram|mdi.check-decagram|fa5s.check-circle",
        "google":   "fa5b.google",
        "save":     "mdi6.content-save-outline|mdi.content-save-outline|fa5s.save",
        "piggy":    "mdi6.piggy-bank-outline|mdi.piggy-bank|fa5s.piggy-bank",
        "hashtag":  "mdi6.pound|mdi.pound|fa5s.hashtag",
        "map_pin":  "mdi6.map-marker-outline|mdi.map-marker-outline|fa5s.map-marker-alt",
        "status":   "mdi6.playlist-check|mdi.playlist-check|fa5s.tasks",
        # ── иконки самого окна (свой заголовок вместо системного) ──
        "win_min":  "mdi6.window-minimize|mdi.window-minimize",
        "win_max":  "mdi6.window-maximize|mdi.window-maximize",
        "win_rest": "mdi6.window-restore|mdi.window-restore",
        "win_close": "mdi6.close|mdi.close",
        "info":     "mdi6.information-outline|mdi.information-outline|fa5s.info-circle",
        "search":   "mdi6.magnify|mdi.magnify|fa5s.search",
        "palette":  "mdi6.palette-outline|mdi.palette|fa5s.palette",
        "theme_light": "mdi6.white-balance-sunny|mdi.white-balance-sunny|fa5s.sun",
        "theme_dark":  "mdi6.weather-night|mdi.weather-night|fa5s.moon",
        "fullscreen":  "mdi6.fullscreen|mdi.fullscreen|fa5s.expand",
        "fullscreen_exit": "mdi6.fullscreen-exit|mdi.fullscreen-exit|fa5s.compress",
    },
}


def get_icon_set(theme_name=None):
    """Набор иконок (роль -> имя/список имён qtawesome)."""
    base = dict(ICON_SETS["vortex_dark"])
    base.update(ICON_SETS.get(theme_name or "vortex_dark", {}))
    return base


# ─────────────────────────  СТИЛЕВЫЕ ТОКЕНЫ  ─────────────────────────
# Шрифт, скругления, толщина рамок и т.п. — одним местом на всю программу,
# чтобы менять пропорции интерфейса, не вычищая цифры из сотен строк QSS.
STYLE_SETS = {
    "vortex_dark": {
        "font": "Manrope", "font_fallback": "Segoe UI", "font_mono": "Consolas",
        "radius_sm": 9, "radius_md": 11, "radius_lg": 18,
        "icon_btn_radius": 11, "border_w": 1, "btn_weight": 700,
        "nav_style": "vortex", "gradient_buttons": True,
        "frameless": True,        # свой заголовок окна вместо системного
        "resizable": False,       # окно фиксированного размера (без растягивания)
        "sidebar_w": 262,         # ширина бокового меню
        "titlebar_h": 46,         # высота своей полосы заголовка
        # Рабочий размер окна. Подобран так, чтобы самая широкая вкладка
        # («Вклады») и таблицы помещались без сжатия — при фиксированном
        # размере ничего не «съезжает». На маленьком экране окно само
        # ужимается до доступной области, но не меньше win_min_*.
        "win_w": 1400, "win_h": 850,
        "win_min_w": 1215, "win_min_h": 640,
    },
}


def get_style_set(theme_name=None):
    """Стилевые токены активной темы."""
    base = dict(STYLE_SETS["vortex_dark"])
    base.update(STYLE_SETS.get(theme_name or "vortex_dark", {}))
    return base


DEFAULT_THEME = "vortex_dark"


def load_saved_theme():
    """Имя темы, сохранённой в прошлый раз (или тема по умолчанию)."""
    try:
        with open(THEME_FILE, "r", encoding="utf-8") as f:
            name = json.load(f).get("theme", DEFAULT_THEME)
            return name if name in THEMES else DEFAULT_THEME
    except Exception:
        return DEFAULT_THEME


def save_theme(name):
    """Сохраняет выбор темы — применится при следующем запуске."""
    try:
        with open(THEME_FILE, "w", encoding="utf-8") as f:
            json.dump({"theme": name}, f)
        return True
    except Exception:
        return False


_active_theme_name = load_saved_theme()
_t = THEMES[_active_theme_name]

# ─────────────────────────────  ПАЛИТРА (из активной темы)  ─────────────────────────────
BG        = _t["BG"]
BG2       = _t["BG2"]
BG3       = _t["BG3"]
NAV_BG    = _t["NAV_BG"]
SURFACE   = _t["SURFACE"]
BORDER    = _t["BORDER"]
DIVIDER   = _t["DIVIDER"]

ACCENT       = _t["ACCENT"]
ACCENT_HV    = _t["ACCENT_HV"]
DUPE_RED     = _t["DUPE_RED"]
DUPE_RED_HV  = _t["DUPE_RED_HV"]
DANGER_BG    = _t["DANGER_BG"]
DANGER_FG    = _t["DANGER_FG"]
GREEN        = _t["GREEN"]
GREEN_HV     = _t["GREEN_HV"]
GRAY_BTN     = _t["GRAY_BTN"]
GRAY_HV      = _t["GRAY_HV"]

FG        = _t["FG"]
FG2       = _t["FG2"]
LOG_FG    = _t["LOG_FG"]

# Новые цвета оформления Vortex (нижний стоп градиента у кнопок, светлый блик,
# синий для второстепенных акцентов и третий уровень серого текста).
ACCENT_DEEP = _t.get("ACCENT_DEEP", _t["ACCENT"])
ACCENT_SOFT = _t.get("ACCENT_SOFT", _t["ACCENT_HV"])
BLUE        = _t.get("BLUE", "#60a5fa")
FG3         = _t.get("FG3", _t["FG2"])

HOVER_ROW = _t["HOVER_ROW"]   # подсветка строки таблицы при наведении
ROW_ODD   = _t["ROW_ODD"]     # фон нечётной строки таблицы

FONT       = ("Segoe UI", 12)
FONT_BOLD  = ("Segoe UI", 12, "bold")
FONT_TITLE = ("Segoe UI", 20, "bold")
FONT_BTN   = ("Segoe UI", 13, "bold")
FONT_BTN_SM = ("Segoe UI", 11, "bold")
FONT_SMALL = ("Segoe UI", 10, "bold")
FONT_HINT  = ("Segoe UI", 11)
FONT_HINT_ITALIC = ("Segoe UI", 11, "italic")

# Единый шрифт таблиц.
TABLE_FONT      = ("Segoe UI", 10)
TABLE_FONT_MONO = ("Consolas", 10)
LOG_FONT_MONO   = ("Consolas", 11)

# Крупный подписанный шрифт для меток над важными полями (например, «ПАПКИ»).
FONT_LABEL_LG = ("Segoe UI", 13, "bold")

_ICON_FONT_CANDIDATES = ("Segoe Fluent Icons", "Segoe MDL2 Assets", "Segoe UI Symbol", "Segoe UI")
ICON_FONT_FAMILY = _ICON_FONT_CANDIDATES[-1]

# Единая шкала размеров значков (раньше значения скакали 9/13/15/18/22/24 —
# отсюда ощущение, что значки "то мелкие, то крупные, то кривые").
ICON_SM = 16
ICON_MD = 18
ICON_LG = 22
ICON_XL = 26

# ── Современные иконки: глифы Segoe Fluent/MDL2, рисуются шрифтом icon_font() ──
ICO_FOLDER   = "\uE8B7"
ICO_SEARCH   = "\uE721"
ICO_DELETE   = "\uEA0D"
ICO_COPY     = "\uE8C8"
ICO_SAVE     = "\uE74E"
ICO_CUT      = "\uE8C6"
ICO_SHUFFLE  = "\uE8B1"
ICO_CHECK    = "\uE73E"
ICO_CLOSE    = "\uE711"
ICO_PHONE    = "\uE717"
ICO_EDIT     = "\uE70F"
ICO_MAP      = "\uE707"
ICO_EXPORT   = "\uE898"
ICO_WARN     = "\uE7BA"

TRIGGER_WORDS = [
    "организация", "больница", "школа", "сад", "суд", "полиция", "охрана", "охранное", "агентство", "бюро",
    "детский сад", "здравоохранение", "медицинские учреждения", "психоневрологический диспансер",
    "наркологическое отделение", "стоматологическая клиника", "ресторан", "фотостудия", "апартаменты",
    "библиотека", "городской", "магазин", "салон", "кулинария", "бухгалтерия", "торговая сеть",
    "военный комиссариат", "муниципальное образование", "атс", "автоответчик", "служба поддержки", "центр",
    "областной", "дезинфекция", "поликлиника", "отдел", "училище", "управление", "фонд", "диспетчерская",
    "диспетчер", "филиал", "отель", "отделение", "дежурный", "клуб", "агент", "мастерская", "банк",
    "факультет", "кабинет", "аптека", "пост", "институт", "приемная", "лаборатория", "служба", "общежитие",
    "гостиница", "парк", "аквапарк", "парикмахерская", "клиника", "кафе", "ломбард", "проходная", "охраны",
    "система", "предприятие", "майор", "росгвардия", "капитан", "стоматология", "полковник", "подполковник",
    "полиции", "садик", "экспертиза", "ИВЦ", "сургутский", "министерство", "собрание", "правовой", "отряд",
    "ЖКХ", "департамент", "штаб", "вахта", "протокол", "оператор", "участок", "помощь", "помощник",
    "федеральная", "вас приветствует", "эксперт", "групп", "дилер", "компанию", "горячую линию",
    "горячая линия", "автомойка", "музей", "колледж", "учебная часть", "оптика", "палата", "цнк", "торговый",
    "инспекция", "общество", "храм", "терминал", "производится запись", "корректорская", "правком",
    "по делам", "консульство", "секретарь", "секретари", "клаб", "дежурная", "оперативный", "театр",
    "гимназия", "дирекция", "медкомиссия", "фирма", "военкомат", "инкассация", "ростелеком", "оборудование",
    "автокомплекс", "набранный вами номер", "в данный момент сотрудник которому", "счетная", "комитет",
    "следователь", "вы позвонили в", "котельная", "неправильно набран", "в целях повышения", "канцелярия",
    "хирургия", "инспектор", "администратор", "суши", "межрегионгаз", "газпром", "рфц", "магнит",
    "типография", "столовая", "мто", "мтс", "страх", "касса", "доктор", "казначейство", "сервис", "комиссия",
    "такси", "дворец спорта", "гарантия", "мебельный", "администрация", "ателье", "овд", "увд", "мвд", "фсб",
    "сжд", "повышение качества", "амурэнергосбыт", "аморэнергосбыт", "внутренний номер", "рдс", "следствие",
    "информационный", "завод", "фабрика", "подразделение", "запчасти", "нейрохирургия", "энергосервис",
    "цсо", "неврология", "регистратура", "сауна", "дом культуры", "бар", "ассоциация", "пекарня",
    "реанимация", "спа-комплекс", "ревматология", "мебель", "холдинг", "квп", "учреждение", "хостел",
    "радиология", "кардиология", "кпп", "теплогаз", "окб", "прокуратура", "почта россии", "ремонт",
    "абонент сейчас", "почта", "химчистка", "терапевт", "техникум", "россервис", "амбулатория",
    "пожарно-спасательная", "стадион", "наркология", "топливный", "пульт", "лицей", "дворец культуры",
    "медсанчасть", "санаторий", "бассейн", "компания", "ск", "типгорафия", "спа-комлекс", "терапия",
    "россельхоз", "офис", "бутик", "маркет", "кооператив", "гараж", "охранник", "энергосбыт", "бти",
    "подстанция", "архив", "ритуал", "инфекционная", "аптечный пункт", "азс", "соцподдержка", "база",
    "пожарный", "техподдержка", "энерго", "медпункт", "редакция", "интернат", "инфекция", "корпус", "вокзал",
    "томография", "флюрография", "роддом", "университет", "управляющая", "лейтенант", "пицерия", "кристал",
    "техно", "набранный вами", "боулинг", "кофейня", "водоканал", "академия", "мрт", "диспансер",
    "психиатрия", "контора", "студия", "телеграф", "коммутатор", "Набран", "номер", "добрый день",
    "позвонили", "комп", "Автосервис", "админист", "нажмите", "бухгалтер", "Лента", "клиники", "канцтовары",
    "АДВОКАТ", "апар", "дежурна", "не исп", "режиме", "смотат", "район", "охран", "разговор", "игротех",
    "хлеб", "Москва", "Термин", "Набор", "Парик", "плазм", "социальный", "Интерфина", "логистика", "логомед",
    "консультация", "Матрас", "выбер", "контроль", "Стройпост", "Системная", "Кадров", "мастер.",
    "информация", "электроник", "коеледж", "реабилитация", "детская", "категорию", "Получать", "смена",
    "комната", "церковь", "бесплатной", "платной", "доставкой", "доставка", "розыск", "управа", "ритуальный",
    "ритуальные", "услуги", "ромашка", "гвоздика", "дантист", "донтист", "газ", "жилищник", "аттестация",
    "телеканал", "отелем", "сон", "стационар", "стацион", "стацио", "дизайн", "галерея", "пресс", "преса",
    "дом", "мосархив", "роспотребнадзор", "надзор", "анатомия", "технический", "поддержки", "вкусна",
    "вкусно", "точка", "минист", "курсы", "кондитерская", "политики", "профилакторий", "контракт", "прием",
    "приема", "альфа-сервис", "арбитраж", "диагностика", "этология", "автофан", "пункт", "здравпункт",
    "литон", "машина", "люкс", "комфорт", "электросети", "автомат", "пульмонология", "общежитию",
    "екатерина", "канал", "гинекология", "старший", "мфц", "зал", "электросеть", "недоступен", "кпк",
    "метод", "обратится", "вопросам", "учебный", "комбинат", "ювелир", "фон", "личная", "мастер", "нейрон",
    "Елена", "Валерия", "обратились", "электрики", "электрик", "метстрой", "строй", "полиграфия", "модуль",
    "техномодуль", "проектов", "финансирование", "искусств", "искусство", "юбилейный", "нарко",
    "реставрация", "радио", "мир", "работу", "работа", "безопасности", "абонент", "александров",
    "страхование", "иванов", "культуры", "начальник", "комиссариат", "просторы", "росгосстрах", "таксит",
    "тдск", "росреестр", "редактор", "справочная", "фнс", "собор", "сектор", "изолит", "запись", "концепт",
    "часть", "пансионат", "бетон", "смены", "заказ", "заказа", "яндекс", "караульная", "вымпел", "генерация",
    "автозапчасти", "городское", "рынок", "гидромонтаж", "таможня", "деканат", "закупок", "цум", "окна",
    "пвх", "прапорщик", "анимация", "спектр", "мото", "батальон", "блок", "рентген", "стп", "кардиологи",
    "сурдологи", "ип", "ипп", "урология", "статистика", "гематология", "пдс", "интернет", "патология",
    "гастроэнтрология", "курс", "озгд", "кпз", "проктология", "упаковка", "раев", "пахов", "голубев",
    "единый", "стажер", "пдо", "город", "приставы", "колония", "аварийная", "заместитель", "кдс",
    "снабжение", "жилищная", "геодезия", "виноградов", "кузнецов", "макаров", "макаев", "родионов", "рякин",
    "сагаев", "сухарев", "хохлов", "шишков", "суботин", "альянс", "круглов", "ресепшн", "кац",
    "территориальная", "сигнал", "лесничество", "пенсионный", "соцзащита", "баня", "попов", "никонов",
    "курьер", "гбо", "громов", "громцев", "манолов", "тинькофф", "птс", "кончев", "гордиенко", "гусев",
    "хроматография", "сериков", "сабуров", "струков", "рашин", "прудников", "дума", "управляющий", "мбу",
    "видеонаблюдение", "спутник", "лобанов", "азк", "михайлов", "ксо", "карманов", "альфа", "бородин", "кп",
    "кадры", "закуп", "эконом", "этаж", "стол", "госпиталь", "минтруд", "спецучет", "стоянка", "парковка",
    "металл", "центральный", "бударин", "после сигнала", "борисов", "боровик", "конюшня", "юниор",
    "сельсовет", "полигон", "пожарное", "анализ", "фото", "ростехнадзор", "приемный", "коттедж",
    "региональный", "склад", "информ", "госконтроль", "оператив", "скорая", "112", "свр", "роскадастр",
    "нефтебаза", "судья", "техотдел", "турагентство", "погода", "подъезд", "конвой", "вас слушают",
    "психологи", "минстрой", "тихомиров", "защита", "ермаков", "росимущество", "турфирма", "союз",
    "психология", "кафедра", "минэнерго", "группа", "архитектура", "прачечная", "зао", "пао", "ооо",
    "недвижимость", "электроснабжение", "чоп", "узел связи", "лечение", "колесник", "никитина", "смирнов",
    "смирнова", "кузнецова", "петров", "петрова", "соколов", "соколова", "михайлова", "новиков", "новикова",
    "фёдоров", "фёдорова", "морозов", "морозова", "волков", "волкова", "алексеев", "алексеева", "лебедев",
    "лебедева", "семёнов", "семёнова", "егоров", "егорова", "павлов", "павлова", "козлов", "козлова",
    "степанов", "степанова", "николаев", "николаева", "орлов", "орлова", "белов", "белова", "голубева",
    "миронов", "миронова", "виноградова", "сергеев", "сергеева", "магомедов", "магомедова", "алиев",
    "алиева", "григорьев", "григорьева", "александрова", "андреев", "андреева", "макарова", "матвеев",
    "матвеева", "хакимов", "хакимова", "галлямов", "галлямова", "ахмадов", "ахмадова", "богданов",
    "богданова", "ооржак", "монгуш", "доржиев", "доржиева", "галиуллин", "галиуллина", "юсупов", "юсупова",
    "сафин", "сафина", "гаджиев", "гаджиева", "омаров", "омарова", "ибрагимов", "ибрагимова", "сидорова",
    "комаров", "комарова", "лазарев", "лазарева", "воробьёв", "воробьёва", "кудрявцев", "кудрявцева",
    "яковлев", "яковлева", "зуев", "зуева", "антонов", "антонова", "гусева", "денисов", "денисова",
    "киселёв", "киселёва", "ефимов", "ефимова", "сорокин", "сорокина", "ковалёв", "ковалёва", "баранов",
    "баранова", "титов", "титова", "шарапов", "шарапова", "мартынов", "мартынова", "кудряшов", "кудряшова",
    "ларионов", "ларионова", "винокуров", "винокурова", "цыденов", "цыденова", "салчак", "куулар",
    "муртазин", "муртазина", "шайхутдинов", "шайхутдинова", "сайдулаев", "сайдулаева", "хасанов", "хасанова",
    "оздоев", "оздоева", "аушев", "аушева", "кшнякин", "кшнякина", "мокшонин", "мокшонина", "батоев",
    "батоева", "лхасаранов", "лхасаранова", "ондар", "гулиев", "гулиева", "беков", "бекова", "тумгоев",
    "тумгоева", "султанов", "султанова", "хабибуллин", "хабибуллина", "мингазов", "мингазова", "дадаев",
    "дадаева", "рустамов", "рустамова", "исмаилов", "исмаилова", "мусаев", "мусаева", "керимов", "керимова",
    "абдуллаев", "абдуллаева", "захаров", "захарова", "тихонов", "тихонова", "борисова", "гаврилов",
    "гаврилова", "ильин", "ильина", "дмитриев", "дмитриева", "никитин", "медведев", "медведева", "карпов",
    "карпова", "фролов", "фролова", "савельев", "савельева", "тимофеев", "тимофеева", "панфилов",
    "панфилова", "константинов", "константинова", "герасимов", "герасимова", "ершов", "ершова", "калинин",
    "калинина", "журавлёв", "журавлёва", "максимов", "максимова", "белоусов", "белоусова", "савин", "савина",
    "потапов", "потапова", "лаврентьев", "лаврентьева", "орёхов", "орёхова", "куприн", "куприна", "чернышёв",
    "чернышёва", "пономарёв", "пономарёва", "галкин", "галкина", "самойлов", "самойлова", "шишкин",
    "шишкина", "минеев", "минеева", "доронин", "доронина", "гришин", "гришина", "бондарев", "бондарева",
    "терентьев", "терентьева", "фомин", "фомина", "абрамов", "абрамова", "жуков", "жукова", "третьяков",
    "третьякова", "суханов", "суханова", "князев", "князева", "беляев", "беляева", "королёв", "королёва",
    "крюков", "крюкова", "рябов", "рябова", "шулин", "шулина", "малинин", "малинина", "павленко", "сазонов",
    "сазонова", "марков", "маркова", "ткачёв", "ткачёва", "булатов", "булатова", "хусаинов", "хусаинова",
    "абдурахманов", "абдурахманова", "салимов", "салимова", "нуриев", "нуриева", "закиров", "закирова",
    "кадыров", "кадырова", "расулов", "расулова", "шарафутдинов", "шарафутдинова", "мухамедов", "мухамедова",
    "гафуров", "гафурова", "исламов", "исламова", "фаттахов", "фаттахова", "гумеров", "гумерова", "байрамов",
    "байрамова", "сабиров", "сабирова", "хамитов", "хамитова", "якупов", "якупова", "садыков", "садыкова",
    "набиуллин", "набиуллина", "агзамов", "агзамова", "мустафин", "мустафина", "ракипов", "ракипова",
    "гаязов", "гаязова", "шакиров", "шакирова", "гараев", "гараева", "абдрахманов", "абдрахманова",
    "идиятов", "идиятова", "мухутдинов", "мухутдинова", "саттаров", "саттарова", "халиуллин", "халиуллина",
    "загидуллин", "загидуллина", "каримов", "каримова", "файзуллин", "файзуллина", "бикбов", "бикбова",
    "гильмутдинов", "гильмутдинова", "нурмухаметов", "нурмухаметова", "абдрахимов", "абдрахимова",
    "шафигуллин", "шафигуллина", "гатауллин", "гатауллина", "саитов", "саитова", "муллаян", "муллаянова",
    "гафиятуллин", "гафиятуллина", "зигануров", "зиганурова", "хаматнуров", "хаматнурова", "ягафаров",
    "ягафарова", "абдуллин", "абдуллина", "магадеев", "магадеева", "гаянова", "нурлыгаянов", "нурлыгаянова",
    "сафиуллин", "сафиуллина", "абдулхаков", "абдулхакова", "мухаметзянов", "мухаметзянова", "устюжанин",
    "устюжанина", "костромин", "костромина", "пинегин", "пинегина", "вологжанин", "вологжанина", "вагин",
    "вагина", "вяткин", "вяткина", "мезенцев", "мезенцева", "каргопольцев", "каргопольцева", "новгородов",
    "новгородова", "иноземцев", "иноземцева", "новик", "сороковиков", "сороковикова", "новокрещенов",
    "новокрещенова", "ермака", "мещериаков", "мещериакова", "пан", "пана", "брязгин", "брязгина", "кольцов",
    "кольцова", "аршинский", "аршинская", "пешков", "пешкова", "плотников", "плотникова", "рукосуев",
    "рукосуева", "гантимуров", "гантимурова", "шубин", "шубина", "демидов", "демидова", "караулов",
    "караулова", "турунов", "турун", "кошкарёв", "кошкарёва", "черепанов", "черепанова", "шадрин", "шадрина",
    "тобольцев", "тобольцева", "томилов", "томилова", "иртышский", "иртышская", "обский", "обская",
    "енисейский", "енисейская", "ленский", "ленская", "ангарский", "ангарская", "байкальский", "байкальская",
    "алтайцев", "алтайцева", "катунский", "катунская", "бийский", "бийская", "барнаульцев", "барнаульцева",
    "новосибирцев", "новосибирцева", "омский", "омская", "томич", "томичка", "красноярец", "красноярка",
    "иркутянин", "иркутянка", "якутянин", "якутянка", "савватciach", "давыдов", "давыдова", "багин",
    "багина", "турок", "туркова", "хандаров", "хандарова", "сангинов", "сангинова", "дашиев", "дашиева",
    "балданов", "балданова", "дамдинов", "дамдинова", "жамсаранов", "жамсаранова", "хингелов", "хингелова",
    "гармаев", "гармаева", "дондуков", "дондукова", "бадмаев", "бадмаева", "цыремпилов", "цыремпилова",
    "атаганов", "атаганова", "дондоков", "дондокова", "содномов", "содномова", "тумуров", "тумурова",
    "халтанов", "халтанова", "буянтуев", "буянтуева", "лхагванов", "лхагванова", "намсараев", "намсараева",
    "раднаев", "раднаева", "содбоев", "содбоева", "тунгусов", "тунгусова", "хараев", "хараева", "эрдынеев",
    "эрдынеева", "жалсанов", "жалсанова", "дашибалов", "дашибалова", "базаров", "базарова", "дамбаев",
    "дамбаева", "лодонов", "лодонова", "саганов", "саганова", "хартаев", "хартаева", "цыбикжапов",
    "цыбикжапова", "дамдинжапов", "дамдинжапова", "тумунов", "тумунова", "жамбалов", "жамбалова",
    "спасатели", "машинисты", "газовая", "заправка", "депо", "компрессорная", "пожарная",
    "электромастерская", "ортопед", "ортопедия", "бурденко", "лисенко", "кашин", "кашина", "филатов",
    "филатова", "семин", "семина", "федоров", "федорова", "учет", "караоке", "бокс", "терехов", "терехова",
    "ахмедов", "ахмедова", "хореография", "насосная", "узи", "росгвардии", "рабочий", "аэродром", "коллеги",
    "лазарет", "вход", "служебный", "дорожная", "станция", "мода", "втб", "пожарно-спасательный",
    "посольство", "зауч", "перевозки", "динамо",
]

TRIGGER_WORDS_FILE = os.path.join(_app_dir(), "trigger_words.txt")

def load_active_trigger_words():
    """Если рядом с программой лежит trigger_words.txt — берём слова оттуда
    (это сохранённый пользователем список, включающий его правки).
    Иначе используем встроенный список TRIGGER_WORDS."""
    if os.path.isfile(TRIGGER_WORDS_FILE):
        try:
            with open(TRIGGER_WORDS_FILE, "r", encoding="utf-8-sig") as f:
                words = [line.strip() for line in f if line.strip()]
            if words:
                return words
        except Exception:
            pass
    return list(TRIGGER_WORDS)

def save_active_trigger_words(words):
    """Сохраняет текущий список триггер-слов в файл рядом с программой,
    чтобы правки не терялись при перезапуске."""
    with open(TRIGGER_WORDS_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(words))

ACTIVE_TRIGGER_WORDS = load_active_trigger_words()

# ─────────────────────────────  ВШИТЫЕ ДАННЫЕ  ─────────────────────────────
ICON_B64 = "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"
_CITY_B64 = "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"
_PHONES_B64 = "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"
INDUSTRY = json.loads(r"""{
  "Обнинск": "☢ наукоград, первая в мире АЭС, атомная наука",
  "Дубна": "☢ наукоград, ядерная физика (ОИЯИ)",
  "Саров": "☢ ЗАТО, ядерный центр РФЯЦ-ВНИИЭФ (атомное оружие)",
  "Снежинск": "☢ ЗАТО, ядерный центр ВНИИТФ",
  "Озёрск": "☢ ЗАТО, ПО «Маяк», радиохимия",
  "Северск": "☢ ЗАТО, Сибирский химкомбинат (атомная)",
  "Новоуральск": "☢ ЗАТО, УЭХК — обогащение урана",
  "Зеленогорск": "☢ ЗАТО, ЭХЗ — обогащение урана",
  "Лесной": "☢ ЗАТО, комбинат «Электрохимприбор»",
  "Трёхгорный": "☢ ЗАТО, приборостроительный завод (атомная)",
  "Заречный": "☢ ЗАТО, приборостроение (атомная)",
  "Железногорск": "☢ ЗАТО, ГХК + спутники (ИСС Решетнёва)",
  "Волгодонск": "⚙ Атоммаш — атомное машиностроение, Ростовская АЭС",
  "Десногорск": "☢ Смоленская АЭС",
  "Курчатов": "☢ Курская АЭС",
  "Полярные Зори": "☢ Кольская АЭС",
  "Удомля": "☢ Калининская АЭС",
  "Нововоронеж": "☢ Нововоронежская АЭС",
  "Сосновый Бор": "☢ Ленинградская АЭС",
  "Димитровград": "☢ НИИАР — атомные реакторы",
  "Балаково": "☢ Балаковская АЭС, химволокно",
  "Череповец": "⛓ металлургия «Северсталь», химия (ФосАгро)",
  "Магнитогорск": "⛓ ММК — чёрная металлургия",
  "Нижний Тагил": "⛓ Уралвагонзавод, металлургия (ЕВРАЗ)",
  "Норильск": "⛓ «Норникель» — цветная металлургия",
  "Липецк": "⛓ НЛМК — чёрная металлургия",
  "Новокузнецк": "⛓ металлургия + уголь (Кузбасс)",
  "Череповец ": "⛓ металлургия",
  "Тольятти": "🚗 АвтоВАЗ, автопром, химия (КуйбышевАзот)",
  "Набережные Челны": "🚚 КамАЗ — грузовики",
  "Нижнекамск": "🧪 нефтехимия (Нижнекамскнефтехим, ТАНЕКО), шины",
  "Дзержинск": "🧪 крупнейший центр химпрома",
  "Кириши": "🛢 нефтепереработка (КИНЕФ)",
  "Ангарск": "🧪 нефтехимия, АЭХК",
  "Салават": "🧪 нефтехимия (Газпром нефтехим Салават)",
  "Стерлитамак": "🧪 химия (Сода, Каустик)",
  "Новомосковск": "🧪 химия (Азот, бытовая химия)",
  "Щёкино": "🧪 химия (Щёкиноазот)",
  "Кемерово": "🧪 химия (Азот) + уголь",
  "Березники": "🧪 калий (Уралкалий), содовая химия",
  "Соликамск": "🧪 калий (Уралкалий), целлюлоза",
  "Тобольск": "🧪 нефтехимия (СИБУР ЗапСибНефтехим)",
  "Омск": "🛢 нефтепереработка (Газпромнефть), нефтехимия",
  "Кстово": "🛢 нефтепереработка (ЛУКОЙЛ), нефтехимия",
  "Волжский": "🧪 химия, трубный и подшипниковый заводы",
  "Рыбинск": "✈ ОДК-Сатурн — авиадвигатели",
  "Тутаев": "⚙ моторный завод",
  "Комсомольск-на-Амуре": "✈ авиастроение (Су), судостроение",
  "Северодвинск": "🚢 Севмаш — атомные подлодки",
  "Миасс": "🚚 «Урал» (грузовики), ГРЦ Макеева (ракеты)",
  "Златоуст": "⛓ металлургия, оружие, часовой завод",
  "Каменск-Уральский": "⛓ алюминий, цветная металлургия",
  "Первоуральск": "⛓ трубный завод (ПНТЗ)",
  "Ковров": "🔫 оборонка (завод им. Дегтярёва)",
  "Тула": "🔫 оборонка, оружие, металлургия",
  "Выкса": "⛓ металлургия, трубы (ОМК)",
  "Старый Оскол": "⛓ металлургия (ОЭМК), руда КМА",
  "Губкин": "⛏ горнодобыча (Лебединский ГОК)",
  "Костомукша": "⛏ руда (Карельский окатыш)",
  "Апатиты": "⛏ апатит-нефелиновая руда (ФосАгро)",
  "Кировск": "⛏ апатиты (ФосАгро)",
  "Воркута": "⛏ угольный бассейн (Заполярье)",
  "Инта": "⛏ уголь",
  "Междуреченск": "⛏ уголь (Кузбасс)",
  "Прокопьевск": "⛏ уголь (Кузбасс)",
  "Ленинск-Кузнецкий": "⛏ уголь (Кузбасс)",
  "Белово": "⛏ уголь (Кузбасс)",
  "Киселёвск": "⛏ уголь (Кузбасс)",
  "Осинники": "⛏ уголь (Кузбасс)",
  "Анжеро-Судженск": "⛏ уголь (Кузбасс)",
  "Шахты": "⛏ уголь (Восточный Донбасс)",
  "Гуково": "⛏ уголь",
  "Новошахтинск": "⛏ уголь",
  "Нефтеюганск": "🛢 нефтедобыча (Юганскнефтегаз)",
  "Нижневартовск": "🛢 нефтедобыча (Самотлор)",
  "Сургут": "🛢 нефть и газ, крупнейшая ГРЭС",
  "Когалым": "🛢 нефтедобыча (ЛУКОЙЛ)",
  "Ноябрьск": "🛢 нефть и газ",
  "Новый Уренгой": "🔥 газовая столица (Уренгойгазпром)",
  "Надым": "🔥 добыча газа",
  "Губкинский": "🔥 нефть и газ",
  "Муравленко": "🛢 нефтедобыча",
  "Усинск": "🛢 нефтедобыча (Коми)",
  "Ухта": "🛢 нефть, газ, переработка",
  "Альметьевск": "🛢 нефть (столица Татнефти)",
  "Нефтекамск": "🛢 нефть, автозаводы (НефАЗ)",
  "Октябрьский": "🛢 нефтедобыча (Башкирия)",
  "Ноябрьск ": "🛢 нефть",
  "Рубцовск": "🚜 сельхозмашиностроение",
  "Северобайкальск": "🚆 БАМ — железнодорожный узел",
  "Тында": "🚆 столица БАМ",
  "Гаджиево": "🚢 ЗАТО, база подводного флота",
  "Североморск": "🚢 ЗАТО, главная база Северного флота",
  "Вилючинск": "🚢 ЗАТО, база подводного флота (Камчатка)",
  "Мирный": "🚀 ЗАТО, космодром Плесецк",
  "Знаменск": "🚀 ЗАТО, полигон Капустин Яр",
  "Краснознаменск": "🛰 ЗАТО, управление космическими аппаратами",
  "Байконур": "🚀 космодром Байконур",
  "Магадан": "⛏ золотодобыча (Колыма)",
  "Мирный ": "💎 добыча алмазов (АЛРОСА)",
  "Удачный": "💎 добыча алмазов (АЛРОСА)",
  "Нерюнгри": "⛏ уголь (Южная Якутия)",
  "Асбест": "⛏ добыча асбеста",
  "Качканар": "⛏ ГОК, ванадиевая руда",
  "Полевской": "⛓ металлургия, криолитовый завод",
  "Ревда": "⛓ цветная металлургия (СУМЗ)",
  "Верхняя Пышма": "⛓ медь (УГМК)",
  "Красноуральск": "⛓ медеплавильный завод",
  "Кировград": "⛓ медь",
  "Медногорск": "⛓ медно-серный комбинат",
  "Гай": "⛏ медная руда (Гайский ГОК)",
  "Сибай": "⛏ медно-цинковая руда",
  "Учалы": "⛏ горно-обогатительный комбинат",
  "Заринск": "🧪 коксохимия",
  "Мончегорск": "⛓ никель (Норникель, Кольская ГМК)",
  "Заполярный": "⛏ руда (Кольская ГМК)",
  "Кандалакша": "⛓ алюминиевый завод",
  "Волхов": "⛓ алюминий, химия",
  "Пикалёво": "🧪 глинозём, цемент, сода",
  "Бокситогорск": "⛏ бокситы, глинозём",
  "Шелехов": "⛓ алюминий (ИркАЗ)",
  "Саяногорск": "⛓ алюминий (СААЗ), Саяно-Шушенская ГЭС рядом",
  "Братск": "⛓ алюминий (БрАЗ), ГЭС, лесохимия",
  "Усть-Илимск": "🌲 лесопромышленный комплекс, ГЭС",
  "Красноярск": "⛓ алюминий (КрАЗ), машиностроение",
  "Ачинск": "⛓ глинозём, нефтепереработка",
  "Назарово": "⛏ уголь, ГРЭС",
  "Шарыпово": "⛏ уголь (Берёзовский разрез), ГРЭС",
  "Чайковский": "🧪 химия (Уралоргсинтез), текстиль",
  "Губаха": "🧪 коксохимия",
  "Кувандык": "⛓ криолитовый завод",
  "Новотроицк": "⛓ металлургия (Уральская Сталь)",
  "Орск": "⛓ металлургия, нефтепереработка",
  "Медногорск ": "⛓ медь",
  "Ясный": "⛏ асбест, ЗАТО ракетная дивизия рядом",
  "Балахна": "📄 целлюлозно-бумажный комбинат",
  "Кондопово": "📄 бумага",
  "Кондопога": "📄 крупный ЦБК (газетная бумага)",
  "Сегежа": "📄 целлюлозно-бумажный комбинат",
  "Коряжма": "📄 целлюлозно-бумажный комбинат",
  "Сокол": "🌲 деревообработка, ЦБК",
  "Светогорск": "📄 целлюлозно-бумажный комбинат",
  "Выборг": "🚢 судостроение, порт",
  "Тихвин": "🚆 вагоностроение (ТВСЗ)",
  "Гатчина": "☢ ПИЯФ — ядерная физика",
  "Протвино": "☢ наукоград, физика высоких энергий",
  "Пущино": "🔬 наукоград, биология",
  "Королёв": "🚀 наукоград, ракетно-космическая (РКК Энергия)",
  "Жуковский": "✈ наукоград, авиация (ЦАГИ, ЛИИ)",
  "Реутов": "🚀 ракетостроение (НПО машиностроения)",
  "Фрязино": "📡 наукоград, электроника, СВЧ",
  "Черноголовка": "🔬 наукоград, химическая физика",
  "Троицк": "🔬 наукоград, физика",
  "Электросталь": "⛓ металлургия, машиностроение, атомная (ТВЭЛ)",
  "Подольск": "⚙ машиностроение (котлы, ЗиО для АЭС)",
  "Коломна": "🚆 тепловозостроение, тяжёлое машиностроение",
  "Егорьевск": "🏭 машиностроение, текстиль",
  "Балашиха": "⚙ криогенное машиностроение",
  "Клин": "🧪 химволокно, стекло",
  "Воскресенск": "🧪 химия (минеральные удобрения)",
  "Ступино": "✈ авиационная металлургия",
  "Дзержинский": "🚀 ракетное топливо (ФЦДТ «Союз»)",
  "Москва": "🏛 столица РФ: финансы, наука, IT, штаб-квартиры",
  "Санкт-Петербург": "🚢 судостроение, тяжёлое машиностроение, порт",
  "Новосибирск": "🔬 наука (Академгородок), IT, транспортный узел",
  "Екатеринбург": "⚙ тяжёлое машиностроение, металлургия Урала",
  "Казань": "✈ авиастроение (КАЗ), нефтехимия, IT",
  "Нижний Новгород": "🚗 ГАЗ, судостроение, оборонка, IT",
  "Челябинск": "⛓ чёрная металлургия (ЧМК), трубы (ЧТПЗ)",
  "Самара": "🚀 ракетно-космическая (РКЦ Прогресс), авиация",
  "Уфа": "🛢 нефтепереработка, нефтехимия, моторы (УМПО)",
  "Ростов-на-Дону": "🚜 сельхозмашиностроение (Ростсельмаш), вертолёты",
  "Краснодар": "🌾 АПК и пищепром, столица Кубани",
  "Воронеж": "✈ авиастроение (ВАСО), радиоэлектроника, ракетные двигатели",
  "Пермь": "🚀 авиадвигатели (ОДК-Пермь), пороха, машиностроение",
  "Волгоград": "⛓ металлургия, тракторный завод, химия",
  "Саратов": "✈ авиация, химия, нефтепереработка",
  "Тюмень": "🛢 нефтегазовый и нефтесервисный центр",
  "Барнаул": "⚙ машиностроение, коксохимия, АПК",
  "Ижевск": "🔫 оружие (Калашников), автопром, металлургия",
  "Махачкала": "🐟 рыба, морской порт, пищепром",
  "Хабаровск": "⚙ машиностроение, НПЗ, транспортный узел",
  "Ульяновск": "✈ авиастроение (Авиастар), автопром (УАЗ)",
  "Иркутск": "✈ авиастроение (Иркут), ГЭС, транспортный узел",
  "Владивосток": "🚢 морской порт, рыбная отрасль, судоремонт",
  "Ярославль": "🛢 нефтепереработка, моторы (ЯМЗ), шины",
  "Томск": "🔬 наука и университеты, нефтехимия",
  "Севастополь": "⚓ база Черноморского флота, судоремонт",
  "Ставрополь": "🌾 АПК, химия, приборостроение",
  "Оренбург": "🔥 газ (Оренбурггазпром), нефть",
  "Рязань": "🛢 нефтепереработка, приборостроение",
  "Пенза": "⚙ приборостроение, машиностроение",
  "Чебоксары": "⚙ электротехника (ABB), тракторы, ХБК",
  "Калининград": "🚢 судостроение, янтарь, рыба, порт",
  "Астрахань": "🐟 рыба, судостроение, газоперереработка",
  "Сочи": "🏝 курорт и туризм",
  "Курск": "⚙ электроаппаратура, АПК, АЭС рядом",
  "Улан-Удэ": "🚁 вертолётный завод (У-УАЗ), локомотивы",
  "Тверь": "🚆 вагоностроение (ТВЗ), экскаваторы",
  "Брянск": "🚆 машиностроение (БМЗ), оборонка",
  "Иваново": "🧵 текстиль (текстильная столица)",
  "Якутск": "💎 центр алмазного края, СВФУ",
  "Владимир": "⚙ машиностроение, тракторы, стекло",
  "Симферополь": "🏛 столица Крыма, пищепром, транспорт",
  "Белгород": "⛏ руда КМА рядом, стройматериалы, АПК",
  "Калуга": "🚗 автомобильный кластер (VW/PSA/Volvo)",
  "Чита": "🚆 ж/д узел, горнодобыча",
  "Грозный": "🛢 историческая нефтяная столица, нефтехимия",
  "Смоленск": "💎 огранка бриллиантов (Кристалл), авиазавод",
  "Саранск": "💡 светотехника (Лисма), кабель, вагоны",
  "Вологда": "🧈 маслоделие, подшипники, машиностроение",
  "Курган": "🚌 автобусы (КАвЗ), сельхозмаш, арматура",
  "Орёл": "⚙ приборостроение, сталепрокат",
  "Архангельск": "🌲 ЦБК, лес, порт, судостроение",
  "Владикавказ": "⛓ цветная металлургия, «Победит»",
  "Йошкар-Ола": "⚙ приборостроение, оборонка",
  "Мурманск": "🐟 рыбный порт, атомный ледокольный флот",
  "Кострома": "💍 ювелирная столица, ЦБК, лён",
  "Новороссийск": "🚢 крупнейший порт, цемент",
  "Тамбов": "⚙ химия, машиностроение, АПК",
  "Химки": "🚀 ракетно-космическая (НПО Лавочкина, Энергомаш)",
  "Мытищи": "🚆 вагоностроение (Метровагонмаш), ПВО (Алмаз-Антей)",
  "Нальчик": "⛓ цветная металлургия, приборостроение, курорт",
  "Таганрог": "✈ авиастроение (ТАНТК Бериева), металлургия (ТМК)",
  "Петрозаводск": "🌲 ЦБК, тракторы (ОТЗ), лес",
  "Энгельс": "🚚 троллейбусы, химия, авиабаза",
  "Великий Новгород": "🧪 химия («Акрон»), электроника",
  "Псков": "⚙ электротехника, АПК",
  "Красногорск": "📷 оптика (КМЗ, «Зенит»)",
  "Армавир": "🚆 машиностроение, ж/д, АПК",
  "Абакан": "🚆 ж/д узел, контейнеры, лёгкая промышленность",
  "Бийск": "🧪 наукоград, химия (пороха), фармацевтика",
  "Южно-Сахалинск": "🛢 нефтегаз (Сахалин), рыба",
  "Уссурийск": "🌾 АПК, сахар, локомотиворемонт",
  "Сызрань": "🛢 нефтепереработка, тяжмаш",
  "Петропавловск-Камчатский": "🐟 рыбная отрасль, морской порт",
  "Новочеркасск": "🚆 электровозы (НЭВЗ), ГРЭС",
  "Керчь": "🚢 судостроение (Залив), порт, металлургия",
  "Копейск": "⛏ уголь, машиностроение",
  "Пятигорск": "🏝 курорт Кавминвод, туризм",
  "Майкоп": "🌾 АПК, машиностроение",
  "Находка": "🚢 порт, рыба, нефтеналив",
  "Домодедово": "✈ крупный аэропорт",
  "Гусь-Хрустальный": "🔮 стекло и хрусталь",
  "Павлово": "🔧 металлоизделия, автобусы (ПАЗ)",
  "Арзамас": "⚙ приборостроение, оборонка",
  "Углич": "⌚ часы («Чайка»), сыроделие",
  "Кольчугино": "⛓ цветной прокат, кабель",
  "Ирбит": "🏍 мотоциклы («Урал»)",
  "Сатка": "🧱 магнезит («Магнезит»)",
  "Кыштым": "⛓ медеэлектролит, каолин",
  "Карабаш": "⛓ медеплавильный завод",
  "Аша": "⛓ металлургия",
  "Юрга": "⚙ машиностроение (Юрмаш)",
  "Новоалтайск": "🚆 вагоностроение",
  "Черногорск": "⛏ уголь",
  "Тавда": "🌲 лес, фанера",
  "Сухой Лог": "🧱 цемент",
  "Богданович": "🧱 огнеупоры",
  "Кулебаки": "⛓ металлургия",
  "Навашино": "🚢 судостроение",
  "Заволжье": "🚗 моторы (ЗМЗ)",
  "Бор": "🪟 стекло",
  "Гаврилов-Ям": "🧵 льнокомбинат",
  "Кимры": "👞 обувь (историческая столица)",
  "Дятьково": "🔮 хрусталь",
  "Клинцы": "🧵 текстиль, машиностроение",
  "Вязьма": "🚆 ж/д узел, машиностроение",
  "Рославль": "⚙ автоагрегаты, ж/д",
  "Сафоново": "⚙ приборы, пластмассы",
  "Мичуринск": "🍎 наукоград, садоводство",
  "Котлас": "🚆 ж/д узел, ЦБК рядом",
  "Великие Луки": "⚙ машиностроение, электротехника",
  "Прохладный": "🚆 ж/д узел, кабель",
  "Черкесск": "🧪 химия (резинотехника), АПК",
  "Элиста": "🏛 столица Калмыкии, буддийский центр",
  "Кызыл": "🏛 столица Тывы, центр Азии",
  "Горно-Алтайск": "🏔 столица Республики Алтай, туризм",
  "Магас": "🏛 молодая столица Ингушетии",
  "Анадырь": "🐟 рыба, порт, столица Чукотки",
  "Нарьян-Мар": "🛢 нефтедобыча (НАО), порт",
  "Салехард": "🐟 рыба, столица Ямала, полярный круг",
  "Ханты-Мансийск": "🛢 нефтяная столица Югры",
  "Биробиджан": "⚙ сельхозмаш, лёгкая промышленность",
  "Кудымкар": "🌲 лес, АПК",
  "Воткинск": "🚀 ракетостроение (Воткинский завод), станки",
  "Глазов": "☢ производство циркония (ЧМЗ, ТВЭЛ)",
  "Сарапул": "📻 радиозавод, электроника",
  "Балашов": "🧪 комбинат плащевых тканей, АПК",
  "Россошь": "🧪 минеральные удобрения",
  "Борисоглебск": "⚙ химмаш, АПК",
  "Мелеуз": "🧪 минеральные удобрения",
  "Кумертау": "🚁 вертолётостроение (КумАПП)",
  "Белорецк": "⛓ металлургия, метизы",
  "Ишимбай": "🛢 нефтедобыча, машиностроение",
  "Туймазы": "🛢 нефтедобыча, автобетоновозы",
  "Волжск": "📄 целлюлозно-бумажный, машиностроение",
  "Канаш": "🚆 вагоноремонт, автоагрегаты",
  "Шумерля": "🚗 автофургоны, химия",
  "Алатырь": "🚆 приборостроение, ж/д",
  "Сердобск": "⚙ электромеханика, часы",
  "Кузнецк": "👞 обувь, кожевенное производство",
  "Кинешма": "🧵 текстиль, химия",
  "Шуя": "🧵 текстиль, гармони, водка",
  "Вичуга": "🧵 текстиль",
  "Муром": "🚆 тепловозостроение, радиозавод",
  "Ковылкино": "🚆 ж/д, электромеханика",
  "Рузаевка": "🚆 крупный ж/д узел, вагоны",
  "Каменск-Шахтинский": "🧪 химволокно, металлоизделия",
  "Донецк": "⛏ уголь (Ростовская обл.)",
  "Зверево": "⛏ уголь",
  "Красный Сулин": "⛓ металлургия",
  "Волгореченск": "⚡ ГРЭС, трубный завод",
  "Буй": "🧪 химия (удобрения), ж/д узел",
  "Мантурово": "🌲 деревообработка",
  "Шарья": "🌲 деревообработка (фанера, плиты)",
  "Няндома": "🚆 ж/д узел",
  "Онега": "🌲 лес, ЦБК, порт",
  "Вельск": "🌲 лес, АПК",
  "Печора": "⛏ уголь, ж/д, ГРЭС",
  "Сосногорск": "🛢 газопереработка, сажевый завод",
  "Емва": "🌲 лес, ж/д",
  "Богородицк": "🔬 электрооптика (БЗТХИ, кристаллы для ЦЕРН), дворец Бобринских",
  "Донской": "⛏ бывшая угледобыча Подмосковного бассейна, стройматериалы",
  "Узловая": "🚗 автозавод «Хавейл», ж/д узел, ОЭЗ «Узловая»",
  "Алексин": "🧪 химия (пороха), стройматериалы, ГРЭС",
  "Ефремов": "🧪 химия (синтетический каучук), АПК",
  "Кимовск": "⛏ бывшая угледобыча, радиоэлектромеханика",
  "Киреевск": "⛓ лёгкие металлоконструкции, стройматериалы",
  "Липки": "⛏ бывшая угледобыча",
  "Болохово": "⚙ машиностроение, стройматериалы",
  "Люберцы": "⚙ вертолётные агрегаты, машиностроение",
  "Одинцово": "🏙 деловой и жилой спутник Москвы",
  "Щёлково": "🧪 химия и биотехнологии, авиабаза Чкаловский",
  "Серпухов": "⚙ приборостроение, мотозавод, конденсаторы",
  "Раменское": "📡 авионика и приборостроение (РПКБ)",
  "Пушкино": "🧵 текстиль, деревообработка",
  "Орехово-Зуево": "🧵 текстиль (Морозовские мануфактуры), машиностроение",
  "Ногинск": "🧵 текстиль, машиностроение",
  "Сергиев Посад": "⛪ Троице-Сергиева лавра, игрушки, оптика (ЗОМЗ)",
  "Видное": "🧪 коксогазовый завод, металлоконструкции",
  "Чехов": "📄 полиграфкомбинат, мебель",
  "Ивантеевка": "🧵 текстиль, машиностроение",
  "Лобня": "⚙ электротехника, стройматериалы",
  "Наро-Фоминск": "🧵 текстиль (шёлк), приборостроение",
  "Дмитров": "⚙ мостовые конструкции, экскаваторы, фрезы",
  "Павловский Посад": "🧣 павловопосадские платки и шали",
  "Лыткарино": "🔭 оптическое стекло (ЛЗОС), авиадвигатели",
  "Солнечногорск": "⚙ металлические сетки, стекловолокно, приборы",
  "Долгопрудный": "🚀 наука (МФТИ), НПО (ДНПП)",
  "Дедовск": "🧵 текстиль",
  "Электрогорск": "⚡ ГРЭС, приборостроение",
  "Луховицы": "✈ авиастроение (МиГ), овощеводство",
  "Кашира": "⚡ Каширская ГРЭС, металлоконструкции",
  "Можайск": "📄 полиграфия, медико-инструментальный завод",
  "Волоколамск": "🌾 АПК, ткацкое производство",
  "Зарайск": "👞 обувь, перо-пуховые изделия",
  "Озёры": "🧵 текстиль",
  "Хотьково": "🎨 художественные промыслы, НИИ материалов",
  "Куровское": "🧵 текстиль",
  "Старая Купавна": "🧪 химия и фармацевтика («Акрихин»)",
  "Ликино-Дулёво": "🚌 автобусы (ЛиАЗ), фарфор (Дулёво)",
  "Апрелевка": "💿 бывший завод грампластинок, стройматериалы",
  "Голицыно": "🏭 керамика, стройматериалы, приборы",
  "Кубинка": "✈ авиабаза, парк «Патриот» (танковый музей)",
  "Бронницы": "💍 ювелирный завод, машиностроение",
  "Рошаль": "🧪 химкомбинат (пороха, пластмассы)",
  "Шатура": "⚡ ГРЭС на торфе, мебель («Шатура»)",
  "Лосино-Петровский": "🧵 камвольный комбинат",
  "Звенигород": "🏝 санатории, туризм",
  "Истра": "🔬 приборостроение, стройматериалы",
  "Котельники": "🏙 логистика, торговля (спутник Москвы)",
  "Высоковск": "🧵 текстиль",
  "Краснозаводск": "🎆 химзавод (пиротехника), ХБК",
  "Пересвет": "🚀 испытания ракетных двигателей (НИЦ РКП)",
  "Всеволожск": "🚗 автосборка, пищепром (спутник СПб)",
  "Кингисепп": "🧪 минеральные удобрения («ЕвроХим»), порт Усть-Луга рядом",
  "Тосно": "🏭 стройматериалы, пищепром, логистика",
  "Луга": "🏭 абразивы, стройматериалы, АПК",
  "Сланцы": "⛏ сланцы, цемент",
  "Отрадное": "🚢 судостроение (Пелла), машиностроение",
  "Коммунар": "📄 бумажная фабрика",
  "Никольское": "🧱 кирпич, стройматериалы",
  "Сясьстрой": "📄 целлюлозно-бумажный комбинат",
  "Лодейное Поле": "🌲 деревообработка, ж/д",
  "Любань": "🚆 ж/д узел",
  "Боровичи": "🧱 огнеупоры, комбинат «Боровичи»",
  "Старая Русса": "🏝 бальнеокурорт, приборостроение, авиаремонт",
  "Остров": "⚡ электромашиностроение, АПК",
  "Новодвинск": "📄 Архангельский ЦБК",
  "Оленегорск": "⛏ железорудный ГОК",
  "Заозёрск": "🚢 ЗАТО, база Северного флота",
  "Кола": "🐟 рыба, стройматериалы",
  "Елец": "⚙ машиностроение, кружево, табак, ж/д узел",
  "Ливны": "⚙ насосы, гидромашины, АПК",
  "Мценск": "⚙ литейное производство (цветмет), приборы",
  "Ржев": "⚙ краностроение, приборы, ж/д узел",
  "Вышний Волочёк": "🧵 текстиль (хлопок), стекло",
  "Торжок": "✈ вертолётный центр, золотное шитьё, вагоностроение",
  "Бежецк": "⚙ автоспецоборудование, льнообработка",
  "Конаково": "⚡ Конаковская ГРЭС, фаянс",
  "Бологое": "🚆 крупный ж/д узел (Москва–СПб)",
  "Гагарин": "🍬 пищепром, светотехника, родина Ю. Гагарина",
  "Ярцево": "🧵 хлопчатобумажный комбинат, литьё",
  "Дорогобуж": "🧪 минеральные удобрения",
  "Новозыбков": "⚙ станки, швейное производство",
  "Унеча": "🚆 ж/д узел, нефтеперекачка",
  "Сельцо": "🧨 оборонный завод (боеприпасы)",
  "Жуковка": "⚙ веломотозавод, деревообработка",
  "Моршанск": "🧵 суконная фабрика, табак, машиностроение",
  "Рассказово": "🧵 текстиль, кожевенное производство",
  "Котовск": "🧨 пороховой завод, лакокраска",
  "Уварово": "🧪 химия (бывший химзавод), АПК",
  "Александров": "📺 радиотехника (телевизоры «Рекорд»), ж/д",
  "Вязники": "🧵 льнокомбинаты, текстиль",
  "Киржач": "💡 светотехника, инструменты, шёлк",
  "Собинка": "🧵 хлопчатобумажный комбинат",
  "Лакинск": "🧵 прядильно-ткацкая фабрика",
  "Карабаново": "🧵 текстиль",
  "Струнино": "🧵 текстиль",
  "Камешково": "🧵 текстиль",
  "Меленки": "🧵 текстиль, литьё",
  "Судогда": "🧵 стекловолокно",
  "Костерёво": "🏭 стройматериалы, пластмассы",
  "Курлово": "🔮 стекольный завод",
  "Гороховец": "🚢 судостроение, туризм",
  "Тейково": "🧵 хлопчатобумажный комбинат, РВСН",
  "Родники": "🧵 текстиль (меланжевый комбинат)",
  "Фурманов": "🧵 текстиль",
  "Приволжск": "💍 ювелирное производство, текстиль",
  "Наволоки": "🧵 хлопчатобумажный комбинат",
  "Кохма": "🧵 текстиль",
  "Ростов": "🎨 финифть, эмаль, туризм (кремль)",
  "Переславль-Залесский": "🎨 туризм, вышивка, химия (магнитная лента)",
  "Данилов": "🚆 ж/д узел, сыроделие",
  "Пошехонье": "🧀 сыроделие (пошехонский сыр)",
  "Нерехта": "🏭 машиностроение, текстиль",
  "Галич": "⚙ автокрановый завод",
  "Чухлома": "🌲 лесное хозяйство",
  "Великий Устюг": "🎅 родина Деда Мороза, кисти, туризм",
  "Нея": "🌲 деревообработка",
  "Зеленодольск": "🚢 судостроение (завод им. Горького), фанера",
  "Чистополь": "⌚ часовой завод («Восток»), судостроение",
  "Елабуга": "🚗 ОЭЗ «Алабуга», автопром, нефтехимия",
  "Бугульма": "🛢 нефтедобыча, машиностроение",
  "Лениногорск": "🛢 нефтедобыча",
  "Азнакаево": "🛢 нефтедобыча",
  "Заинск": "⚡ Заинская ГРЭС, автозапчасти",
  "Нурлат": "🛢 нефтедобыча, сахар",
  "Менделеевск": "🧪 химия (удобрения)",
  "Можга": "🔮 стекольный завод, деревообработка",
  "Агрыз": "🚆 ж/д узел",
  "Буинск": "🍬 сахарный завод, АПК",
  "Бавлы": "🛢 нефтедобыча",
  "Арск": "🧵 меховое и валяльное производство",
  "Козловка": "⚙ комбинат автофургонов",
  "Мариинский Посад": "🌲 деревообработка, судостроение",
  "Ядрин": "🍬 пищепром",
  "Козьмодемьянск": "🌲 деревообработка, судоремонт",
  "Звенигово": "🚢 судостроение и судоремонт",
  "Новокуйбышевск": "🛢 нефтепереработка, нефтехимия",
  "Чапаевск": "🧨 химия (боеприпасы, азот)",
  "Жигулёвск": "⚡ Жигулёвская ГЭС, стройматериалы",
  "Октябрьск": "🚆 ж/д узел на Волге",
  "Кинель": "🚆 ж/д узел, АПК",
  "Отрадный": "🛢 нефтедобыча, газопереработка",
  "Похвистнево": "🛢 нефтедобыча",
  "Нефтегорск": "🛢 нефтедобыча",
  "Бузулук": "🛢 нефтедобыча, тяжмаш",
  "Бугуруслан": "🛢 нефтедобыча (старейшая в регионе)",
  "Сорочинск": "🛢 нефтепереработка, АПК",
  "Соль-Илецк": "🧂 добыча соли, бальнеокурорт, арбузы",
  "Абдулино": "🚆 ж/д узел",
  "Вольск": "🧱 цемент",
  "Пугачёв": "🌾 АПК, механический завод",
  "Маркс": "⚙ дизельная аппаратура, АПК",
  "Аткарск": "🌾 маслоэкстракционный завод, АПК",
  "Петровск": "⚙ автозапчасти, АПК",
  "Ртищево": "🚆 крупный ж/д узел",
  "Нижний Ломов": "🥫 пищепром (консервы, спички), фанера",
  "Каменка": "🌾 АПК, машиностроение",
  "Сурск": "🧵 текстиль",
  "Городище": "🌾 АПК",
  "Богородск": "👞 кожевенное производство",
  "Городец": "🎨 городецкая роспись, судостроение, туризм",
  "Семёнов": "🎨 хохломская роспись, матрёшки",
  "Лысково": "⚙ электротехника, металлоизделия",
  "Ворсма": "🔪 ножи и медицинский инструмент",
  "Урень": "🌲 деревообработка",
  "Сергач": "🍬 сахарный завод",
  "Княгинино": "👗 швейное производство",
  "Лукоянов": "🌾 АПК",
  "Горбатов": "🎣 рыболовные снасти",
  "Грязи": "🚆 ж/д узел, пищепром",
  "Данков": "🧪 химия (доломит, силикатная)",
  "Лебедянь": "🥫 соки («Лебедянский»), машиностроение",
  "Усмань": "🌾 АПК, табак, литьё",
  "Чаплыгин": "⚙ агрегатный завод",
  "Кирсанов": "🍬 сахар, АПК",
  "Лиски": "🚆 крупный ж/д узел, пищепром",
  "Острогожск": "🌾 АПК, консервы",
  "Семилуки": "🧱 огнеупоры",
  "Бобров": "🌾 АПК, сыроделие",
  "Бутурлиновка": "🍞 пищепром, ликёро-водочный",
  "Калач": "🌾 АПК, сыроделие",
  "Шебекино": "🧪 бытовая химия (ПАВ), машиностроение",
  "Алексеевка": "🌻 подсолнечное масло (родина маргарина), АПК",
  "Валуйки": "🚆 ж/д узел, АПК",
  "Строитель": "🏭 стройматериалы, АПК",
  "Бирюч": "🌾 АПК",
  "Льгов": "🍬 сахар, АПК",
  "Щигры": "⚙ геологоразведочное оборудование",
  "Обоянь": "🌾 АПК, консервы",
  "Батайск": "🚆 ж/д узел, машиностроение",
  "Азов": "🚢 порт, машиностроение (кузнечно-прессовое)",
  "Аксай": "🏭 стекло, логистика",
  "Сальск": "🌾 АПК, машиностроение",
  "Миллерово": "🌻 маслоэкстракция, АПК",
  "Белая Калитва": "⛓ алюминиевый прокат (металлургия)",
  "Морозовск": "🌾 АПК, авиабаза",
  "Семикаракорск": "🎨 семикаракорская керамика, АПК",
  "Константиновск": "🌾 АПК, порт на Дону",
  "Зерноград": "🚜 сельхознаука и АПК",
  "Пролетарск": "🌾 АПК",
  "Ейск": "🏝 приморский курорт, порт, машиностроение",
  "Анапа": "🏝 детский черноморский курорт",
  "Геленджик": "🏝 черноморский курорт",
  "Туапсе": "🛢 порт, нефтепереработка (нефтеналив)",
  "Кропоткин": "🚆 ж/д узел, химмаш, пищепром",
  "Славянск-на-Кубани": "🌾 АПК (рис), нефтегаз",
  "Тихорецк": "🚆 ж/д машиностроение, пищепром",
  "Крымск": "🌾 консервы, АПК, ж/д узел",
  "Тимашёвск": "🥫 пищепром (сахар, мясо, «Нестле»)",
  "Лабинск": "🌾 АПК, сахар, машиностроение",
  "Курганинск": "🌾 сахар, АПК",
  "Белореченск": "🧪 минеральные удобрения",
  "Апшеронск": "🌲 деревообработка, курорт",
  "Усть-Лабинск": "🌾 АПК, сахар",
  "Кореновск": "🥛 молочный комбинат, сахар, авиабаза",
  "Темрюк": "🍇 виноделие, порт (Тамань)",
  "Абинск": "⛓ электрометаллургический завод",
  "Горячий Ключ": "🏝 бальнеокурорт",
  "Новокубанск": "🍾 коньячный завод, АПК",
  "Хадыженск": "🛢 нефтедобыча, бальнеокурорт",
  "Гулькевичи": "🌾 сахар, АПК",
  "Приморско-Ахтарск": "🐟 рыба, курорт",
  "Кисловодск": "🏝 бальнеокурорт (нарзан)",
  "Ессентуки": "🏝 бальнеокурорт (минеральные воды)",
  "Железноводск": "🏝 бальнеокурорт",
  "Лермонтов": "⛏ бывшая добыча урана, химия",
  "Минеральные Воды": "✈ аэропорт, ж/д узел КМВ",
  "Невинномысск": "🧪 химия («Невинномысский Азот»), ГРЭС",
  "Изобильный": "🔥 газоперекачка, ГРЭС",
  "Будённовск": "🧪 нефтехимия («Ставролен»)",
  "Георгиевск": "🌾 АПК, машиностроение",
  "Светлоград": "🌾 АПК, газ",
  "Зеленокумск": "🌾 АПК",
  "Благодарный": "🌾 АПК",
  "Ипатово": "🌾 АПК",
  "Новоалександровск": "🌾 АПК",
  "Новопавловск": "🌾 АПК",
  "Нефтекумск": "🛢 нефтедобыча",
  "Хасавюрт": "🐑 АПК, ковроткачество, торговля",
  "Дербент": "🍇 виноделие и коньяк, туризм (крепость)",
  "Каспийск": "⚙ оборонный завод («Дагдизель»), порт",
  "Буйнакск": "⚙ приборостроение, агрегатный завод",
  "Избербаш": "🛢 нефтедобыча, приборостроение",
  "Кизляр": "🍾 коньячный завод, электроаппаратура, оружие (ножи)",
  "Кизилюрт": "⚡ ГЭС, приборостроение",
  "Дагестанские Огни": "🔮 стекло",
  "Назрань": "🏭 электромеханика, АПК",
  "Малгобек": "🛢 нефтедобыча",
  "Карабулак": "🏭 АПК, стройматериалы",
  "Моздок": "🚆 ж/д узел, АПК, авиабаза",
  "Беслан": "🌽 кукурузно-крахмальный комбинат, ж/д узел",
  "Алагир": "⛓ цветная металлургия, ГЭС",
  "Ардон": "🌾 АПК",
  "Дигора": "🌾 АПК",
  "Баксан": "⚡ каскад ГЭС, АПК",
  "Нарткала": "🍬 сахар, химия",
  "Майский": "🌾 АПК, ж/д",
  "Терек": "⚙ алмазный инструмент, электроаппаратура",
  "Тырныауз": "⛏ вольфрамо-молибденовый комбинат",
  "Чегем": "🌾 АПК, стройматериалы",
  "Усть-Джегута": "🧱 цемент",
  "Карачаевск": "🏔 туризм, приборостроение",
  "Теберда": "🏝 горный курорт",
  "Серов": "⛓ металлургия (сталь, ферросплавы)",
  "Краснотурьинск": "⛓ алюминиевый завод (БАЗ)",
  "Карпинск": "⛏ бывшая угледобыча, машиностроение",
  "Североуральск": "⛏ бокситы (СУБР)",
  "Реж": "⛓ никелевый завод, химия",
  "Невьянск": "⛓ цемент, машиностроение, туризм (башня)",
  "Алапаевск": "⛓ металлургия, деревообработка",
  "Артёмовский": "⛏ бывшая угледобыча, машиностроение",
  "Кушва": "⛏ железная руда, металлургия",
  "Красноуфимск": "🚆 ж/д узел, АПК",
  "Камышлов": "🧵 кожевенно-текстильное, стройматериалы",
  "Сысерть": "🏭 гидромашины, фарфор, туризм",
  "Арамиль": "🧵 текстиль, логистика",
  "Среднеуральск": "⚡ Среднеуральская ГРЭС",
  "Верхняя Салда": "✈ титан (ВСМПО-АВИСМА)",
  "Дегтярск": "⛏ бывшая медная руда",
  "Верхний Уфалей": "⛓ никель",
  "Чебаркуль": "⛓ металлургический завод (поковки)",
  "Коркино": "⛏ угольный разрез",
  "Еманжелинск": "⛏ бывшая угледобыча, машиностроение",
  "Южноуральск": "⚡ ГРЭС, изоляторный завод",
  "Карталы": "🚆 крупный ж/д узел",
  "Усть-Катав": "🚋 вагоностроение (трамваи)",
  "Юрюзань": "⚙ механический завод",
  "Миньяр": "⚙ метизно-металлургическое",
  "Бакал": "⛏ железная руда",
  "Куса": "⚙ художественное литьё, машиностроение",
  "Пласт": "⛏ золотодобыча",
  "Катав-Ивановск": "🧱 цемент, приборостроение",
  "Нязепетровск": "⚙ краностроение",
  "Шадринск": "⚙ автоагрегаты, телефонный завод, АПК",
  "Далматово": "🌾 АПК, консервы, монастырь",
  "Куртамыш": "🌾 АПК",
  "Ишим": "🚆 ж/д узел, АПК, машиностроение",
  "Ялуторовск": "🥛 пищепром (молоко, лес), нефтедобыча",
  "Заводоуковск": "🌾 АПК, машиностроение",
  "Нягань": "🛢 нефтедобыча, ГРЭС",
  "Мегион": "🛢 нефтедобыча",
  "Урай": "🛢 нефтедобыча (первая в Зап. Сибири)",
  "Лангепас": "🛢 нефтедобыча (ЛУКОЙЛ)",
  "Лянтор": "🛢 нефтедобыча",
  "Пыть-Ях": "🛢 нефтедобыча",
  "Югорск": "🔥 транспорт газа («Газпром трансгаз Югорск»)",
  "Советский": "🌲 лес, газотранспорт",
  "Белоярский": "🔥 газотранспорт",
  "Стрежевой": "🛢 нефтедобыча (Томскнефть)",
  "Тарко-Сале": "🔥 нефть и газ",
  "Лабытнанги": "🚆 ж/д, ворота Ямала",
  "Бердск": "📡 электроника, приборостроение, пищепром",
  "Искитим": "🧱 цемент",
  "Куйбышев": "🧪 химия, спирт (быв. Каинск)",
  "Барабинск": "🚆 ж/д узел, ГРЭС",
  "Татарск": "🚆 ж/д узел, пищепром",
  "Карасук": "🚆 ж/д узел, пищепром",
  "Обь": "✈ пристанционный (аэропорт Толмачёво)",
  "Черепаново": "🏭 стройматериалы, АПК",
  "Тогучин": "🌾 АПК, лес",
  "Каргат": "🌾 АПК",
  "Канск": "🌲 ЦБК, деревообработка, ГРЭС",
  "Минусинск": "🍅 АПК (овощи), электротехника",
  "Лесосибирск": "🌲 лесопиление, ЛДК, канифоль",
  "Дивногорск": "⚡ Красноярская ГЭС, низковольтная аппаратура",
  "Сосновоборск": "🚜 бывший завод прицепов, машиностроение",
  "Дудинка": "🚢 морской порт Норильска (Енисей)",
  "Заозёрный": "⛏ слюда, уголь",
  "Бородино": "⛏ крупнейший угольный разрез",
  "Уяр": "🚆 ж/д узел",
  "Артёмовск": "⛏ золотодобыча",
  "Черемхово": "⛏ угледобыча, машиностроение",
  "Усолье-Сибирское": "🧪 химия (хлор, соль), фармацевтика",
  "Усть-Кут": "🚢 порт Осетрово, начало БАМ",
  "Тайшет": "⛓ алюминиевый завод, ж/д узел БАМ",
  "Тулун": "⛏ угольный разрез, гидролизный",
  "Саянск": "🧪 химия («Саянскхимпласт»)",
  "Зима": "🧪 химия, ж/д узел",
  "Нижнеудинск": "🚆 ж/д узел, слюда",
  "Вихоревка": "🚆 ж/д узел БАМ",
  "Железногорск-Илимский": "⛏ железорудный ГОК",
  "Бирюсинск": "🌲 гидролизный, деревообработка",
  "Байкальск": "📄 бывший БЦБК, туризм (Байкал)",
  "Свирск": "⚙ аккумуляторы, металлообработка",
  "Алзамай": "🌲 деревообработка",
  "Мыски": "⛏ уголь (разрез), ГРЭС",
  "Топки": "🧱 цемент",
  "Мариинск": "🍾 спирт, АПК, лесопиление",
  "Тайга": "🚆 ж/д узел",
  "Салаир": "⛏ полиметаллическая руда",
  "Таштагол": "⛏ железная руда",
  "Калтан": "⚡ Южно-Кузбасская ГРЭС",
  "Полысаево": "⛏ угледобыча",
  "Камень-на-Оби": "🌾 АПК, порт, деревообработка",
  "Славгород": "🧪 химия (соли), АПК",
  "Алейск": "🌾 АПК, мукомольное",
  "Яровое": "🧪 химический завод",
  "Змеиногорск": "⛏ полиметаллическая руда",
  "Горняк": "⛏ полиметаллическая руда",
  "Артём": "⛏ бывшая угледобыча, ГРЭС, аэропорт Владивостока",
  "Арсеньев": "✈ авиастроение («Прогресс», вертолёты Ка), судомашины",
  "Большой Камень": "🚢 судостроение (ССК «Звезда»)",
  "Спасск-Дальний": "🧱 цемент",
  "Партизанск": "⛏ бывшая угледобыча, ГРЭС",
  "Дальнегорск": "⛏ бор и полиметаллы (химкомбинат)",
  "Дальнереченск": "🌲 деревообработка, ж/д",
  "Лесозаводск": "🌲 деревообработка",
  "Свободный": "🚀 космодром Восточный рядом, ж/д, газохимия (Амурский ГПЗ)",
  "Зея": "⚡ Зейская ГЭС",
  "Райчихинск": "⛏ угледобыча",
  "Шимановск": "🚆 ж/д, машиностроение",
  "Циолковский": "🚀 космодром Восточный (ЗАТО)",
  "Амурск": "📄 бывший ЦКК, деревообработка, машиностроение",
  "Советская Гавань": "🚢 судоремонт, порт, рыба",
  "Николаевск-на-Амуре": "🐟 рыба, судоремонт, порт",
  "Вяземский": "🌲 деревообработка, ж/д",
  "Корсаков": "🐟 рыба, порт, СПГ рядом",
  "Холмск": "🐟 рыба, порт (паром на материк)",
  "Оха": "🛢 нефтедобыча (север Сахалина)",
  "Поронайск": "🐟 рыба, ЦБК",
  "Долинск": "🌾 АПК, пищепром",
  "Анива": "🌾 АПК, СПГ-завод рядом",
  "Невельск": "🐟 рыба, порт",
  "Елизово": "✈ аэропорт Камчатки, АПК",
  "Алдан": "⛏ золотодобыча",
  "Ленск": "🛢 нефтегаз, порт (АЛРОСА, «Сила Сибири»)",
  "Томмот": "🚆 ж/д (АЯМ), золото",
  "Покровск": "🏭 стройматериалы, газ",
  "Нюрба": "💎 алмазодобыча (АЛРОСА)",
  "Певек": "🚢 самый северный порт, плавучая АЭС",
  "Билибино": "☢ Билибинская АЭС, золото",
  "Камышин": "🧵 хлопчатобумажный комбинат, стеклотара, машиностроение",
  "Михайловка": "🧱 цемент, АПК",
  "Фролово": "🛢 нефтегаз, сталелитейный, АПК",
  "Урюпинск": "🐐 пуховязание (столица провинции), АПК",
  "Котельниково": "🔥 газоперекачка, АПК",
  "Суровикино": "🌾 АПК",
  "Калач-на-Дону": "🚢 порт (Волго-Дон), судоремонт",
  "Дубовка": "🍇 АПК, виноградарство",
  "Палласовка": "🌾 АПК, ж/д",
  "Петров Вал": "🚆 ж/д узел",
  "Котово": "🛢 нефтедобыча",
  "Евпатория": "🏝 детский приморский курорт",
  "Ялта": "🏝 южнобережный курорт",
  "Феодосия": "🏝 курорт, порт, оптика (завод «Море»)",
  "Джанкой": "🚆 ж/д узел, АПК",
  "Алушта": "🏝 приморский курорт",
  "Саки": "🏝 бальнеогрязевой курорт, химия (содовый)",
  "Бахчисарай": "🏭 стройматериалы, туризм (ханский дворец)",
  "Красноперекопск": "🧪 химия (сода, бром)",
  "Армянск": "🧪 химия (диоксид титана — «Крымский Титан»)",
  "Старый Крым": "🏝 туризм, АПК",
  "Щёлкино": "⚡ недостроенная Крымская АЭС, курорт",
  "Судак": "🍇 виноделие, курорт (крепость)",
  "Алупка": "🏝 курорт (Воронцовский дворец)",
  "Инкерман": "🍷 виноделие, добыча камня",
  "Балаклава": "⚓ бывшая база подлодок, курорт, виноделие"
}""")


def _norm(s):
    return s.strip().lower().replace("ё", "е")


# ──────────────────  ПОИСК ГОРОДОВ ПО СПЕЦИАЛИЗАЦИИ (ключевым словам)  ──────────────────
# Каждая запись в INDUSTRY начинается с эмодзи-категории (☢ атом, 🚀 ракеты/космос,
# 🌾 агро, 🔫 оборонка и т.д.) плюс текстовое описание. Обычный поиск подстроки не
# найдёт, например, «ядерщики» в тексте «ЗАТО, ядерный центр» — поэтому у частых
# бытовых запросов есть словарь синонимов на реальные слова/эмодзи-категории базы.
SPEC_SYNONYMS = {
    "впк": ["🔫", "🧨", "🚀", "оборон", "оружи", "боеприпас", "зато"],
    "оборонка": ["🔫", "🧨", "оборон", "оружи", "боеприпас", "зато"],
    "оборонный": ["🔫", "🧨", "оборон", "оружи", "боеприпас", "зато"],
    "оружие": ["🔫", "оружи"],
    "ракета": ["🚀", "ракет", "космодром", "космич", "грц"],
    "ракеты": ["🚀", "ракет", "космодром", "космич", "грц"],
    "ракетостроение": ["🚀", "ракет", "грц"],
    "космос": ["🚀", "космодром", "космич", "спутник"],
    "космонавтика": ["🚀", "космодром", "космич", "спутник"],
    "агро": ["апк", "🌾", "сельхоз", "агропром"],
    "агропром": ["апк", "🌾", "сельхоз"],
    "сельское хозяйство": ["апк", "🌾", "сельхоз"],
    "ядерщик": ["☢", "атом", "ядерн", "аэс", "уран", "радиохим"],
    "ядерщики": ["☢", "атом", "ядерн", "аэс", "уран", "радиохим"],
    "атом": ["☢", "атом", "ядерн", "аэс", "уран", "радиохим"],
    "атомщики": ["☢", "атом", "ядерн", "аэс", "уран", "радиохим"],
    "атомная промышленность": ["☢", "атом", "ядерн", "аэс", "уран", "радиохим"],
    "аэс": ["☢", "аэс"],
    "нефть": ["🛢", "нефт"],
    "нефтяники": ["🛢", "нефт"],
    "нефтедобыча": ["🛢", "нефт"],
    "газ": ["🔥", "газ"],
    "газовики": ["🔥", "газ"],
    "металлургия": ["⛓", "металлург"],
    "металлурги": ["⛓", "металлург"],
    "горнодобыча": ["⛏", "гок", "руда", "горнодобыч"],
    "шахтёры": ["⛏", "уголь", "угледобыч"],
    "уголь": ["уголь", "угледобыч"],
    "текстиль": ["🧵", "текстиль"],
    "лес": ["🌲", "лес", "деревообработ", "цбк", "лесопромышлен"],
    "лесники": ["🌲", "лес", "деревообработ"],
    "судостроение": ["🚢", "судостро", "судоремонт"],
    "судостроители": ["🚢", "судостро"],
    "авиация": ["✈", "авиа"],
    "авиастроение": ["✈", "авиа"],
    "автопром": ["🚗", "автопром", "автомобил"],
    "автомобилестроение": ["🚗", "автопром", "автомобил"],
    "туризм": ["🏝", "туризм", "курорт"],
    "курорт": ["🏝", "курорт"],
    "рыба": ["🐟", "рыб"],
    "рыбаки": ["🐟", "рыб"],
    "рыболовство": ["🐟", "рыб"],
    "оптика": ["оптик"],
    "часы": ["часов", "часы"],
    "стекло": ["🔮", "стекл", "хрусталь"],
    "наукоград": ["наукоград"],
    "физики": ["наукоград", "физик"],
    "it": ["it", "айти"],
    "порт": ["порт"],
    "химия": ["🧪", "хими"],
    "химики": ["🧪", "хими"],
}


# ──────────────────  ИЗВЛЕЧЕНИЕ ГОРОДА ИЗ АДРЕСА (для сопоставления ФИО+Города)  ──────────────────
_ADDR_REGION_WORDS = {"область", "обл", "край", "республика", "р-н", "район", "рн", "ао",
                       "автономный", "округ", "бл"}
_ADDR_STREET_MARKERS = (r"ул\.?|улица|пр-кт|проспект|просп\.?|пл\.?|площадь|мкр\.?|переулок|"
                         r"пер\.?|дом|д\.|кв\.?|квартал|кв-л|шоссе|бульвар|б-р")
# города, у которых первое слово стандартно продолжается вторым словом
_ADDR_MULTI_WORD_STARTS = {"нижний", "нижняя", "нижнее", "верхний", "верхняя", "великий",
                            "великие", "старый", "старая", "новый", "новая", "усть",
                            "красный", "южно", "северо", "санкт", "ростов"}


def _addr_clean_token(s):
    return s.strip(" .,\t\"'")


def _addr_capture_city(words):
    if not words:
        return ""
    city = _addr_clean_token(words[0])
    if city.lower().rstrip(".,") in _ADDR_MULTI_WORD_STARTS and len(words) > 1:
        nxt = _addr_clean_token(words[1])
        if nxt and nxt[0].isalpha():
            city = f"{city} {nxt}"
    return city


def _addr_drop_trailing_street_word(words):
    out = []
    for w in words:
        if re.fullmatch(_ADDR_STREET_MARKERS, w, flags=re.IGNORECASE):
            break
        out.append(w)
    return out


def extract_city_from_address(addr):
    """Достаёт название города из свободно записанного адреса — независимо от
    порядка «область/город/улица», регистра букв и наличия пробела после «г.».
    Проверено на реальных «грязных» адресах разных форматов."""
    if not addr:
        return ""
    addr = re.sub(r"\s+", " ", str(addr).strip())

    # 1) явный маркер "г."/"гор."/"город" перед названием (в любом регистре,
    #    с пробелом или без — "г.Мурманск", "Г. МУРМАНСК", "город Заринск")
    m = re.search(
        r"(?:^|[,.\s])(?:город\b|гор\.|гор\b|г\.|г\b)\s*([А-ЯЁ][а-яё\-]*(?:\s[А-ЯЁ][а-яё\-]*)?)",
        addr, flags=re.IGNORECASE)
    if m:
        words = [w for w in m.group(1).split(" ") if w]
        words = _addr_drop_trailing_street_word(words)
        if words:
            return _addr_capture_city(words)

    # 2) название города ПЕРЕД словом "г." (напр. «Красноярск г., Красноярский край.»)
    m = re.match(r"^\s*([А-ЯЁ][а-яё\-]+)\s+г\.?[,.]", addr, flags=re.IGNORECASE)
    if m:
        return _addr_clean_token(m.group(1))

    # 3) маркера нет вообще — берём начало строки до запятой/маркера улицы,
    #    отбрасывая слова региона (область/край/республика и т.п.)
    first_part = re.split(r",", addr)[0].strip(". ").strip()
    first_part = re.split(rf"\s+(?:{_ADDR_STREET_MARKERS})\b", first_part, maxsplit=1,
                           flags=re.IGNORECASE)[0]
    words = [w for w in first_part.split(" ") if w]
    while words and words[0].strip(".,").lower() in _ADDR_REGION_WORDS:
        words.pop(0)
    return _addr_capture_city(words)


def normalize_city_case(city):
    if not city:
        return city
    def cap(w):
        return w[:1].upper() + w[1:].lower() if w else w
    return " ".join("-".join(cap(p) for p in w.split("-")) for w in city.split(" "))


def search_by_specialization(query, cities):
    """Ищет города по слову/фразе специализации — независимо от какого-либо
    выбранного центра/радиуса. Сначала смотрит словарь синонимов (частые
    бытовые формулировки вроде «ВПК», «ядерщики», «ракеты», «агро»), затем,
    если синонимов нет, просто ищет запрос как подстроку в тексте специализации.
    Возвращает список городов (без km/code — это заполняет вызывающий код)."""
    q = _norm(query)
    if not q:
        return []
    needles = SPEC_SYNONYMS.get(q)
    if needles is None:
        needles = [q]
    result = []
    for c in cities:
        if c.get("ext"):
            continue   # у сёл/посёлков из расширенной базы специализаций нет
        spec = INDUSTRY.get(c["n"], "")
        if not spec:
            continue
        spec_l = spec.lower()
        if any(n in spec_l for n in needles):
            result.append(c)
    result.sort(key=lambda c: -(c["pop"] or 0))
    return result


# ───────────────  РАБОЧИЕ ПОСЁЛКИ (ПГТ) — добавка к основной базе  ───────────────
# Основная база (_CITY_B64) — это города. Ниже отдельно добавлены рабочие
# посёлки/ПГТ — населённые пункты без статуса города. Население и координаты —
# по открытым данным (перепись 2021 и более поздние оценки, округлённо),
# специализация — по известным градообразующим предприятиям.
# ВАЖНО: телефонных кодов для этих посёлков в базе нет (в отличие от ph_codes/
# ph_ranges для городов) — сознательно, чтобы не выдумывать телеком-данные,
# которые программа не может проверить. Как и у части городов без кода, в
# столбце «код» у них будет прочерк — это ожидаемое поведение, не ошибка.
EXTRA_SETTLEMENTS = [
    {"n": "Яблоновский", "lat": 45.0167, "lon": 38.9667, "pop": 55000, "sub": "Адыгея"},
    {"n": "Нахабино", "lat": 55.8386, "lon": 37.1067, "pop": 51000, "sub": "Московская область"},
    {"n": "Кольцово", "lat": 54.9333, "lon": 83.1833, "pop": 15000, "sub": "Новосибирская область"},
    {"n": "Дрожжино", "lat": 55.5167, "lon": 37.4667, "pop": 43000, "sub": "Московская область"},
    {"n": "Свердловский", "lat": 55.8167, "lon": 38.1833, "pop": 21000, "sub": "Московская область"},
    {"n": "Черёмушки", "lat": 52.7833, "lon": 91.3667, "pop": 9000, "sub": "Хакасия"},
    {"n": "Сириус", "lat": 43.4056, "lon": 39.9564, "pop": 13000, "sub": "Краснодарский край"},
    {"n": "Горячеводский", "lat": 44.0167, "lon": 43.1, "pop": 37000, "sub": "Ставропольский край"},
    {"n": "Тикси", "lat": 71.6333, "lon": 128.8667, "pop": 4300, "sub": "Якутия"},
    {"n": "Восход", "lat": 55.19, "lon": 36.62, "pop": 8000, "sub": "Московская область"},
    {"n": "Молодёжный", "lat": 52.30, "lon": 104.20, "pop": 4500, "sub": "Иркутская область"},
    {"n": "Солнечный", "lat": 50.72, "lon": 136.63, "pop": 8000, "sub": "Хабаровский край"},
    {"n": "Пойковский", "lat": 61.05, "lon": 74.45, "pop": 30000, "sub": "ХМАО"},
    {"n": "Излучинск", "lat": 60.83, "lon": 76.62, "pop": 14000, "sub": "ХМАО"},
    {"n": "Фёдоровский", "lat": 61.27, "lon": 73.63, "pop": 21000, "sub": "ХМАО"},
    {"n": "Барсово", "lat": 61.20, "lon": 73.15, "pop": 12000, "sub": "ХМАО"},
    {"n": "Белый Яр", "lat": 61.63, "lon": 73.13, "pop": 20000, "sub": "ХМАО"},
    {"n": "Уренгой", "lat": 65.97, "lon": 78.37, "pop": 10000, "sub": "ЯНАО"},
    {"n": "Тазовский", "lat": 67.47, "lon": 78.70, "pop": 7000, "sub": "ЯНАО"},
    {"n": "Пангоды", "lat": 65.83, "lon": 74.48, "pop": 10000, "sub": "ЯНАО"},
    {"n": "Никель", "lat": 69.4, "lon": 30.22, "pop": 11000, "sub": "Мурманская область"},
]

EXTRA_INDUSTRY = {
    "Яблоновский": "🏙 спальный пригород Краснодара, логистика, торговля",
    "Нахабино": "🚀 полигон ГИРД-09 (первая советская ракета), моторвагонное депо",
    "Кольцово": "🔬 наукоград, вирусология (ГНЦ «Вектор»)",
    "Дрожжино": "🏙 спальный пригород Москвы",
    "Свердловский": "🧵 текстиль (шёлк), пригород Щёлково",
    "Черёмушки": "⚡ Саяно-Шушенская ГЭС",
    "Сириус": "🏝 Олимпийский парк, туризм, IT (федеральная территория)",
    "Горячеводский": "🏝 курорт (Пятигорск), туризм",
    "Тикси": "⚓ морской порт, погранзастава, полярная станция",
    "Восход": "🚀 ЗАТО, управление космодромом Байконур",
    "Молодёжный": "🔫 ЗАТО, авиабаза",
    "Солнечный": "⛏ горнодобыча, оловянный рудник",
    "Пойковский": "🛢 нефтедобыча (Приобьнефтегаз)",
    "Излучинск": "🛢 нефтедобыча (Нижневартовскнефтегаз)",
    "Фёдоровский": "🛢 нефтедобыча (Сургутнефтегаз)",
    "Барсово": "🛢 нефтедобыча",
    "Белый Яр": "🌲 лес, нефтедобыча",
    "Уренгой": "🔥 газодобыча",
    "Тазовский": "🔥 газ, 🐟 рыба",
    "Пангоды": "🔥 газ (Медвежье месторождение)",
    "Никель": "⛏ никелевый комбинат «Печенганикель»",
}
INDUSTRY.update(EXTRA_INDUSTRY)


# Регионы, полностью исключаемые из базы городов по просьбе пользователя.
EXCLUDED_REGIONS = {"Дагестан"}


# ───────────────  РАСШИРЕННАЯ БАЗА: сёла, посёлки, ПГТ  ───────────────
# Лежит отдельным файлом settlements_ext.bin рядом с программой и
# подгружается ПОВЕРХ основной базы (_CITY_B64/_PHONES_B64 не меняются).
# Внутри (gzip + JSON):
#   "s": [[название, lat, lon, население, регион, тип, код, [группы диапазонов]]]
#   "g": группы диапазонов [[from, span, cap, op, kind], ...]
#   "o": операторы для этих диапазонов
# Источники: реестр нумерации Минцифры (через зеркало на GitHub) и набор
# «Населённые пункты России» (ЦПУР). В реестре диапазоны выделены только
# до уровня города / городского округа / муниципального района — поэтому у
# сёл и посёлков диапазоны — это номера их района (kind=1, «район»), а свои
# (kind=0, «город») — только там, где посёлок указан в реестре отдельно.
SETTLEMENT_TYPES = {
    "г": "город", "пгт": "ПГТ", "рп": "рабочий посёлок", "дп": "дачный посёлок",
    "кп": "курортный посёлок", "гп": "городской посёлок", "с": "село", "п": "посёлок",
    "д": "деревня", "х": "хутор", "ст-ца": "станица", "аул": "аул", "сл": "слобода",
    "ст": "станция", "нп": "населённый пункт", "п/ст": "посёлок при станции",
    "ж/д_ст": "ж/д станция", "у": "улус", "с/п": "сельское поселение",
    "мкр": "микрорайон", "аал": "аал", "м": "местечко",
}
EXT_KEYS = set()      # ключи «название|регион» из расширенной базы
# Сёла внутри городских округов крупных городов: их «район» — это номера
# самого города, поэтому при определении города по номеру их не используем.
EXT_IN_CITY = set()
_EXT_CACHE = None


def _load_ext_db():
    global _EXT_CACHE
    if _EXT_CACHE is None:
        try:
            with open(os.path.join(_app_dir(), "settlements_ext.bin"), "rb") as f:
                _EXT_CACHE = json.loads(gzip.decompress(f.read()).decode("utf-8"))
        except Exception:
            _EXT_CACHE = {"s": [], "g": [], "o": []}
    return _EXT_CACHE


def load_city_db():
    raw = gzip.decompress(base64.b64decode(_CITY_B64)).decode("utf-8")
    arr = json.loads(raw)
    cities = [{"n": a[0], "lat": a[1], "lon": a[2], "pop": a[3], "sub": a[4]} for a in arr]
    cities.extend(EXTRA_SETTLEMENTS)  # рабочие посёлки — см. пояснение выше
    for a in _load_ext_db().get("s", []):
        cities.append({"n": a[0], "lat": a[1], "lon": a[2], "pop": a[3], "sub": a[4],
                       "t": a[5], "ext": True})
    cities = [c for c in cities if c["sub"] not in EXCLUDED_REGIONS]
    index = {}
    for c in cities:
        index.setdefault(_norm(c["n"]), []).append(c)
    return cities, index

def load_phone_db():
    """-> (ops_list, codes{key}, ranges{key:[[from,span,cap,opidx,kind]]})"""
    raw = gzip.decompress(base64.b64decode(_PHONES_B64)).decode("utf-8")
    d = json.loads(raw)
    ops, codes, ranges = list(d["o"]), dict(d["c"]), dict(d["r"])
    ext = _load_ext_db()
    off = len(ops)
    ops.extend(ext.get("o", []))
    groups = [[[r[0], r[1], r[2], r[3] + off, r[4]] for r in g] for g in ext.get("g", [])]
    for a in ext.get("s", []):
        key = f"{a[0]}|{a[4]}"
        if key in codes or key in ranges:
            continue          # основная база всегда в приоритете
        EXT_KEYS.add(key)
        if len(a) > 8 and a[8]:
            EXT_IN_CITY.add(key)
        codes[key] = a[6]
        rr = []
        for gi in a[7]:
            rr.extend(groups[gi])
        ranges[key] = rr
    return ops, codes, ranges

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    p = math.radians
    dlat = p(lat2 - lat1); dlon = p(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p(lat1)) * math.cos(p(lat2)) * math.sin(dlon / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))

def fmt_num(n):
    s = str(n).zfill(10)
    return f"+7-{s[:3]}-{s[3:6]}-{s[6:]}"


# ─────────────────────────────  ЦВЕТОВЫЕ ХЕЛПЕРЫ  ─────────────────────────────
def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def _rgb_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(c)))) for c in rgb)

def _lerp(c1, c2, t):
    a, b = _hex_to_rgb(c1), _hex_to_rgb(c2)
    return _rgb_to_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


# ─────────────────────────────  ЛОГИКА ОЧИСТКИ  ─────────────────────────────
def get_city_from_called(called_dir):
    try:
        for f in os.listdir(called_dir):
            if f.endswith(".csv") or f.endswith(".xlsx"):
                name = os.path.splitext(f)[0]
                name = re.sub(r'^[\d_.\s]+', '', name).strip()
                city = re.split(r'[\s_]+', name)[0]
                if city:
                    return city
    except:
        pass
    return "result"

def load_phones_from_dir(called_dir):
    phones = set()
    try:
        for f in os.listdir(called_dir):
            path = os.path.join(called_dir, f)
            try:
                if f.endswith(".csv"):
                    df = pd.read_csv(path)
                    col = "phone_number" if "phone_number" in df.columns else df.columns[0]
                    for num in df[col].dropna():
                        n = str(num).lstrip("+").strip()
                        if n.isdigit() and len(n) in (10, 11):
                            phones.add(n)
                elif f.endswith(".xlsx"):
                    df = pd.read_excel(path, header=None)
                    for num in df.iloc[:, 0].dropna():
                        n = str(num).lstrip("+").strip()
                        if n.isdigit() and len(n) in (10, 11):
                            phones.add(n)
            except:
                pass
    except:
        pass
    return phones

def load_uncalled_numbers(uncalled_dir):
    nums = []
    try:
        for f in sorted(os.listdir(uncalled_dir)):
            if not f.endswith(".xlsx"):
                continue
            path = os.path.join(uncalled_dir, f)
            try:
                df = pd.read_excel(path, header=None)
                df.columns = ["Номер"]
                df["Номер"] = df["Номер"].astype(str).str.lstrip("+").str.strip()
                nums.extend(df["Номер"].tolist())
            except:
                pass
    except:
        pass
    return nums


# ───────────────────  ПОДДЕРЖКА ВЫБОРА НЕСКОЛЬКИХ CSV-ФАЙЛОВ  ───────────────────
# «Прозвоненные» можно указать не только папкой, но и списком конкретных CSV
# файлов (через FolderRow.allow_files). В этом случае FolderRow.get() возвращает
# строку вида "путь1.csv; путь2.csv; ...". Функции ниже понимают оба варианта.

def _split_file_sources(source):
    """Если source — это один/несколько файлов (не папка), вернуть список путей.
    Если source — папка (или пусто/не существует), вернуть None."""
    if not source:
        return None
    if os.path.isdir(source):
        return None
    parts = [p.strip() for p in source.split(";") if p.strip()]
    if parts and all(os.path.isfile(p) for p in parts):
        return parts
    return None


def source_is_valid(source):
    """Папка существует, либо все указанные CSV/XLSX файлы существуют."""
    if not source:
        return False
    files = _split_file_sources(source)
    if files is not None:
        return True
    return os.path.isdir(source)


def load_phones_by_status_from_source(source):
    """Как load_phones_from_source, но дополнительно группирует номера по
    статусу обзвона («Недозвон», «Перезвон», «Дозвон» и т.д.), если такой
    столбец есть в файле — тот же формат выгрузки («phone_number» /
    «status»), что использует вкладка «Статусы». Нужна вкладке «Очистка»,
    чтобы перед удалением дублей спросить, какие статусы считать дублями
    (удалить), а какие — оставить в результате, даже если номер совпал
    с «прозвоненными».

    Возвращает (all_phones, phones_by_status):
      all_phones       — set всех найденных номеров (как раньше возвращала
                          load_phones_from_source, для обратной совместимости)
      phones_by_status — {статус: set(номеров)}; номера из файлов без
                          столбца статуса (старый формат — просто список
                          номеров) попадают под ключ "" — они, как и раньше,
                          считаются дублями безусловно.
    """
    files = _split_file_sources(source)
    if files is None:
        if source and os.path.isdir(source):
            try:
                files = [os.path.join(source, f) for f in sorted(os.listdir(source))
                         if f.lower().endswith((".csv", ".xlsx"))]
            except Exception:
                files = []
        else:
            files = []

    all_phones = set()
    by_status = {}

    def _add(raw_num, status):
        n = str(raw_num).lstrip("+").strip()
        if n.isdigit() and len(n) in (10, 11):
            all_phones.add(n)
            by_status.setdefault(status, set()).add(n)

    for path in files:
        low = path.lower()
        try:
            if low.endswith(".csv"):
                raw, last_err = None, None
                for enc in ("utf-8-sig", "utf-8", "cp1251"):
                    try:
                        raw = pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
                        break
                    except Exception as e:
                        last_err = e
                if raw is None:
                    continue
                cols_lower = [str(c).strip().lower() for c in raw.columns]
                if "phone_number" in cols_lower:
                    raw.columns = cols_lower
                    phone_col = "phone_number"
                    status_col = "status" if "status" in cols_lower else None
                else:
                    phone_col, status_col = raw.columns[0], None
            elif low.endswith(".xlsx"):
                raw = pd.read_excel(path, header=None, dtype=str).fillna("")
                first_row = ([str(x).strip().lower() for x in raw.iloc[0].tolist()]
                             if len(raw) else [])
                if "phone_number" in first_row:
                    raw.columns = first_row
                    raw = raw.iloc[1:].reset_index(drop=True)
                    phone_col = "phone_number"
                    status_col = "status" if "status" in first_row else None
                else:
                    phone_col, status_col = 0, None
            else:
                continue

            for i in range(len(raw)):
                num = raw[phone_col].iloc[i]
                status = str(raw[status_col].iloc[i]).strip() if status_col is not None else ""
                _add(num, status)
        except Exception:
            pass

    return all_phones, by_status


def get_city_from_source(source):
    """Как get_city_from_called, но также понимает список CSV-файлов."""
    files = _split_file_sources(source)
    if files is None:
        return get_city_from_called(source)
    for path in files:
        name = os.path.splitext(os.path.basename(path))[0]
        name = re.sub(r'^[\d_.\s]+', '', name).strip()
        city = re.split(r'[\s_]+', name)[0]
        if city:
            return city
    return "result"


def load_phones_from_source(source):
    """Как load_phones_from_dir, но также понимает список конкретных CSV/XLSX файлов
    (несколько файлов выбранных вручную вместо папки)."""
    files = _split_file_sources(source)
    if files is None:
        return load_phones_from_dir(source)
    phones = set()
    for path in files:
        try:
            if path.endswith(".csv"):
                df = pd.read_csv(path)
                col = "phone_number" if "phone_number" in df.columns else df.columns[0]
                for num in df[col].dropna():
                    n = str(num).lstrip("+").strip()
                    if n.isdigit() and len(n) in (10, 11):
                        phones.add(n)
            elif path.endswith(".xlsx"):
                df = pd.read_excel(path, header=None)
                for num in df.iloc[:, 0].dropna():
                    n = str(num).lstrip("+").strip()
                    if n.isdigit() and len(n) in (10, 11):
                        phones.add(n)
        except Exception:
            pass
    return phones


# ═══════════════════════════════════════════════════════════════════════════
#  ОРГАНИЗАЦИИ ЧЕРЕЗ parser-2gis (бесплатный сборщик поверх сайта 2ГИС, без
#  платного API-ключа) — детские сады, заводы, школы, столовые, казначейство,
#  полиция, администрация, офисы. Используется во вкладке «Очистка», чтобы
#  убирать такие номера вместе с обычными дублями.
#  Требует: pip install parser-2gis  и установленный Google Chrome.
# ═══════════════════════════════════════════════════════════════════════════

ORG_SEARCH_CATEGORIES = [
    "детский сад", "завод", "школа", "столовая",
    "казначейство", "полиция", "администрация", "офис",
]

ORGS_CACHE_FILE = os.path.join(_app_dir(), "orgs_phones_cache.json")

# Крупные города 2ГИС называет по-своему в адресе страницы (не транслитом),
# поэтому для них — явные исключения. Для всех остальных городов ниже
# используется обычная транслитерация.
_CITY_SLUG_OVERRIDES = {
    "москва": "moscow", "санкт-петербург": "spb", "питер": "spb",
    "новосибирск": "novosibirsk", "екатеринбург": "ekaterinburg",
    "казань": "kazan", "нижний новгород": "nnovgorod",
    "челябинск": "chelyabinsk", "самара": "samara", "омск": "omsk",
    "ростов-на-дону": "rostov", "уфа": "ufa", "красноярск": "krasnoyarsk",
    "воронеж": "voronezh", "пермь": "perm", "волгоград": "volgograd",
    "краснодар": "krasnodar", "тюмень": "tyumen", "мурманск": "murmansk",
}

_TRANSLIT_MAP = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '',
    'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya', ' ': '-', '-': '-',
}


_city_db_index_cache = None


def _get_city_db_index():
    """load_city_db() один раз кэшируем — база большая (gzip+base64), не нужно
    распаковывать её заново на каждый вызов city_to_2gis_slug."""
    global _city_db_index_cache
    if _city_db_index_cache is None:
        _, index = load_city_db()
        _city_db_index_cache = index
    return _city_db_index_cache


def _canonical_city_name(city):
    """Ищет город во вшитой базе (та же, что используется во всей остальной
    программе для городов/специализаций) и возвращает его официальное
    написание из базы — так лечатся опечатки/неверный регистр в названии
    города, из-за которых транслитерация могла давать неправильный адрес
    2ГИС. Если города в базе нет — вернёт None, и в ход пойдёт то, что
    передали изначально."""
    try:
        index = _get_city_db_index()
    except Exception:
        return None
    matches = index.get(_norm(city))
    if not matches:
        return None
    # если несколько городов с одинаковым названием в разных регионах —
    # берём самый крупный по населению (наиболее вероятный вариант)
    best = max(matches, key=lambda c: (c.get("pop") or 0))
    return best.get("n")


def city_to_2gis_slug(city):
    """Пытается угадать сегмент города в адресе 2ГИС (например, «Мурманск» →
    «murmansk»). Для крупных городов со специальными названиями — явный
    список исключений. Если города нет в этом списке, сверяем название с уже
    вшитой базой городов (та же, что используется остальной программой) —
    это поправит опечатки/регистр перед транслитерацией. Для всех прочих —
    простая транслитерация. Стопроцентной гарантии совпадения всё равно нет:
    если 2ГИС не найдёт такой город, парсинг вернёт 0 организаций, и это
    будет видно в логе программы."""
    key = city.strip().lower()
    if key in _CITY_SLUG_OVERRIDES:
        return _CITY_SLUG_OVERRIDES[key]

    canonical = _canonical_city_name(city)
    if canonical:
        canon_key = canonical.strip().lower()
        if canon_key in _CITY_SLUG_OVERRIDES:
            return _CITY_SLUG_OVERRIDES[canon_key]
        key = canon_key

    return "".join(_TRANSLIT_MAP.get(ch, ch) for ch in key).strip("-")


def _find_parser2gis_cmd():
    """Ищет parser-2gis не только в PATH текущего процесса (shutil.which),
    но и в типичных местах, куда pip кладёт консольные скрипты — это чинит
    ситуацию «в cmd работает, а из запущенной программы — не найден»:
    у собранного .exe / у GUI-приложения, запущенного двойным кликом, PATH
    может отличаться от того, что видно в открытом cmd (Windows не всегда
    обновляет PATH уже запущенных процессов/Проводника после установки
    пакета), либо pip ставил пакет в другой Python, чем тот, что первым
    найдётся в PATH. Возвращает список для subprocess (например
    ["parser-2gis"] или [sys.executable, "-m", "parser_2gis"]), либо
    None, если нигде не нашли."""
    exe = shutil.which("parser-2gis") or shutil.which("parser-2gis.exe")
    if exe:
        return [exe]

    py_dir = os.path.dirname(sys.executable)
    candidates = [
        os.path.join(py_dir, "Scripts", "parser-2gis.exe"),   # Windows venv/обычный python.exe
        os.path.join(py_dir, "parser-2gis.exe"),
        os.path.join(py_dir, "Scripts", "parser-2gis"),
        os.path.join(py_dir, "bin", "parser-2gis"),           # Linux/Mac
    ]
    # pip install --user — свой Scripts/bin, который может отсутствовать в PATH
    # вообще где бы то ни было.
    try:
        import site
        user_base = site.getuserbase()
        candidates += [
            os.path.join(user_base, "Scripts", "parser-2gis.exe"),
            os.path.join(user_base, "bin", "parser-2gis"),
        ]
    except Exception:
        pass

    for path in candidates:
        if os.path.isfile(path):
            return [path]

    # Запасной вариант: если пакет виден тому же интерпретатору, из-под
    # которого работает сама программа, запускаем его как модуль — сработает,
    # даже если готового exe-скрипта нигде не нашлось.
    try:
        import importlib.util
        if importlib.util.find_spec("parser_2gis") is not None:
            return [sys.executable, "-m", "parser_2gis"]
    except Exception:
        pass

    return None


def parser2gis_available():
    return _find_parser2gis_cmd() is not None


def build_2gis_search_urls(city_slug, categories=None):
    categories = categories or ORG_SEARCH_CATEGORIES
    return [f"https://2gis.ru/{city_slug}/search/{_url_quote(cat)}" for cat in categories]


def run_parser_2gis(urls, timeout=1800, on_line=None):
    """Запускает parser-2gis (реальный Chrome под капотом) и возвращает путь
    к временному xlsx-файлу с результатом. on_line(str) — необязательный
    колбэк, получает строки вывода парсера по ходу работы (для отображения
    прогресса в логе).

    --chrome.headless yes — прячет окно браузера (по умолчанию у parser-2gis
    оно ВИДИМОЕ — тот самый «гугл», который у тебя мелькал). С headless
    парсинг идёт полностью «в консоли», без всплывающего окна Chrome, и
    заодно чуть быстрее (не тратится время на отрисовку интерфейса).

    PYTHONIOENCODING=utf-8 в окружении дочернего процесса — чинит «крокозябры»
    в логе: сам parser-2gis пишет свои логи в кодировке консоли Windows
    (обычно cp1251), а мы читаем вывод как UTF-8 — без принудительного
    UTF-8 на стороне дочернего процесса кириллица превращается в мусор.

    ВНИМАНИЕ: все urls идут одним запуском — если хотя бы одна категория
    словит TimeoutError внутри самого parser-2gis (известный неисправленный
    баг библиотеки, см. github.com/interlark/parser-2gis issue #74 — 2ГИС
    иногда не укладывается в её внутренний таймаут ожидания ответа), это
    может обнулить результат по остальным категориям тоже. Для города
    предпочтительнее run_parser_2gis_by_category ниже — она гоняет каждую
    категорию отдельным (и вдобавок параллельным) запуском, так один
    зависший запрос не топит другие и не тормозит их по очереди."""
    base_cmd = _find_parser2gis_cmd()
    if base_cmd is None:
        raise RuntimeError(
            "Не найден parser-2gis. Установите: pip install parser-2gis "
            "(и Google Chrome, если он ещё не установлен)")
    fd, out_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    cmd = [*base_cmd, "-i", *urls, "-o", out_path, "-f", "xlsx", "--chrome.headless", "yes"]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace", env=env)
    try:
        for line in proc.stdout:
            if on_line:
                on_line(line.rstrip())
        proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        raise RuntimeError("parser-2gis не уложился в отведённое время ожидания")
    if proc.returncode != 0:
        raise RuntimeError(f"parser-2gis завершился с ошибкой (код {proc.returncode})")
    return out_path


def run_parser_2gis_by_category(urls, timeout_per_category=300, retries=1,
                                 on_line=None, max_workers=1):
    """Как run_parser_2gis, но гоняет parser-2gis отдельным запуском на КАЖДУЮ
    категорию (каждый url) вместо одного общего запуска на все сразу.

    ВАЖНО про max_workers: по умолчанию 1 (строго по очереди), и это
    осознанно. Параллельный запуск нескольких категорий одновременно ловит
    другой подтверждённый баг: сервер 2ГИС начинает отвечать ошибкой
    «Bound is incorrect. Set 'viewpoint1' ... 'viewpoint2' ...» — это описано
    в её баг-трекере (issue #66 у interlark/parser-2gis) как раз про
    ситуацию «если парсить сразу несколько (одновременно)». Похоже, сам
    2ГИС (или общая сессия/куки между параллельными Chrome-процессами)
    путается в границах карты при нескольких одновременных запросах.
    Поэтому категории идут последовательно — надёжность важнее скорости,
    когда параллельность буквально ломает ответы сервера.

    Раздельные (пусть и последовательные) запуски всё равно лечат исходную
    беду: у parser-2gis есть неисправленный автором баг (TimeoutError:
    MainParser._wait_requests_finished — сайт 2ГИС иногда не успевает
    ответить в её внутренний таймаут). Если гнать все категории одним
    запуском, зависание ОДНОЙ категории (например «завод») обнуляет весь
    результат по городу, хотя остальные 7 категорий могли уже собраться.
    Раздельные запуски + одна повторная попытка на зависшую категорию
    решают это — теряется максимум одна категория, а не весь город.

    Возвращает (phones, failed): phones — объединённый набор телефонов по
    всем успешно обработанным категориям; failed — список (url, ошибка) по
    категориям, которые не удалось собрать даже после повтора."""
    import concurrent.futures

    def _category_label(url):
        # достаём последний сегмент URL (категория) для читаемых логов
        from urllib.parse import unquote as _unq
        tail = url.rstrip("/").rsplit("/", 1)[-1]
        return _unq(tail)

    def _run_one(url):
        label = _category_label(url)
        last_err = None
        for attempt in range(retries + 1):
            out_path = None
            try:
                def _prefixed(line, _label=label):
                    if on_line and line.strip():
                        on_line(f"[{_label}] {line}")
                out_path = run_parser_2gis([url], timeout=timeout_per_category,
                                            on_line=_prefixed)
                phones = extract_phones_from_parser_result(out_path)
                return url, phones, None
            except Exception as e:
                last_err = e
                if on_line:
                    tries_left = retries - attempt
                    suffix = (f", ещё {tries_left} попытк(а/и)" if tries_left > 0
                              else ", пропускаю категорию")
                    on_line(f"[{label}] ⚠ Не удалось: {e}{suffix}")
            finally:
                if out_path:
                    try:
                        os.remove(out_path)
                    except Exception:
                        pass
        return url, None, last_err

    all_phones = set()
    failed = []
    if max_workers <= 1:
        results = (_run_one(url) for url in urls)  # строго по очереди
    else:
        pool = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        results = pool.map(_run_one, urls)
    for url, phones, err in results:
        if err is None:
            all_phones |= phones
        else:
            failed.append((url, err))
    if max_workers > 1:
        pool.shutdown(wait=True)
    return all_phones, failed


def extract_phones_from_parser_result(xlsx_path):
    """Достаёт все телефоны (столбцы «Телефон 1», «Телефон 2», ...) из
    результата parser-2gis и приводит их к 10-значному формату без кода
    страны — тому же, что используется во всей остальной программе."""
    df = pd.read_excel(xlsx_path, dtype=str)
    phone_cols = [c for c in df.columns if str(c).startswith("Телефон")]
    phones = set()
    for col in phone_cols:
        for raw in df[col].dropna():
            digits = re.sub(r"\D", "", str(raw))
            if len(digits) == 11 and digits[0] in "78":
                digits = digits[1:]
            if len(digits) == 10:
                phones.add(digits)
    return phones


def load_orgs_cache():
    try:
        with open(ORGS_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_orgs_cache_for_city(city, phones):
    cache = load_orgs_cache()
    cache[city] = {"phones": sorted(phones), "updated": datetime.date.today().isoformat()}
    try:
        with open(ORGS_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass


def get_cached_orgs_phones(city):
    """None — города ещё никогда не парсили, set() (возможно пустой) — уже
    парсили раньше, используем сохранённый результат без повторного похода
    в браузер."""
    entry = load_orgs_cache().get(city)
    if entry is None:
        return None
    return set(entry.get("phones", []))


