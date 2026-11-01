# -*- coding: utf-8 -*-
"""Обновление программы с GitHub по кнопке «Обновить».

Как это устроено
----------------
В репозитории на GitHub лежат файлы программы и version.json:

    {
      "version": "3.2",
      "files": ["cleaner_qt.pyw", "cleaner_core.py", "updater.py"],
      "changes": ["что нового 1", "что нового 2"]
    }

Кнопка «Обновить» скачивает version.json, сравнивает версию с текущей
(cleaner_core.APP_VERSION) и, если на GitHub новее, скачивает файлы из
списка "files" во временную папку, проверяет, что .py-файлы не битые,
кладёт старые файлы в app/_backup/<старая версия>/ и только потом
подменяет их новыми. Если что-то пошло не так на любом шаге — текущие
файлы остаются нетронутыми.

Пользовательские файлы (theme.json, trigger_words.txt, кэши) обновлением
не трогаются, если их специально не указать в "files".

Настройки — в update_config.json рядом с программой:
    {"repo": "владелец/репозиторий", "branch": "main", "token": ""}
token нужен только для приватного репозитория.
"""

import json
import os
import py_compile
import shutil
import ssl
import sys
import tempfile
import urllib.error
import urllib.request

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "update_config.json")
BACKUP_DIR = os.path.join(APP_DIR, "_backup")

DEFAULT_CONFIG = {"repo": "", "branch": "main", "token": ""}

# Файлы, которые нельзя перезаписывать никогда — настройки и данные пользователя.
PROTECTED = {"update_config.json", "theme.json", "orgs_phones_cache.json"}


class UpdateError(Exception):
    pass


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            cfg.update({k: v for k, v in json.load(f).items() if v is not None})
    except FileNotFoundError:
        save_config(cfg)
    except Exception:
        pass
    cfg["repo"] = str(cfg.get("repo", "")).strip().strip("/")
    cfg["branch"] = str(cfg.get("branch") or "main").strip()
    cfg["token"] = str(cfg.get("token") or "").strip()
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _ver_tuple(v):
    out = []
    for part in str(v).strip().split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        out.append(int(digits) if digits else 0)
    while out and out[-1] == 0:
        out.pop()
    return tuple(out)


def is_newer(remote, local):
    return _ver_tuple(remote) > _ver_tuple(local)


def _ssl_context():
    ctx = ssl.create_default_context()
    try:
        import certifi  # если вдруг установлен — надёжнее системного хранилища
        ctx.load_verify_locations(certifi.where())
    except Exception:
        pass
    return ctx


def _fetch(cfg, path, timeout=30):
    """Скачивает один файл из репозитория (через GitHub API — без задержки
    кэша, которая бывает у raw.githubusercontent.com)."""
    if not cfg["repo"]:
        raise UpdateError(
            "Не указан репозиторий для обновлений.\n"
            f"Впиши его в файл:\n{CONFIG_PATH}")
    url = (f"https://api.github.com/repos/{cfg['repo']}/contents/"
           f"{urllib.request.quote(path)}?ref={urllib.request.quote(cfg['branch'])}")
    req = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github.raw",
        "User-Agent": "Base-updater",
        "Cache-Control": "no-cache",
    })
    if cfg["token"]:
        req.add_header("Authorization", f"Bearer {cfg['token']}")
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as r:
            return r.read()
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise UpdateError(
                f"На GitHub не найден файл «{path}».\n"
                f"Проверь репозиторий «{cfg['repo']}» и ветку «{cfg['branch']}»"
                + ("" if cfg["token"] else " (если репозиторий приватный — нужен token)") + ".")
        if e.code in (401, 403):
            raise UpdateError(
                "GitHub отказал в доступе (неверный token или превышен лимит запросов).\n"
                "Попробуй ещё раз через несколько минут.")
        raise UpdateError(f"Ошибка GitHub: {e.code} {e.reason}")
    except urllib.error.URLError as e:
        raise UpdateError(f"Нет связи с GitHub: {e.reason}")
    except Exception as e:
        raise UpdateError(f"Не удалось скачать «{path}»: {e}")


def check(current_version):
    """Возвращает (есть_обновление, информация_из_version.json)."""
    cfg = load_config()
    raw = _fetch(cfg, "version.json", timeout=15)
    try:
        info = json.loads(raw.decode("utf-8-sig"))
    except Exception:
        raise UpdateError("Файл version.json на GitHub повреждён.")
    if "version" not in info:
        raise UpdateError("В version.json нет поля \"version\".")
    info.setdefault("files", [])
    info.setdefault("changes", [])
    return is_newer(info["version"], current_version), info


def _safe_rel(path):
    rel = os.path.normpath(str(path).replace("\\", "/")).replace("\\", "/")
    if rel.startswith("../") or rel == ".." or os.path.isabs(rel) or ":" in rel:
        raise UpdateError(f"Недопустимый путь в version.json: {path}")
    return rel


def install(info, current_version, progress=None):
    """Скачивает и устанавливает файлы из info["files"].
    progress(текст) — необязательный колбэк для статуса."""
    cfg = load_config()
    files = [_safe_rel(p) for p in info.get("files", [])]
    files = [p for p in files if os.path.basename(p) not in PROTECTED]
    if not files:
        raise UpdateError("В version.json пустой список файлов.")

    tmp = tempfile.mkdtemp(prefix="base_update_")
    try:
        # 1) скачиваем всё во временную папку
        for i, rel in enumerate(files, 1):
            if progress:
                progress(f"Скачиваю {i}/{len(files)}: {rel}")
            data = _fetch(cfg, rel)
            dst = os.path.join(tmp, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            with open(dst, "wb") as f:
                f.write(data)

        # 2) проверяем, что код целый (не обрезан, без синтаксических ошибок)
        if progress:
            progress("Проверяю файлы…")
        for rel in files:
            if rel.endswith((".py", ".pyw")):
                try:
                    py_compile.compile(os.path.join(tmp, rel), doraise=True,
                                       cfile=os.path.join(tmp, "_check.pyc"))
                except py_compile.PyCompileError as e:
                    raise UpdateError(f"Файл {rel} с ошибкой, обновление отменено:\n{e.msg}")

        # 3) резервная копия текущих файлов
        if progress:
            progress("Сохраняю резервную копию…")
        bdir = os.path.join(BACKUP_DIR, f"v{current_version}")
        for rel in files:
            src = os.path.join(APP_DIR, rel)
            if os.path.isfile(src):
                b = os.path.join(bdir, rel)
                os.makedirs(os.path.dirname(b), exist_ok=True)
                shutil.copy2(src, b)

        # 4) подменяем файлы
        if progress:
            progress("Устанавливаю…")
        done = []
        try:
            for rel in files:
                dst = os.path.join(APP_DIR, rel)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                os.replace(os.path.join(tmp, rel), dst)
                done.append(rel)
        except Exception as e:
            # откатываем уже заменённые файлы
            for rel in done:
                b = os.path.join(bdir, rel)
                if os.path.isfile(b):
                    try:
                        shutil.copy2(b, os.path.join(APP_DIR, rel))
                    except Exception:
                        pass
            raise UpdateError(f"Не удалось заменить файлы ({e}). Всё возвращено как было.")

        # старый скомпилированный кэш больше не нужен
        shutil.rmtree(os.path.join(APP_DIR, "__pycache__"), ignore_errors=True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def restart():
    """Запускает программу заново и завершает текущий процесс."""
    import subprocess
    script = os.path.join(APP_DIR, "cleaner_qt.pyw")
    exe = sys.executable
    # всегда запускаем через pythonw.exe, чтобы не появлялось чёрное окно консоли
    if os.name == "nt" and os.path.basename(exe).lower() == "python.exe":
        w = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.isfile(w):
            exe = w
    flags = 0
    if os.name == "nt":
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([exe, script], cwd=APP_DIR, close_fds=True, creationflags=flags)
