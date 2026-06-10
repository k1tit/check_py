# -*- coding: utf-8 -*-
"""
Локальная SQLite-БД для выгрузок SOrg (обход «битого» xlsx в pandas).
По умолчанию для 3804 при запуске спрашиваем в консоли.
"""
from __future__ import annotations

import json
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from build_checks import BASE_DIR, _get_file, load_runtime_paths_dict

PARTS = ("base", "bp", "py", "zy")
DEFAULT_ASK_SORG = frozenset({"3804"})


def db_root() -> Path:
    return load_runtime_paths_dict()["data_dir"] / "db"


def db_path(folder: str) -> Path:
    return db_root() / f"{folder}.sqlite"


def _source_files(folder: str) -> dict[str, Path | None]:
    fp = BASE_DIR / folder
    return {
        "base": _get_file(fp, "*Base*.xlsx"),
        "bp": _get_file(fp, "*BP*.xlsx"),
        "py": _get_file(fp, "*PY*.xlsx"),
        "zy": _get_file(fp, "*ZY*.xlsx"),
    }


def _file_meta(path: Path | None) -> dict[str, Any] | None:
    if not path or not path.is_file():
        return None
    st = path.stat()
    return {
        "path": str(path.resolve()),
        "name": path.name,
        "size": st.st_size,
        "mtime": st.st_mtime,
    }


def build_manifest(folder: str) -> dict[str, Any]:
    src = _source_files(folder)
    return {
        "folder": folder,
        "built_at": datetime.now().isoformat(timespec="seconds"),
        "sources": {k: _file_meta(v) for k, v in src.items()},
    }


def _read_manifest(conn: sqlite3.Connection) -> dict[str, Any] | None:
    try:
        row = conn.execute("SELECT json FROM _manifest LIMIT 1").fetchone()
    except sqlite3.OperationalError:
        return None
    if not row:
        return None
    try:
        return json.loads(row[0])
    except (TypeError, ValueError):
        return None


def manifest_matches(folder: str) -> bool:
    path = db_path(folder)
    if not path.is_file():
        return False
    try:
        with sqlite3.connect(path) as conn:
            saved = _read_manifest(conn)
    except sqlite3.Error:
        return False
    if not saved:
        return False
    current = build_manifest(folder)
    return saved.get("sources") == current.get("sources")


def db_exists(folder: str) -> bool:
    path = db_path(folder)
    if not path.is_file():
        return False
    try:
        with sqlite3.connect(path) as conn:
            tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            return "base" in tables and _read_manifest(conn) is not None
    except sqlite3.Error:
        return False


def db_info(folder: str) -> str:
    path = db_path(folder)
    if not path.is_file():
        return "БД отсутствует"
    try:
        with sqlite3.connect(path) as conn:
            meta = _read_manifest(conn)
        built = (meta or {}).get("built_at", "?")
        stale = "" if manifest_matches(folder) else " (УСТАРЕЛА — xlsx изменились)"
        size_mb = path.stat().st_size / (1024 * 1024)
        return f"{path} ({size_mb:.1f} MB){stale}, залита {built}"
    except (OSError, sqlite3.Error):
        return f"{path} (ошибка чтения manifest)"


def _table_names(conn: sqlite3.Connection) -> set[str]:
    return {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def _read_table(conn: sqlite3.Connection, name: str) -> pd.DataFrame | None:
    if name not in _table_names(conn):
        return None
    df = pd.read_sql_query(f'SELECT * FROM "{name}"', conn, dtype=str)
    if df.empty:
        return None
    return df.fillna("")


def load_db(
    folder: str,
) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None] | None:
    if not db_exists(folder) or not manifest_matches(folder):
        return None
    path = db_path(folder)
    try:
        with sqlite3.connect(path) as conn:
            base = _read_table(conn, "base")
            if base is None:
                return None
            bp = _read_table(conn, "bp")
            py = _read_table(conn, "py")
            zy = _read_table(conn, "zy")
        print(f"[db] SO {folder}: загружено из {path}", flush=True)
        return base, bp, py, zy
    except Exception as exc:
        print(f"[db] SO {folder}: не удалось прочитать БД — {exc}", flush=True)
        return None


def save_db(
    folder: str,
    base: pd.DataFrame,
    bp: pd.DataFrame | None,
    py: pd.DataFrame | None,
    zy: pd.DataFrame | None,
) -> Path:
    path = db_path(folder)
    path.parent.mkdir(parents=True, exist_ok=True)
    manifest = json.dumps(build_manifest(folder), ensure_ascii=False, indent=2)

    def _prep(df: pd.DataFrame | None) -> pd.DataFrame | None:
        if df is None or df.empty:
            return None
        return df.fillna("").astype(str)

    with sqlite3.connect(path) as conn:
        for name, df in (("base", base), ("bp", bp), ("py", py), ("zy", zy)):
            prepared = _prep(df)
            if prepared is not None:
                prepared.to_sql(name, conn, if_exists="replace", index=False)
            else:
                conn.execute(f'DROP TABLE IF EXISTS "{name}"')
        conn.execute("CREATE TABLE IF NOT EXISTS _manifest (id INTEGER PRIMARY KEY, json TEXT)")
        conn.execute("DELETE FROM _manifest")
        conn.execute("INSERT INTO _manifest (json) VALUES (?)", (manifest,))
    print(f"[db] SO {folder}: сохранено в {path}", flush=True)
    return path


def clear_db(folder: str) -> None:
    path = db_path(folder)
    if path.is_file():
        path.unlink()
        print(f"[db] SO {folder}: БД удалена ({path})", flush=True)


def ask_use_db(folder: str) -> bool:
    if not db_exists(folder):
        return False
    if os.environ.get("REPORTS_DB_AUTO") == "1":
        return manifest_matches(folder)
    if os.environ.get("REPORTS_DB_AUTO") == "0":
        return False
    if not sys.stdin.isatty():
        return False
    stale = "" if manifest_matches(folder) else " [БД устарела — лучше перезалить]"
    print(f"[new_access] SO {folder}: найдена локальная БД{stale}", flush=True)
    print(f"  {db_info(folder)}", flush=True)
    try:
        ans = input(f"Использовать локальную БД для SO {folder} вместо чтения xlsx? [Y/n]: ").strip().lower()
    except EOFError:
        return False
    return ans in ("", "y", "yes", "д", "да")
