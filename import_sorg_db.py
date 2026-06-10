# -*- coding: utf-8 -*-
"""
Залить одну SOrg в локальную SQLite (data/db/{SO}.sqlite).
Для «битых» xlsx в pandas — чтение через Excel COM (Windows + pywin32).
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_checks import (  # noqa: E402
    BASE_DIR,
    _get_file,
    _normalize_base_df,
    _normalize_partner_df,
    _read_excel_checked,
    load_runtime_paths_dict,
)
from new_access_pf_checks import dedupe_base_access  # noqa: E402
from sorg_db import clear_db, db_info, save_db  # noqa: E402


def _file_mb(path: Path | None) -> str:
    if not path or not path.is_file():
        return "?"
    return f"{path.stat().st_size / (1024 * 1024):.1f} MB"


def _read_base_for_db(path: Path, folder: str, *, via_excel: bool):
    raw = _read_excel_checked(path, kind="Base", folder=folder, allow_com=via_excel)
    return dedupe_base_access(_normalize_base_df(raw, folder))


def _read_partner_for_db(path: Path, folder: str, label: str, *, via_excel: bool):
    try:
        return _read_partner_with_com(path, folder, label, via_excel=via_excel)
    except Exception:
        if not via_excel and sys.platform == "win32":
            print(f"  {label}: pandas не смог — повтор через Excel COM…", flush=True)
            return _read_partner_with_com(path, folder, label, via_excel=True)
        raise


def _read_partner_with_com(path: Path, folder: str, label: str, *, via_excel: bool):
    raw = _read_excel_checked(path, kind=label, folder=folder, allow_com=via_excel)
    return _normalize_partner_df(raw, folder, kind=label)


def load_sorg_for_db(folder: str, *, via_excel: bool = False):
    fp = BASE_DIR / folder
    f_base = _get_file(fp, "*Base*.xlsx")
    if not f_base:
        import pandas as pd

        return pd.DataFrame(), None, None, None

    f_bp = _get_file(fp, "*BP*.xlsx")
    f_py = _get_file(fp, "*PY*.xlsx")
    f_zy = _get_file(fp, "*ZY*.xlsx")

    print(f"  Base {f_base.name} ({_file_mb(f_base)})", flush=True)
    t0 = time.perf_counter()
    base = _read_base_for_db(f_base, folder, via_excel=via_excel)
    print(f"  Base: {len(base)} строк, {time.perf_counter() - t0:.0f} с", flush=True)

    def _part(path: Path | None, label: str):
        if not path:
            print(f"  {label}: не найден", flush=True)
            return None
        print(f"  {label} {path.name} ({_file_mb(path)})", flush=True)
        t1 = time.perf_counter()
        df = _read_partner_for_db(path, folder, label, via_excel=via_excel)
        rows = len(df) if df is not None and not df.empty else 0
        print(f"  {label}: {rows} строк, {time.perf_counter() - t1:.0f} с", flush=True)
        return df

    bp = _part(f_bp, "BP")
    py = _part(f_py, "PY")
    zy = _part(f_zy, "ZY")
    return base, bp, py, zy


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Залить SOrg в локальную SQLite (data/db/{SO}.sqlite)"
    )
    parser.add_argument("folder", help="Номер SOrg, например 3804")
    parser.add_argument("--clear", action="store_true", help="Удалить БД")
    parser.add_argument(
        "--via-excel",
        action="store_true",
        help="Всегда читать через Excel COM (медленно, но обходит BadZipFile)",
    )
    args = parser.parse_args()

    folder = args.folder.strip()
    paths = load_runtime_paths_dict()
    print(f"data_dir: {paths['data_dir']}")
    print(f"base_dir: {paths['base_dir']}")
    print(f"db:       {paths['data_dir'] / 'db' / (folder + '.sqlite')}")

    if not BASE_DIR.exists():
        print(f"Нет каталога выгрузок: {BASE_DIR}")
        return 1

    if args.clear:
        clear_db(folder)
        return 0

    os.environ["REPORTS_PARALLEL"] = "0"
    mode = "Excel COM" if args.via_excel else "pandas (BP/PY/ZY — fallback на Excel при ошибке)"
    print(f"\nЧитаю SO {folder} ({mode}) — может занять несколько минут…\n", flush=True)
    base, bp, py, zy = load_sorg_for_db(folder, via_excel=args.via_excel)
    if base.empty:
        print(f"Нет Base для SO {folder}")
        return 1

    save_db(folder, base, bp, py, zy)
    print(f"\n{db_info(folder)}")
    print("\nДальше: python new_access_pf_checks.py — при SO 3804 спросит, брать БД или xlsx.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
