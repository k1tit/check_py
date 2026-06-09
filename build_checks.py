# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional, Set, Tuple

import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


def _project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


_ROOT = _project_root()
RUNTIME_PATHS_JSON = _ROOT / "runtime_paths.json"
BUILD_PROGRESS_JSON = _ROOT / "build_progress.json"

ZERO_FILES_SUBDIR = "1 Нулевые файлы выгрузки макроса + файл исключений"
DATA_DIR = _ROOT / "data"


def _normalize_path_text(raw: object) -> str:
    txt = str(raw or "").strip().strip('"').strip("'")
    if not txt:
        return ""
    return str(Path(os.path.expandvars(os.path.expanduser(txt))))


def _resolve_config_path(raw: object) -> Path:
    txt = _normalize_path_text(raw)
    if not txt:
        return Path()
    p = Path(txt)
    if not p.is_absolute():
        p = _ROOT / p
    return p.resolve()


def _path_for_config(p: Path) -> str:
    p = p.resolve()
    try:
        return p.relative_to(_ROOT.resolve()).as_posix()
    except ValueError:
        return str(p)


def _paths_from_data_dir(data_dir: Path) -> dict[str, Path]:
    return {
        "data_dir": data_dir,
        "base_dir": data_dir / ZERO_FILES_SUBDIR,
        "output_dir": data_dir / "result",
        "exception_file": data_dir / "Exception.xlsx",
    }


def load_runtime_paths_dict() -> dict[str, Path]:
    """Все рабочие пути. Достаточно задать data_dir — остальное выводится автоматически."""
    result = _paths_from_data_dir(DATA_DIR)
    try:
        if RUNTIME_PATHS_JSON.is_file():
            cfg = json.loads(RUNTIME_PATHS_JSON.read_text(encoding="utf-8"))
            if cfg.get("data_dir"):
                result = _paths_from_data_dir(_resolve_config_path(cfg["data_dir"]))
            else:
                if cfg.get("base_dir"):
                    result["base_dir"] = _resolve_config_path(cfg["base_dir"])
                if cfg.get("output_dir"):
                    result["output_dir"] = _resolve_config_path(cfg["output_dir"])
                if cfg.get("exception_file"):
                    result["exception_file"] = _resolve_config_path(cfg["exception_file"])
                if result["base_dir"].name == ZERO_FILES_SUBDIR:
                    result["data_dir"] = result["base_dir"].parent
                elif cfg.get("base_dir"):
                    result["data_dir"] = result["base_dir"].parent
    except Exception:
        pass

    if os.environ.get("REPORTS_DATA_DIR"):
        result = _paths_from_data_dir(_resolve_config_path(os.environ["REPORTS_DATA_DIR"]))
    if os.environ.get("REPORTS_BASE_DIR"):
        result["base_dir"] = _resolve_config_path(os.environ["REPORTS_BASE_DIR"])
    if os.environ.get("REPORTS_OUTPUT_DIR"):
        result["output_dir"] = _resolve_config_path(os.environ["REPORTS_OUTPUT_DIR"])
    if os.environ.get("REPORTS_EXCEPTION_FILE"):
        result["exception_file"] = _resolve_config_path(os.environ["REPORTS_EXCEPTION_FILE"])
    if result["base_dir"].name == ZERO_FILES_SUBDIR:
        result["data_dir"] = result["base_dir"].parent
    return result


def save_runtime_paths(data_dir: str) -> dict[str, str]:
    paths = _paths_from_data_dir(_resolve_config_path(data_dir))
    payload = {"data_dir": _path_for_config(paths["data_dir"])}
    tmp = RUNTIME_PATHS_JSON.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(RUNTIME_PATHS_JSON)
    return {k: str(v) for k, v in paths.items()}


def _load_runtime_paths() -> tuple[Path, Path, Path]:
    d = load_runtime_paths_dict()
    return d["base_dir"], d["output_dir"], d["exception_file"]


BASE_DIR, OUTPUT_DIR, WEB_EXCEPTION_FILE = _load_runtime_paths()
#нормализация имен колонок
COLUMN_MAP = {
    "KUNNR": "Customer",
    "NAME1": "Name",
    "STCD": "Tax Number 1",
    "VKORD": "SO"
}

PAIRS = {
    "3801_3803": {"folders": ["3801", "3803"]},
    "3802_3804": {"folders": ["3802", "3804"]},
    "3805_3806": {"folders": ["3805", "3806"]},
}

NAME_BLACKLIST = (
    "sampling",
    "сэмплинг",
    "самлинг",
    "самплинг",
    "семплинг", 
    "dummy", 
    "дамми",
    "внутреннее", 
    "внп", 
    "bf_", 
    "test", 
    "тест",
)
# Как в Access (Joined_1): только <> "S" и <> "PR" для OrBlk и OrBlk1; TS не исключается.
BLOCKED_ORBLK = frozenset({"S", "PR"})
FAKE_INN = ("1111111111", "2222222222")

# Сообщение в Excel, если после apply_filters не осталось строк — не сбой, а допустимый исход.
NO_DATA_AFTER_FILTERS_MSG = (
    "Нет данных после применения фильтров: все строки отсеклись по правилам "
    "(исключения, CGrp≠Z, OrBlk, маски в имени, тестовые ИНН и т.д.). Это нормально."
)

BILL_TO_COLS = [
    "SO",
    "Customer",
    "Name",
    "Tax Number 1",
    "CGrp",
    "Grp4",
    "Search Term 2",
    "SDst",
    "A7",
    "OrBlk1",
    "OrBlk2",
    "Cust OrBlk Bill-to",
    "BP now",
    "Check Cust&BP",
    "BP SO",
    "BP Customer",
    "BP Name",
    "BP Tax Number 1",
    "BP CGrp",
    "BP OrBlk1",
    "BP OrBlk2",
    "BP status",
]

# Лист несоответствий — как эталон Excel / Joined_1_Checks (порядок столбцов 1:1 со скрином).
MISMATCH_EXPORT_COLS = [
    "SO",
    "Customer",
    "Name",
    "Tax Number 1",
    "CGrp",
    "Grp4",
    "Search Term 2",
    "SDst",
    "A7",
    "OrBlk1",
    "OrBlk2",
    "Cust OrBlk Bill-to",
    "BP",
    "BP Name",
    "BP Tax Number 1",
    "BP OrBlk1",
    "BP OrBlk2",
    "BP OrBlk Bill-to",
    "PY",
    "PY Name",
    "PY Tax Number 1",
    "PY OrBlk1",
    "PY OrBlk2",
    "ZY",
    "ZY Name",
    "ZY Tax Number 1",
    "ZY OrBlk1",
    "ZY OrBlk2",
    "Check BP-PY-ZY",
    "Check Tax Number Cust&BP",
    "Check Cust&BP",
    "Comment MD Analyst",
    "Комментарий OM",
]

# Поля bill-to b из pool2 в merge (могут быть шире BILL_TO_COLS — для _pref и полноты join).
_BILL_TO_B_SIDE_COLS = [
    "BP SO", "BP Customer", "BP Name", "BP Tax Number 1", "BP CGrp",
    "BP Grp4", "BP Search Term 2", "BP SDst", "BP A7",
    "BP OrBlk1", "BP OrBlk2",
]


def report_build_progress(percent: int, message: str, *, job_index: int = 1, job_total: int = 1) -> None:
    payload = {"percent": max(0, min(100, int(percent))), "message": message,
               "job_index": job_index, "job_total": job_total}
    try:
        tmp = BUILD_PROGRESS_JSON.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        tmp.replace(BUILD_PROGRESS_JSON)
    except OSError:
        pass


def _norm(v: object) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def _norm_cust(v: object) -> str:
    s = _norm(v)
    if s.isdigit():
        return s.lstrip("0") or "0"
    return s


def _ns(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)


def _nc(s: pd.Series) -> pd.Series:
    base = _ns(s)
    return base.where(~base.str.match(r"^\d+$"), base.str.lstrip("0").replace("", "0"))


def _align_merge_keys(df: pd.DataFrame) -> pd.DataFrame:
    """
    Приводаит Customer и _folder к согласованным строковым ключам.
    Иначе pandas.merge падает с int64 vs object — цепочка рвётся, отчёт без данных.
    """
    if df is None or df.empty:
        return df
    out = df.copy()
    if "Customer" in out.columns:
        out["Customer"] = _nc(out["Customer"].fillna("").astype(str))
    if "_folder" in out.columns:
        out["_folder"] = out["_folder"].map(_norm)
    return out


def _col(df: pd.DataFrame, *candidates: str) -> str | None:
    lo = {c.casefold(): c for c in df.columns}
    for c in candidates:
        found = lo.get(c.casefold())
        if found is not None:
            return found
    return None


def _so_col(df: pd.DataFrame) -> str | None:
    return _col(df, "SOrg#", "SOrg.", "VKORG", "SO", "Sales Org.")


def _partner_lookup_extra_columns(base_lookup: pd.DataFrame, prefix: str) -> dict[str, str]:
    """
    Как в Access join к BP/PY/ZY: только поля master из Base для номера партнёра (без адресной «витрины»).
    """
    extra: dict[str, str] = {}
    logical_maps: list[tuple[tuple[str, ...], str]] = [
        (("Name",), f"{prefix} Name"),
        (("Tax Number 1", "STCD"), f"{prefix} Tax Number 1"),
        (("OrBlk",), f"{prefix} OrBlk1"),
        (("OrBlk1", "OrBlk 1"), f"{prefix} OrBlk2"),
    ]
    for candidates, dst in logical_maps:
        src = _col(base_lookup, *candidates)
        if src and src not in extra:
            extra[src] = dst
    return extra


def _get_file(folder: Path, pattern: str) -> Path | None:
    if not folder.exists():
        return None
    files = list(folder.glob(pattern))
    return files[0] if files else None


_CYR = re.compile(r"[А-Яа-яЁё]")


def _read_base(path: Path, folder: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    df = df.rename(columns={"OrBlk.1": "OrBlk1"})
    so = _so_col(df)
    if so and so != "SO":
        df = df.rename(columns={so: "SO"})
    cust = _col(df, "Customer", "KUNNR")
    if cust and cust != "Customer":
        df = df.rename(columns={cust: "Customer"})
    if "Customer" in df.columns:
        df["Customer"] = _nc(df["Customer"])
    for old, new in COLUMN_MAP.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})
    g4 = _col(df, "Grp4", "GROUP4", "Group4", "GRP4")
    if g4 and g4 != "Grp4":
        df = df.rename(columns={g4: "Grp4"})
    df["_folder"] = folder
    return df


def dedupe_base(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    # DelDup: в одной выгрузке не схлопывать разные строки одного KUNNR с разным ИНН
    # (иначе теряются ship-to для Bill-to / несоответствия).
    if "_folder" in df.columns and "Customer" in df.columns:
        if "Tax Number 1" in df.columns:
            group_keys: list[str] = ["_folder", "Customer", "Tax Number 1"]
        else:
            group_keys = ["_folder", "Customer"]
    elif "SO" in df.columns and "Customer" in df.columns:
        if "Tax Number 1" in df.columns:
            group_keys = ["SO", "Customer", "Tax Number 1"]
        else:
            group_keys = ["SO", "Customer"]
    else:
        group_keys = ["Customer"]
    rows = []
    for _, grp in df.groupby(group_keys, sort=False, dropna=False):
        row = grp.iloc[0].to_dict()
        if "Name" in grp.columns:
            names = [s.strip() for s in grp["Name"].fillna("").astype(str) if s.strip()]
            cyr = [n for n in names if _CYR.search(n)]
            row["Name"] = min(cyr) if cyr else (min(names) if names else "")
        rows.append(row)
    result = pd.DataFrame(rows)
    for c in df.columns:
        if c not in result.columns:
            result[c] = ""
    return result[df.columns].reset_index(drop=True)


def _read_partner(path: Path, folder: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    kunnr = _col(df, "KUNNR", "Customer", "Sold-to")
    ktonr = _col(df, "KTONR", "BP", "PY", "ZY")
    if not kunnr or not ktonr:
        return pd.DataFrame(columns=["KUNNR", "KTONR", "_folder"])
    out = df[[kunnr, ktonr]].rename(columns={kunnr: "KUNNR", ktonr: "KTONR"})
    out["KUNNR"] = _nc(out["KUNNR"])
    out["KTONR"] = _nc(out["KTONR"])
    out["_folder"] = folder
    return out.dropna(subset=["KUNNR"]).drop_duplicates(subset=["KUNNR", "_folder"], keep="first")


def merge_partner(base_lookup: pd.DataFrame, partner_df: pd.DataFrame | None, prefix: str) -> pd.DataFrame:
    keep = {"Customer": prefix}
    extra = _partner_lookup_extra_columns(base_lookup, prefix)
    keys = list(keep.keys()) + list(extra.keys())
    bk = base_lookup[keys].rename(columns={**keep, **extra})
    bk = bk.drop_duplicates(subset=[prefix], keep="first")

    extra_cols = list(extra.values())
    if partner_df is None or partner_df.empty:
        return pd.DataFrame(columns=["Customer", "_folder", prefix] + extra_cols)

    if "_folder" not in partner_df.columns:
        partner_df = partner_df.copy()
        partner_df["_folder"] = ""
    pf = partner_df[["KUNNR", "KTONR", "_folder"]].rename(columns={"KUNNR": "Customer", "KTONR": prefix})
    pf = _align_merge_keys(pf)
    pf[prefix] = _nc(pf[prefix].fillna("").astype(str))
    bk = bk.copy()
    bk[prefix] = _nc(bk[prefix].fillna("").astype(str))
    merged = pf.merge(bk, on=prefix, how="left")
    merged = _align_merge_keys(merged)
    return merged.drop_duplicates(subset=["Customer", "_folder"], keep="first")


def merge_all_partners(base: pd.DataFrame, bp_df, py_df, zy_df) -> pd.DataFrame:
    result = _align_merge_keys(base.copy())
    for prefix, pf_df in [("BP", bp_df), ("PY", py_df), ("ZY", zy_df)]:
        lk = merge_partner(base, pf_df, prefix)
        lk = _align_merge_keys(lk)
        merge_on = ["Customer", "_folder"] if "_folder" in result.columns and "_folder" in lk.columns else "Customer"
        try:
            result = result.merge(lk, on=merge_on, how="left")
        except ValueError as e:
            print(f"[build_checks] merge {prefix}: ключи не сходятся по типам, пробуем строковый привод: {e}", flush=True)
            result = _align_merge_keys(result)
            lk = _align_merge_keys(lk)
            result = result.merge(lk, on=merge_on, how="left")

        if prefix not in result.columns:
            result[prefix] = ""
        else:
            result[prefix] = (
                result[prefix].fillna("").astype(str).str.strip()
                .replace({"nan": "", "None": "", "<NA>": ""})
            )

        extra_map = _partner_lookup_extra_columns(base, prefix)
        for src_col, dst in extra_map.items():
            if dst not in result.columns:
                result[dst] = ""
            result[dst] = (
                result[dst].fillna("").astype(str).str.strip()
                .replace({"nan": "", "None": "", "<NA>": ""})
            )
    return result


def load_data(folders: list[str]) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    bases, bps, pys, zys = [], [], [], []
    for folder in folders:
        fp = BASE_DIR / folder
        f = _get_file(fp, "*Base*.xlsx")
        if f:
            bases.append(_read_base(f, folder))
        for lst, pat in [(bps, "*BP*.xlsx"), (pys, "*PY*.xlsx"), (zys, "*ZY*.xlsx")]:
            f = _get_file(fp, pat)
            if f:
                lst.append(_read_partner(f, folder))
    if not bases:
        return pd.DataFrame(), None, None, None
    base = dedupe_base(pd.concat(bases, ignore_index=True))

    def _cat(parts):
        if not parts:
            return None
        df = pd.concat(parts, ignore_index=True)
        return df.drop_duplicates(subset=["KUNNR", "_folder"], keep="first")

    return base, _cat(bps), _cat(pys), _cat(zys)


def load_exception(base_dir: Path) -> pd.DataFrame:
    parts = []
    for f in base_dir.rglob("Exception.xlsx"):
        try:
            parts.append(pd.read_excel(f, dtype=str))
        except Exception:
            pass
    if WEB_EXCEPTION_FILE.exists():
        try:
            parts.append(pd.read_excel(WEB_EXCEPTION_FILE, dtype=str))
        except Exception:
            pass
    if not parts:
        return pd.DataFrame(columns=["SO", "Customer", "Comment OM"])
    df = pd.concat(parts, ignore_index=True)
    so_c = _col(df, "SO", "SOrg#", "SOrg.", "VKORG")
    cu_c = _col(df, "Customer", "KUNNR")
    if not so_c or not cu_c:
        return pd.DataFrame(columns=["SO", "Customer", "Comment OM"])
    co_c = _col(df, "Comment OM", "Comment", "Комментарий")
    cols: dict[str, str] = {so_c: "SO", cu_c: "Customer"}
    if co_c:
        cols[co_c] = "Comment OM"
    out = df.rename(columns=cols)[list(cols.values())].copy()
    if "Comment OM" not in out.columns:
        out["Comment OM"] = ""
    out["SO"] = _ns(out["SO"])
    out["Customer"] = _nc(out["Customer"])
    out = out[(out["SO"] != "") & (out["Customer"] != "")]
    return out.drop_duplicates(subset=["SO", "Customer"], keep="last").reset_index(drop=True)


def exception_keys(exc_df: pd.DataFrame) -> Set[Tuple[str, str]]:
    if exc_df.empty:
        return set()
    return {(_norm(r["SO"]), _norm_cust(r["Customer"])) for _, r in exc_df.iterrows()}


def apply_filters(df: pd.DataFrame, exc_keys: Set[Tuple[str, str]]) -> pd.DataFrame:
    if df.empty:
        return df
    # Exception exclusion
    if exc_keys:
        so_s = _ns(df.get("SO", pd.Series("", index=df.index)))
        cu_s = _nc(df["Customer"])
        keep = pd.Series([(_norm(s), _norm_cust(c)) not in exc_keys
                          for s, c in zip(so_s, cu_s)], index=df.index)
        df = df[keep].copy()
    if df.empty:
        return df

    name_s = df["Name"].fillna("").astype(str).str.lower() if "Name" in df.columns else pd.Series("", index=df.index)
    mask_name = ~name_s.apply(lambda n: any(b in n for b in NAME_BLACKLIST))

    cgrp_s = df["CGrp"].fillna("").astype(str).str.upper().str.strip() if "CGrp" in df.columns else pd.Series("", index=df.index)
    mask_cgrp = cgrp_s != "Z"

    def _ob_series(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series("", index=df.index)
        t = df[col].fillna("").astype(str).str.strip().str.upper()
        return t.replace({"NAN": "", "<NA>": "", "NONE": ""})

    orblk_s = _ob_series("OrBlk")
    orblk1_s = _ob_series("OrBlk1")
    mask_orblk = ~orblk_s.isin(BLOCKED_ORBLK) & ~orblk1_s.isin(BLOCKED_ORBLK)

    tax_s = df["Tax Number 1"].fillna("").astype(str).str.strip() if "Tax Number 1" in df.columns else pd.Series("", index=df.index)
    mask_inn = ~tax_s.apply(lambda t: any(f in t for f in FAKE_INN))

    return df[mask_name & mask_cgrp & mask_orblk & mask_inn].copy()


def add_checks(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    def _s(col: str) -> pd.Series:
        return out[col].fillna("").astype(str).str.strip() if col in out.columns else pd.Series("", index=out.index)

    if "OrBlk" in out.columns and "OrBlk1" not in out.columns:
        out = out.rename(columns={"OrBlk": "OrBlk1"})
    elif "OrBlk" in out.columns and "OrBlk1" in out.columns:
        out = out.rename(columns={"OrBlk1": "OrBlk2", "OrBlk": "OrBlk1"})

    orblk1 = _s("OrBlk1").str.upper()
    bp_orblk1 = _s("BP OrBlk1").str.upper()
    bp = _s("BP")
    py = _s("PY")
    zy = _s("ZY")
    cust = _s("Customer")
    tax = _s("Tax Number 1")
    bp_tax = _s("BP Tax Number 1")
    bp_name = _s("BP Name")

    out["Cust OrBlk Bill-to"] = np.where(orblk1 == "M", "Bill-to", "Ship-to")
    out["BP OrBlk Bill-to"] = np.where(bp_orblk1 == "M", "Bill-to", "Ship-to")
    out["Check BP-PY-ZY"] = np.where((bp == py) & (bp == zy), "TRUE", "FALSE")
    out["Check Tax Number Cust&BP"] = np.where(tax == bp_tax, "TRUE", "FALSE")
    out["Check Cust&BP"] = np.where(cust == bp, "TRUE", "FALSE")

    bp_mismatch = (bp != py) | (bp != zy)
    tax_mismatch = tax != bp_tax
    cust_ne_bp = cust != bp

    comment = pd.Series("", index=out.index)
    comment = comment.mask(bp_mismatch & ~tax_mismatch, "Несоответствие BP-PY-ZY")
    comment = comment.mask(bp_mismatch & tax_mismatch, "Несоответствие BP-PY-ZY; несоответствие ИНН SP и BP")
    comment = comment.mask(~bp_mismatch & tax_mismatch, "Несоответствие ИНН Cust и BP")
    comment = comment.mask(~bp_mismatch & ~tax_mismatch & (bp_name == ""), "BP помечен на удаление")
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk1 == "M") & cust_ne_bp,
        "Bill-to прикреплён к другому Bill-to"
    )
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk1 != "M") & (bp_orblk1 != "M") & cust_ne_bp,
        "Ship-to прикреплён к BP без OB M"
    )
    out["Comment MD Analyst"] = comment
    return out


def build_bill_to(df_checks: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    if df_checks.empty or base.empty:
        return pd.DataFrame()

    pool1 = df_checks[
        (df_checks["Cust OrBlk Bill-to"] == "Ship-to") &
        (df_checks["Check Cust&BP"] == "TRUE") &
        (df_checks["Comment MD Analyst"] == "")
    ].copy()
    if pool1.empty:
        print("[build_checks] build_bill_to: pool1 пуст (нет Ship-to + Check Cust&BP=TRUE без комментария)", flush=True)
        return pd.DataFrame()

    # SO для листов и фильтра: из колонки SO или из папки выгрузки
    pool1 = pool1.copy()
    if "_folder" in pool1.columns:
        so_series = _ns(pool1["SO"]) if "SO" in pool1.columns else pd.Series("", index=pool1.index, dtype=object)
        need = so_series.eq("") | so_series.isna()
        if need.any():
            if "SO" not in pool1.columns:
                pool1["SO"] = ""
            pool1.loc[need, "SO"] = _ns(pool1.loc[need, "_folder"])

    # Первый OB в Base (как b.OrBlk в Access), не OrBlk1 — второй уровень.
    orblk_col = _col(base, "OrBlk", "OrBlk 1")
    if not orblk_col:
        print("[build_checks] build_bill_to: в Base нет колонки OrBlk — лист Bill-to пуст", flush=True)
        return pd.DataFrame()
    pool2 = base[base[orblk_col].fillna("").astype(str).str.upper().str.strip() == "M"].copy()
    if pool2.empty:
        print("[build_checks] build_bill_to: pool2 пуст (нет строк с OrBlk=M в Base)", flush=True)
        return pd.DataFrame()

    so_c = _so_col(pool2)
    rename_p2: dict[str, str] = {"Customer": "BP Customer", orblk_col: "BP OrBlk1"}
    if so_c:
        rename_p2[so_c] = "BP SO"
    extra_map = {
        "Name": "BP Name",
        "Tax Number 1": "BP Tax Number 1",
        "CGrp": "BP CGrp",
    }
    for src, dst in extra_map.items():
        if src in pool2.columns:
            rename_p2[src] = dst
    g4 = _col(pool2, "Grp4", "GROUP4", "Group4")
    if g4:
        rename_p2[g4] = "BP Grp4"
    st2 = _col(pool2, "Search Term 2", "Search term 2", "SORT2", "Sort term 2")
    if st2 and st2 not in rename_p2:
        rename_p2[st2] = "BP Search Term 2"
    sdst = _col(pool2, "SDst", "SDIST", "Sdist")
    if sdst:
        rename_p2[sdst] = "BP SDst"
    a7 = _col(pool2, "A7")
    if a7:
        rename_p2[a7] = "BP A7"
    ob2 = _col(pool2, "OrBlk1", "OrBlk 1", "OrBlk2", "OrBlk 2")
    if ob2:
        rename_p2[ob2] = "BP OrBlk2"
    pool2 = pool2.rename(columns=rename_p2)

    required_p2_cols = [
        "BP Customer", "BP Name", "BP Tax Number 1", "BP OrBlk1", "BP OrBlk2", "BP CGrp",
        "BP Grp4", "BP Search Term 2", "BP SDst", "BP A7",
    ]
    for col in required_p2_cols:
        if col not in pool2.columns:
            pool2[col] = ""
    if "BP SO" not in pool2.columns:
        pool2["BP SO"] = ""
    # BP SO из SOrg в файле; иначе фильтр «та же SOrg» гасит все строки (сравнение с пустым BP SO)
    bpso = pool2["BP SO"].fillna("").astype(str).map(_norm) if "BP SO" in pool2.columns else pd.Series("", index=pool2.index)
    if "_folder" in pool2.columns:
        miss = bpso.eq("")
        if miss.any():
            pool2 = pool2.copy()
            pool2.loc[miss, "BP SO"] = pool2.loc[miss, "_folder"].map(_norm)

    pool1["_tax"] = _ns(pool1["Tax Number 1"]) if "Tax Number 1" in pool1.columns else ""
    pool2["_tax"] = _ns(pool2["BP Tax Number 1"]) if "BP Tax Number 1" in pool2.columns else ""

    # Колонки Bill-to кандидата b (как в Access j INNER JOIN b). В pool1 уже есть одноимённые
    # BP SO, BP OrBlk1, … из строки j — иначе merge даёт BP OrBlk2_x/_y и «BP OrBlk2» остаётся пустым.
    pool1_for_join = pool1.drop(columns=[c for c in _BILL_TO_B_SIDE_COLS if c in pool1.columns], errors="ignore")

    p2_cols = ["_tax"] + _BILL_TO_B_SIDE_COLS
    existing_p2_cols = [c for c in p2_cols if c in pool2.columns]
    merge_on = ["_tax"]
    p2_use = list(existing_p2_cols)
    # Access: INNER JOIN только по Tax Number 1 (SO / _folder не в ключе join).

    p2_sub = pool2[p2_use]
    joined = pool1_for_join.merge(p2_sub, on=merge_on, how="inner")
    joined = joined.drop(columns=["_tax"], errors="ignore")
    if "BP" in joined.columns:
        joined = joined.rename(columns={"BP": "BP now"})

    print(
        f"[build_checks] build_bill_to: pool1={len(pool1)} pool2={len(pool2)} joined={len(joined)} "
        f"(merge по {merge_on})",
        flush=True,
    )
    if joined.empty and len(pool1) > 0 and len(pool2) > 0:
        print(
            "[build_checks] build_bill_to: merge по ИНН пуст — нет совпадения Tax Number 1 (ship-to vs Bill-to M)",
            flush=True,
        )

    # Заполняем отсутствующие колонки в joined (все из _BILL_TO_B_SIDE_COLS + BP SO уже в required_p2_cols / ниже)
    for col in required_p2_cols + ["BP SO"]:
        if col not in joined.columns:
            joined[col] = ""

    # Access размножает строки после join — без groupby/drop_duplicates.

    ob1 = joined["BP OrBlk1"].fillna("").astype(str).str.strip().str.upper()
    ob2s = joined["BP OrBlk2"].fillna("").astype(str).str.strip()
    joined["BP status"] = np.select(
        [(ob1 == "M") & (ob2s == ""), (ob1 == "M") & (ob2s.str.upper() == "M")],
        ["Ошибка для OB", "active"],
        default="block",
    )

    status_order = {"active": 0, "block": 1, "Ошибка для OB": 2}
    joined["_st"] = joined["BP status"].map(status_order).fillna(1).astype(int)
    sort_cols = [c for c in ["Tax Number 1", "Customer", "_st", "BP Customer"] if c in joined.columns]
    if sort_cols:
        joined = joined.sort_values(sort_cols, kind="stable").drop(columns=["_st"], errors="ignore")

    return _prep_bill_to_sheet(joined)

def collect_and_persist_global_exception(base_dir: Path, output_dir: Path) -> pd.DataFrame:
    exc = load_exception(base_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    if exc.empty:
        try:
            (output_dir / "00_Exception_собрано_нет_данных.txt").write_text(
                "Исключения не загружены\n", encoding="utf-8")
        except Exception:
            pass
    else:
        try:
            exc.to_excel(output_dir / "00_Exception_собрано_перед_отчётами.xlsx",
                         index=False, sheet_name="Exception")
        except Exception:
            pass
    return exc


def _prep_bill_to_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Шапка листа «Привязка Bill-to по ИНН» — q_*_Joined_3_BillToByINN."""
    if df is None or df.empty:
        return pd.DataFrame(columns=BILL_TO_COLS)

    out = df.copy()
    if "Grp4" not in out.columns:
        g4 = _col(out, "Grp4", "GROUP4", "Group4", "GRP4")
        if g4:
            out = out.rename(columns={g4: "Grp4"})
    if "Grp4" not in out.columns:
        out["Grp4"] = ""

    if "BP now" not in out.columns and "BP" in out.columns:
        out = out.rename(columns={"BP": "BP now"})

    if "_folder" in out.columns:
        so_series = _ns(out["SO"]) if "SO" in out.columns else pd.Series("", index=out.index, dtype=object)
        need = so_series.eq("") | so_series.isna()
        if need.any():
            if "SO" not in out.columns:
                out["SO"] = ""
            out.loc[need, "SO"] = _ns(out.loc[need, "_folder"])

    for col in BILL_TO_COLS:
        if col not in out.columns:
            out[col] = ""
    return out[BILL_TO_COLS]


def _excel_sheet_name_safe(name: str) -> str:
    """Имя листа Excel ≤31 символа, без : \\ / ? * [ ]"""
    s = re.sub(r'[:\\/?*\[\]]', " ", str(name)).strip()
    return (s[:31] if s else "Sheet").strip() or "Sheet"


def _unique_excel_sheet_name(proposed: str, used: set[str]) -> str:
    base = _excel_sheet_name_safe(proposed)
    if base not in used:
        return base
    for i in range(2, 999):
        suf = f" {i}"
        cand = _excel_sheet_name_safe(base[: max(1, 31 - len(suf))] + suf)
        if cand not in used:
            return cand
    return _excel_sheet_name_safe(base[:28] + "...")[:31]


def _bill_to_sheet_label(so_token: str) -> str:
    """Как в ТЗ: «3805 Привязка Bill-to по ИНН»."""
    so = _norm(so_token)
    if not so:
        so = "SO"
    return f"{so} Привязка Bill-to по ИНН"


def _bill_to_rows_for_sorg(bt: pd.DataFrame, so_key: str) -> pd.DataFrame:
    """
    Строки для листа «{SOrg} Привязка Bill-to по ИНН»:
    SO = SOrg листа (как j.[SO] в Access).
    BP SO: если пусто — не режем; если совпадает с листом — оставляем; если не совпадает, но строка
    из выгрузки этой же орг. (_folder == лист), всё равно оставляем — в Access join к b по ИНН без
    фильтра по VKORG bill-to; у части ship-to в master BP другой SOrg, иначе «пропадают» единичные ИНН.
    """
    if bt.empty or not so_key:
        return pd.DataFrame()
    if "SO" not in bt.columns:
        return pd.DataFrame()
    m = _ns(bt["SO"]) == so_key
    if "BP SO" in bt.columns:
        bso = _ns(bt["BP SO"])
        bp_empty = bt["BP SO"].fillna("").astype(str).str.strip().eq("")
        same_folder = (
            (_ns(bt["_folder"]) == so_key)
            if "_folder" in bt.columns
            else pd.Series(False, index=bt.index)
        )
        m &= (bso == so_key) | bp_empty | same_folder
    return bt[m].copy()


def add_bill_to_sheets(
    writer: pd.ExcelWriter,
    bill_to_df: pd.DataFrame,
    sorg_folders: list[str] | None,
    *,
    placeholder_message: str | None = None,
) -> None:
    """
    Один лист на каждую SOrg из задачи (пара или группа), даже без строк Bill-to.
    Отбор строк: SO = SOrg листа; BP SO = SOrg или пусто или строка из выгрузки этой SOrg (_folder),
    иначе часть Bill-to по ИНН пропадает на листе (см. _bill_to_rows_for_sorg).
    Фильтр по полным колонкам — до _prep_bill_to_sheet, чтобы не терять _folder.
    sorg_folders — номера папок из process_pair (например ['3805','3806']).
    placeholder_message — если задано, подставляется вместо «Нет данных для SO …» на пустых листах.
    """
    used: set[str] = set()

    if sorg_folders:
        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_bill_to_sheet_label(so_key), used)
            used.add(title)
            grp_raw = _bill_to_rows_for_sorg(bill_to_df, so_key) if not bill_to_df.empty else pd.DataFrame()
            grp = _prep_bill_to_sheet(grp_raw) if not grp_raw.empty else pd.DataFrame(columns=BILL_TO_COLS)
            if grp.empty:
                pd.DataFrame(columns=BILL_TO_COLS).to_excel(writer, sheet_name=title, index=False)
            else:
                grp.to_excel(writer, sheet_name=title, index=False)
        return

    if not bill_to_df.empty:
        bt2 = bill_to_df.copy()
        if "SO" in bt2.columns:
            for so_val, grp_raw in bt2.groupby("SO", dropna=False):
                so_key = _norm(so_val) or "ALL"
                title = _unique_excel_sheet_name(_bill_to_sheet_label(so_key), used)
                used.add(title)
                grp_f = _bill_to_rows_for_sorg(grp_raw, so_key)
                grp = _prep_bill_to_sheet(grp_f) if not grp_f.empty else pd.DataFrame(columns=BILL_TO_COLS)
                if grp.empty:
                    pd.DataFrame(columns=BILL_TO_COLS).to_excel(writer, sheet_name=title, index=False)
                else:
                    grp.to_excel(writer, sheet_name=title, index=False)
        else:
            _prep_bill_to_sheet(bt2).to_excel(
                writer, sheet_name=_excel_sheet_name_safe("Привязка Bill-to по ИНН"), index=False
            )
    else:
        pd.DataFrame(columns=BILL_TO_COLS).to_excel(
            writer, sheet_name=_excel_sheet_name_safe("Привязка Bill-to по ИНН"), index=False
        )


def _attach_comment_om(df: pd.DataFrame, exc_df: pd.DataFrame) -> pd.DataFrame:
    """Подмешать Comment OM из Exception по нормализованным SO + Customer."""
    if df.empty:
        return df
    if exc_df is None or exc_df.empty:
        out = df.copy()
        if "Comment OM" not in out.columns:
            out["Comment OM"] = ""
        return out

    left = df.copy()
    if "Comment OM" in left.columns:
        left = left.drop(columns=["Comment OM"])

    ex = exc_df[["SO", "Customer", "Comment OM"]].copy()
    ex["_j_so"] = _ns(ex["SO"])
    ex["_j_cu"] = _nc(ex["Customer"])
    ex = ex.drop_duplicates(subset=["_j_so", "_j_cu"], keep="last")[["_j_so", "_j_cu", "Comment OM"]]

    left["_j_so"] = _ns(left["SO"]) if "SO" in left.columns else ""
    left["_j_cu"] = _nc(left["Customer"])
    merged = left.merge(ex, on=["_j_so", "_j_cu"], how="left")
    merged = merged.drop(columns=["_j_so", "_j_cu"])
    merged["Comment OM"] = merged["Comment OM"].fillna("")
    return merged


def _dedupe_mismatch_report_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Одна строка отчёта на комбинацию (орг. выгрузка, ИНН, BP, текст Comment MD Analyst).
    Одинаковая проблема по многим ship-to не раздувает список; одна и та же проблема в 3801 и 3803 —
    остаётся две строки (разные _folder).
    Оставляем первую строку после сортировки по SO, Customer (стабильный «представитель»).
    """
    if df.empty:
        return df
    if not all(k in df.columns for k in ("Tax Number 1", "BP", "Comment MD Analyst")):
        return df
    work = df.copy()
    sort_cols = [c for c in ("SO", "Customer") if c in work.columns]
    if sort_cols:
        work = work.sort_values(by=sort_cols, kind="stable")
    work["_d_tax"] = work["Tax Number 1"].fillna("").astype(str).str.strip()
    work["_d_bp"] = work["BP"].fillna("").astype(str).map(_norm_cust)
    work["_d_cm"] = work["Comment MD Analyst"].fillna("").astype(str).str.strip()
    dedupe_cols = ["_d_tax", "_d_bp", "_d_cm"]
    if "_folder" in work.columns:
        work["_d_fd"] = work["_folder"].fillna("").astype(str).map(_norm)
        dedupe_cols.insert(0, "_d_fd")
    out = work.drop_duplicates(subset=dedupe_cols, keep="first").reset_index(drop=True)
    return out.drop(columns=[c for c in ("_d_tax", "_d_bp", "_d_cm", "_d_fd") if c in out.columns])


def _mismatch_export_value(row: pd.Series, export_col: str) -> str:
    """Значение для колонки экспорта; имена в Excel могут отличаться от внутренних (checked)."""
    src_cols: tuple[str, ...]
    if export_col == "Check Tax Number Cust&BP":
        src_cols = ("Check Tax Number Cust&BP", "Check Tax Number Cust-BP")
    elif export_col == "Комментарий OM":
        src_cols = ("Комментарий OM", "Comment OM", "Comment")
    else:
        src_cols = (export_col,)

    for src in src_cols:
        if src not in row.index:
            continue
        v = row[src]
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        s = str(v).strip()
        if s.lower() in ("nan", "<na>", "none"):
            continue
        return s
    return ""


def _prep_mismatch_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Строки = результат add_checks (filtered); шапка как Joined_1_Checks, не витрина адресов."""
    if df is None or df.empty:
        return pd.DataFrame(columns=MISMATCH_EXPORT_COLS)

    out = df.copy()
    if "Grp4" not in out.columns:
        g4 = _col(out, "Grp4", "GROUP4", "Group4", "GRP4")
        if g4:
            out = out.rename(columns={g4: "Grp4"})
    if "Grp4" not in out.columns:
        out["Grp4"] = ""

    if "_folder" in out.columns:
        so_series = _ns(out["SO"]) if "SO" in out.columns else pd.Series("", index=out.index, dtype=object)
        need = so_series.eq("") | so_series.isna()
        if need.any():
            if "SO" not in out.columns:
                out["SO"] = ""
            out.loc[need, "SO"] = _ns(out.loc[need, "_folder"])

    rows: list[dict[str, str]] = []
    for _, row in out.iterrows():
        rows.append({c: _mismatch_export_value(row, c) for c in MISMATCH_EXPORT_COLS})

    return pd.DataFrame(rows, columns=MISMATCH_EXPORT_COLS)


def _mismatch_sheet_label(so_token: str) -> str:
    """Как у Bill-to: префикс SOrg + тема листа."""
    so = _norm(so_token)
    if not so:
        so = "SO"
    return f"{so} Несоответствия"


def add_mismatch_sheets(
    writer: pd.ExcelWriter,
    errors_df: pd.DataFrame,
    sorg_folders: list[str] | None,
    *,
    placeholder_message: str | None = None,
) -> None:
    """
    Один лист на каждую SOrg из задачи. Шапка как Joined_1_Checks (строки из checked / errors_only).
    Разрез по листу: при наличии _folder — только строки этой выгрузки (иначе SO в файле часто пустой/не тот).
    """
    used: set[str] = set()

    def _msg_empty_so(so_key: str) -> str:
        if placeholder_message:
            return placeholder_message
        return f"Нет несоответствий для SO {so_key}"

    def _slice_errors_for_sorg(folder_token: str) -> pd.DataFrame:
        if errors_df.empty:
            return errors_df
        so_key = _norm(folder_token)
        if not so_key:
            return errors_df
        if "_folder" in errors_df.columns:
            m = errors_df["_folder"].fillna("").astype(str).map(_norm) == so_key
            return errors_df[m].copy()
        if "SO" in errors_df.columns:
            return errors_df[_ns(errors_df["SO"]) == so_key].copy()
        return errors_df

    if sorg_folders:
        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_mismatch_sheet_label(so_key), used)
            used.add(title)
            sub = _slice_errors_for_sorg(folder_so)
            prep = _prep_mismatch_sheet(sub)
            if prep.empty:
                pd.DataFrame({"Сообщение": [_msg_empty_so(so_key)]}).to_excel(
                    writer, sheet_name=title, index=False
                )
            else:
                prep.to_excel(writer, sheet_name=title, index=False)
        return

    prep = _prep_mismatch_sheet(errors_df)
    if not prep.empty:
        if "SO" in prep.columns:
            for so_val, grp in prep.groupby("SO", dropna=False):
                so_key = _norm(so_val) or "ALL"
                title = _unique_excel_sheet_name(_mismatch_sheet_label(so_key), used)
                used.add(title)
                grp.to_excel(writer, sheet_name=title, index=False)
        else:
            prep.to_excel(writer, sheet_name=_excel_sheet_name_safe("Несоответствия"), index=False)
    else:
        msg = placeholder_message or "Нет данных"
        pd.DataFrame({"Сообщение": [msg]}).to_excel(
            writer, sheet_name=_excel_sheet_name_safe("Несоответствия"), index=False
        )


def save_excel(
    pair_name: str,
    errors_df: pd.DataFrame,
    bill_to_df: pd.DataFrame,
    exc_df: pd.DataFrame,
    *,
    bill_to_sorg_folders: list[str] | None = None,
    after_filters_empty: bool = False,
) -> None:
    pair_dir = OUTPUT_DIR / pair_name
    pair_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%d.%m.%Y")
    out_file = pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str} - Необходимый итоговый файл.xlsx"

    ph = NO_DATA_AFTER_FILTERS_MSG if after_filters_empty else None

    def _write(path: Path) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            # Листы «3805 Несоответствия», … — Joined_1_Checks (не Bill-to и не адресная витрина)
            add_mismatch_sheets(
                w,
                errors_df,
                bill_to_sorg_folders,
                placeholder_message=ph if after_filters_empty else None,
            )

            # Привязка Bill-to по ИНН — по одному листу на каждую SOrg из пары/группы
            add_bill_to_sheets(w, bill_to_df, bill_to_sorg_folders, placeholder_message=ph)

            # Sheet 3: Exception
            if not exc_df.empty:
                exc_df.to_excel(w, sheet_name="Exception", index=False)
            else:
                pd.DataFrame({"Сообщение": ["Нет записей-исключений"]}).to_excel(
                    w, sheet_name="Exception", index=False)

    try:
        _write(out_file)
    except PermissionError:
        ts = datetime.now().strftime("%H%M%S")
        _write(pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str} - Необходимый итоговый файл_{ts}.xlsx")

    if not errors_df.empty:
        _prep_mismatch_sheet(errors_df).to_excel(
            pair_dir / f"{pair_name}_ErrorsOnly.xlsx", index=False)
    with pd.ExcelWriter(pair_dir / f"{pair_name}_BillToByINN.xlsx", engine="openpyxl") as w:
        add_bill_to_sheets(w, bill_to_df, bill_to_sorg_folders, placeholder_message=ph)


def process_pair(pair_name: str, folders: list[str], exc_df: pd.DataFrame,
                 *, progress_lo: int = 5, progress_hi: int = 95,
                 job_index: int = 1, job_total: int = 1) -> bool:
    def prog(step: float, msg: str) -> None:
        span = max(1, progress_hi - progress_lo)
        pct = int(progress_lo + span * max(0.0, min(1.0, step)))
        report_build_progress(pct, f"[{job_index}/{job_total}] {msg}",
                              job_index=job_index, job_total=job_total)

    prog(0.02, f"«{pair_name}»: чтение файлов…")
    base, bp_df, py_df, zy_df = load_data(folders)
    if base.empty:
        print(
            f"[build_checks] «{pair_name}»: нет *Base*.xlsx в папках {folders} — сохраняю пустой итог пары",
            flush=True,
        )
        save_excel(pair_name, pd.DataFrame(), pd.DataFrame(), exc_df, bill_to_sorg_folders=folders)
        return True

    print(f"[build_checks] «{pair_name}»: base={len(base)} строк, "
          f"bp={len(bp_df) if bp_df is not None else 0}, "
          f"py={len(py_df) if py_df is not None else 0}, "
          f"zy={len(zy_df) if zy_df is not None else 0}", flush=True)

    prog(0.20, f"«{pair_name}»: merge BP/PY/ZY…")
    merged = merge_all_partners(base, bp_df, py_df, zy_df)

    prog(0.38, f"«{pair_name}»: фильтры…")
    exc_keys_set = exception_keys(exc_df)
    filtered = apply_filters(merged, exc_keys_set)
    print(f"[build_checks] «{pair_name}»: после фильтров {len(filtered)} строк", flush=True)

    if filtered.empty:
        print(
            f"[build_checks] «{pair_name}»: после фильтров 0 строк — отчёт с пояснением (нормальная ситуация).",
            flush=True,
        )
        save_excel(
            pair_name,
            pd.DataFrame(),
            pd.DataFrame(),
            exc_df,
            bill_to_sorg_folders=folders,
            after_filters_empty=True,
        )
        return True

    prog(0.55, f"«{pair_name}»: проверки MD…")
    checked = add_checks(filtered)

    errors_only = checked[checked["Comment MD Analyst"] != ""].copy()
    if not errors_only.empty:
        errors_only = errors_only.drop_duplicates(subset=["Customer"], keep="first")
        errors_only = errors_only[errors_only["BP Name"].fillna("").astype(str).str.strip().ne("")]
    errors_only = _attach_comment_om(errors_only, exc_df)
    print(f"[build_checks] «{pair_name}»: несоответствий {len(errors_only)}", flush=True)

    prog(0.72, f"«{pair_name}»: Bill-to по ИНН…")
    bill_to = build_bill_to(checked, base)
    print(f"[build_checks] «{pair_name}»: Bill-to строк {len(bill_to)}", flush=True)

    prog(0.88, f"«{pair_name}»: сохранение Excel…")
    if os.environ.get("REPORTS_DEBUG"):
        print(
            f"[build_checks] DEBUG checked.columns ({len(checked.columns)}): "
            f"{checked.columns.tolist()}",
            flush=True,
        )
        print(
            f"[build_checks] DEBUG bill_to.columns ({len(bill_to.columns)}): "
            f"{bill_to.columns.tolist()}",
            flush=True,
        )
        print(
            f"[build_checks] DEBUG errors_only {errors_only.shape}, "
            f"MISMATCH шапка: {MISMATCH_EXPORT_COLS[:6]}…",
            flush=True,
        )
    save_excel(pair_name, errors_only, bill_to, exc_df, bill_to_sorg_folders=folders)
    prog(1.0, f"«{pair_name}»: готово")
    return True


def _all_so_folders() -> list[str]:
    known = ["3801", "3802", "3803", "3804", "3805", "3806"]
    return [f for f in known if (BASE_DIR / f).is_dir()] or known


def _parse_folders(raw: str) -> list[str]:
    if not raw:
        return []
    seen: set[str] = set()
    out = []
    for token in raw.split(","):
        for f in re.findall(r"\d{4}", token.strip()):
            if f not in seen:
                out.append(f)
                seen.add(f)
    return out


def build_jobs(mode: str, folders_csv: str = "") -> list[tuple[str, list[str]]]:
    if mode == "pairs":
        jobs = [(name, cfg["folders"]) for name, cfg in PAIRS.items()]
        raw = folders_csv.strip()
        if raw:
            sel = set(_parse_folders(folders_csv))
            if sel:
                filtered = [(n, folders) for n, folders in jobs if sel.intersection(set(folders))]
                if filtered:
                    jobs = filtered
        return jobs
    all_f = _all_so_folders()
    if mode == "single":
        return [(f, [f]) for f in all_f]
    sel = _parse_folders(folders_csv) or all_f
    if mode == "custom_single":
        return [(f, [f]) for f in sel]
    if mode == "custom_group":
        name = "_".join(sel)
        return [(f"custom_{name}", sel)]
    raise ValueError(f"Неизвестный режим: {mode}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Анализ PF BP-PY-ZY")
    parser.add_argument("--mode", choices=["pairs", "single", "custom_single", "custom_group"], default="pairs")
    parser.add_argument("--folders", default="")
    parser.add_argument("--skip-manual-exceptions", action="store_true")  # ignored, kept for web UI compat
    args = parser.parse_args()

    if not BASE_DIR.exists():
        print(f"[build_checks] Нет каталога: {BASE_DIR}", flush=True)
        return 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_build_progress(2, "Загрузка исключений…", job_index=0, job_total=1)
    exc_df = collect_and_persist_global_exception(BASE_DIR, OUTPUT_DIR)

    try:
        jobs = build_jobs(args.mode, args.folders)
    except ValueError as e:
        print(f"[build_checks] {e}", flush=True)
        return 1

    if not jobs:
        return 0

    nj = len(jobs)
    report_build_progress(5, f"Задач: {nj}", job_index=0, job_total=nj)
    produced = False
    try:
        for i, (job_name, job_folders) in enumerate(jobs):
            lo = 6 + int(88 * i / nj)
            hi = 6 + int(88 * (i + 1) / nj)
            report_build_progress(lo, f"Задача «{job_name}» ({i+1}/{nj})…",
                                  job_index=i+1, job_total=nj)
            if process_pair(job_name, job_folders, exc_df,
                            progress_lo=lo, progress_hi=hi,
                            job_index=i+1, job_total=nj):
                produced = True
    except Exception:
        traceback.print_exc()
        return 1

    if not produced:
        report_build_progress(95, "Ни один отчёт не сохранён", job_index=nj, job_total=max(1, nj))
        print("[build_checks] Ни один отчёт не сохранён — нет файлов Base.", flush=True)

    report_build_progress(100, "Готово", job_index=nj, job_total=max(1, nj))
    return 0


if __name__ == "__main__":
    sys.exit(main())

