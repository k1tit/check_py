# -*- coding: utf-8 -*-
"""
PF BP-PY-ZY — логика 1:1 с Access (q_*_Joined_1_Checks, q_*_Joined_3_BillToByINN).
Итоговый Excel — парами (3801_3803, 3802_3804, 3805_3806).
"""
from __future__ import annotations

import argparse
import asyncio
import os
import sys
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None  # type: ignore

from parallel_io import async_io, gather_limited, parallel_enabled, shutdown_executor, worker_count

from build_checks import (
    BASE_DIR,
    BILL_TO_COLS,
    BLOCKED_ORBLK,
    FAKE_INN,
    OUTPUT_DIR,
    PAIRS,
    _CYR,
    _col,
    _excel_sheet_name_safe,
    _get_file,
    _mismatch_sheet_label,
    _nc,
    _norm,
    _ns,
    _prep_bill_to_sheet,
    _attach_comment_om,
    _prep_bill_to_sheet,
    _prep_mismatch_sheet,
    _read_base,
    _read_partner,
    _so_col,
    _unique_excel_sheet_name,
    collect_and_persist_global_exception,
    exception_keys,
    load_runtime_paths_dict,
    _norm_cust,
)

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# Как в Access q_*_Joined_1 (без «самлинг» — его нет в SQL)
ACCESS_NAME_BLACKLIST = (
    "sampling",
    "сэмплинг",
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

_BILL_TO_B_SIDE = [
    "BP SO", "BP Customer", "BP Name", "BP Tax Number 1", "BP CGrp",
    "BP Grp4", "BP Search Term 2", "BP SDst", "BP A7", "BP OrBlk1", "BP OrBlk2",
]


def dedupe_base_access(df: pd.DataFrame) -> pd.DataFrame:
    """q_*_Base_DelDup: GROUP BY Customer, First() = первая строка группы (кроме Name — кириллица)."""
    if df.empty:
        return df

    rows = []
    for _, grp in df.groupby("Customer", sort=False, dropna=False):
        # Access First([field]) — значение из первой физической строки группы, не «первое непустое».
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
    out = result[df.columns].reset_index(drop=True)
    return out.drop_duplicates(subset=["Customer"], keep="first").reset_index(drop=True)


def _effective_so(df: pd.DataFrame, folder_so: str) -> pd.Series:
    so = _ns(df["SO"]) if "SO" in df.columns else pd.Series("", index=df.index, dtype=object)
    if "_folder" in df.columns:
        so = so.where(so.ne(""), _ns(df["_folder"]))
    if folder_so:
        so = so.where(so.ne(""), folder_so)
    return so


def apply_exception_access(base_dd: pd.DataFrame, exc_keys: set[tuple[str, str]], folder_so: str) -> pd.DataFrame:
    """Access: LEFT JOIN Exception … WHERE e.Customer IS NULL (до checks, на base)."""
    if base_dd.empty or not exc_keys:
        return base_dd
    so = _effective_so(base_dd, folder_so)
    cu = _nc(base_dd["Customer"])
    keep = pd.Series(
        [(_norm(s), _norm_cust(c)) not in exc_keys for s, c in zip(so, cu)],
        index=base_dd.index,
    )
    return base_dd[keep].copy()


def _master_lookup(base_dd: pd.DataFrame) -> pd.DataFrame:
    """Строки Base DelDup для LEFT JOIN d ON pf.KTONR = d.Customer."""
    ob = _col(base_dd, "OrBlk", "OrBlk 1")
    ob1 = _col(base_dd, "OrBlk1", "OrBlk 1")
    cols = ["Customer", "Name", "Tax Number 1"]
    if ob:
        cols.append(ob)
    if ob1:
        cols.append(ob1)
    present = [c for c in cols if c in base_dd.columns]
    return base_dd[present].drop_duplicates(subset=["Customer"], keep="first").copy()


def attach_partner_access(
    base_dd: pd.DataFrame,
    partner_df: pd.DataFrame | None,
    prefix: str,
) -> pd.DataFrame:
    """
    Access:
      FROM base b LEFT JOIN (
        SELECT pf.KUNNR AS Customer, pf.KTONR AS {prefix}, d.*
        FROM partner pf LEFT JOIN Base_DelDup d ON pf.KTONR = d.Customer
      ) ON b.Customer = sub.Customer
    """
    result = base_dd.copy()
    ob = _col(base_dd, "OrBlk", "OrBlk 1")
    ob1 = _col(base_dd, "OrBlk1", "OrBlk 1")
    pcols = [prefix, f"{prefix} Name", f"{prefix} Tax Number 1", f"{prefix} OrBlk1", f"{prefix} OrBlk2"]

    if partner_df is None or partner_df.empty:
        # Access: LEFT JOIN к подзапросу BP/PY/ZY без строки → NULL (пусто), не Customer.
        for c in pcols:
            result[c] = ""
        return result

    lookup = _master_lookup(base_dd)
    rename_d = {"Customer": "_lk_cust", "Name": "_lk_name", "Tax Number 1": "_lk_tax"}
    if ob:
        rename_d[ob] = "_lk_ob1"
    if ob1:
        rename_d[ob1] = "_lk_ob2"
    d = lookup.rename(columns=rename_d)

    pf = partner_df[["KUNNR", "KTONR"]].drop_duplicates(subset=["KUNNR"], keep="first").copy()
    pf["KUNNR"] = _nc(pf["KUNNR"])
    pf["KTONR"] = _nc(pf["KTONR"])
    pf = pf.rename(columns={"KUNNR": "Customer", "KTONR": prefix})

    sub = pf.merge(d, left_on=prefix, right_on="_lk_cust", how="left")
    rename_sub = {
        "_lk_name": f"{prefix} Name",
        "_lk_tax": f"{prefix} Tax Number 1",
        "_lk_ob1": f"{prefix} OrBlk1",
        "_lk_ob2": f"{prefix} OrBlk2",
    }
    sub = sub.rename(columns={k: v for k, v in rename_sub.items() if k in sub.columns})
    sub = sub.drop(columns=["_lk_cust"], errors="ignore")
    for c in pcols:
        if c not in sub.columns:
            sub[c] = ""

    sub = sub[["Customer"] + pcols].drop_duplicates(subset=["Customer"], keep="first")

    result = result.drop(columns=[c for c in pcols if c in result.columns], errors="ignore")
    result = result.merge(sub, on="Customer", how="left", validate="one_to_one")

    for c in pcols:
        if c not in result.columns:
            result[c] = ""
        result[c] = (
            result[c].fillna("")
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": "", "<NA>": ""})
        )

    return result


def merge_all_partners_access(
    base_dd: pd.DataFrame,
    bp_df: pd.DataFrame | None,
    py_df: pd.DataFrame | None,
    zy_df: pd.DataFrame | None,
) -> pd.DataFrame:
    """Цепочка LEFT JOIN BP → PY → ZY; ровно одна строка на Customer."""
    n0 = len(base_dd)
    result = attach_partner_access(base_dd, bp_df, "BP")
    result = attach_partner_access(result, py_df, "PY")
    result = attach_partner_access(result, zy_df, "ZY")
    if len(result) != n0:
        raise RuntimeError(f"merge partners: строк {n0} → {len(result)} (many-to-many)")
    if result["Customer"].duplicated().any():
        dup = result.groupby("Customer").size().sort_values(ascending=False)
        raise RuntimeError(f"merge partners: дубли Customer, max={dup.iloc[0]}")
    return result


def apply_filters_access(df: pd.DataFrame) -> pd.DataFrame:
    """Фильтры Access (Exception уже применён на base)."""
    if df.empty:
        return df

    name_s = df["Name"].fillna("").astype(str).str.lower() if "Name" in df.columns else pd.Series("", index=df.index)
    mask_name = ~name_s.apply(lambda n: any(b in n for b in ACCESS_NAME_BLACKLIST))

    cgrp_s = df["CGrp"].fillna("").astype(str).str.upper().str.strip() if "CGrp" in df.columns else pd.Series("", index=df.index)
    mask_cgrp = cgrp_s != "Z"

    def _ob(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series("", index=df.index)
        t = df[col].fillna("").astype(str).str.strip().str.upper()
        return t.replace({"NAN": "", "<NA>": "", "NONE": ""})

    mask_orblk = ~_ob("OrBlk").isin(BLOCKED_ORBLK) & ~_ob("OrBlk1").isin(BLOCKED_ORBLK)

    tax_s = df["Tax Number 1"].fillna("").astype(str).str.strip() if "Tax Number 1" in df.columns else pd.Series("", index=df.index)
    mask_inn = ~tax_s.apply(lambda t: any(f in t for f in FAKE_INN))

    return df[mask_name & mask_cgrp & mask_orblk & mask_inn].copy()


def load_sorg(folder: str) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    fp = BASE_DIR / folder
    f_base = _get_file(fp, "*Base*.xlsx")
    if not f_base:
        return pd.DataFrame(), None, None, None

    f_bp = _get_file(fp, "*BP*.xlsx")
    f_py = _get_file(fp, "*PY*.xlsx")
    f_zy = _get_file(fp, "*ZY*.xlsx")

    def _file_mb(path: Path | None) -> str:
        if not path or not path.is_file():
            return "?"
        return f"{path.stat().st_size / (1024 * 1024):.1f} MB"

    def _load_base() -> pd.DataFrame:
        print(f"[new_access] SO {folder}: читаю Base {f_base.name} ({_file_mb(f_base)})…", flush=True)
        t0 = time.perf_counter()
        df = dedupe_base_access(_read_base(f_base, folder))
        print(f"[new_access] SO {folder}: Base готов, {len(df)} строк, {time.perf_counter() - t0:.0f} с", flush=True)
        return df

    def _load_partner(path: Path | None, label: str) -> pd.DataFrame | None:
        if not path:
            print(f"[new_access] SO {folder}: {label} не найден", flush=True)
            return None
        print(f"[new_access] SO {folder}: читаю {label} {path.name} ({_file_mb(path)})…", flush=True)
        t0 = time.perf_counter()
        df = _read_partner(path, folder, kind=label)
        rows = len(df) if df is not None and not df.empty else 0
        print(f"[new_access] SO {folder}: {label} готов, {rows} строк, {time.perf_counter() - t0:.0f} с", flush=True)
        return df

    if parallel_enabled():
        jobs: dict[str, object] = {"base": _load_base}
        if f_bp:
            jobs["BP"] = lambda: _load_partner(f_bp, "BP")
        if f_py:
            jobs["PY"] = lambda: _load_partner(f_py, "PY")
        if f_zy:
            jobs["ZY"] = lambda: _load_partner(f_zy, "ZY")
        with ThreadPoolExecutor(max_workers=min(4, len(jobs)), thread_name_prefix=f"so{folder}") as pool:
            futs = {k: pool.submit(fn) for k, fn in jobs.items()}
            base = futs["base"].result()
            bp = futs["BP"].result() if "BP" in futs else None
            py = futs["PY"].result() if "PY" in futs else None
            zy = futs["ZY"].result() if "ZY" in futs else None
        return base, bp, py, zy

    base = _load_base()
    return base, _load_partner(f_bp, "BP"), _load_partner(f_py, "PY"), _load_partner(f_zy, "ZY")


def add_checks_access(df: pd.DataFrame) -> pd.DataFrame:
    """q_*_Joined_1_Checks — тексты комментариев как в Access (SP, не Cust)."""
    out = df.copy()

    def _s(col: str) -> pd.Series:
        return out[col].fillna("").astype(str).str.strip() if col in out.columns else pd.Series("", index=out.index)

    if "OrBlk" in out.columns and "OrBlk1" not in out.columns:
        out = out.rename(columns={"OrBlk": "OrBlk1"})
    elif "OrBlk" in out.columns and "OrBlk1" in out.columns:
        out = out.rename(columns={"OrBlk1": "OrBlk2", "OrBlk": "OrBlk1"})

    orblk1 = _s("OrBlk1").str.upper()
    bp_orblk1 = _s("BP OrBlk1").str.upper()
    bp, py, zy = _s("BP"), _s("PY"), _s("ZY")
    cust, tax, bp_tax, bp_name = _s("Customer"), _s("Tax Number 1"), _s("BP Tax Number 1"), _s("BP Name")

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
        "Bill-to прикреплён к другому Bill-to",
    )
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk1 != "M") & (bp_orblk1 != "M") & cust_ne_bp,
        "Ship-to прикреплён к BP без OB M",
    )
    out["Comment MD Analyst"] = comment
    return out


def build_bill_to_access(df_checks: pd.DataFrame, base_dd: pd.DataFrame) -> pd.DataFrame:
    """
    q_*_Joined_3_BillToByINN:
    INNER JOIN b ON j.[Tax Number 1] = b.[Tax Number 1], b.OrBlk = 'M'.
    Без groupby/drop_duplicates после join — Access размножает строки.
    """
    if df_checks.empty or base_dd.empty:
        return pd.DataFrame()

    pool1 = df_checks[
        (df_checks["Cust OrBlk Bill-to"] == "Ship-to")
        & (df_checks["Check Cust&BP"] == "TRUE")
        & (df_checks["Comment MD Analyst"].fillna("").astype(str).str.strip() == "")
    ].copy()
    if pool1.empty:
        return pd.DataFrame()

    pool1 = pool1.copy()
    if "_folder" in pool1.columns:
        so_series = _ns(pool1["SO"]) if "SO" in pool1.columns else pd.Series("", index=pool1.index, dtype=object)
        need = so_series.eq("") | so_series.isna()
        if need.any():
            if "SO" not in pool1.columns:
                pool1["SO"] = ""
            pool1.loc[need, "SO"] = _ns(pool1.loc[need, "_folder"])

    orblk_col = _col(base_dd, "OrBlk", "OrBlk 1")
    if not orblk_col:
        return pd.DataFrame()

    pool2 = base_dd[base_dd[orblk_col].fillna("").astype(str).str.upper().str.strip() == "M"].copy()
    if pool2.empty:
        return pd.DataFrame()

    rename_p2: dict[str, str] = {"Customer": "BP Customer", orblk_col: "BP OrBlk1"}
    so_c = _so_col(pool2)
    if so_c:
        rename_p2[so_c] = "BP SO"
    for src, dst in {
        "Name": "BP Name",
        "Tax Number 1": "BP Tax Number 1",
        "CGrp": "BP CGrp",
    }.items():
        if src in pool2.columns:
            rename_p2[src] = dst
    g4 = _col(pool2, "Grp4", "GROUP4", "Group4")
    if g4:
        rename_p2[g4] = "BP Grp4"
    ob2 = _col(pool2, "OrBlk1", "OrBlk 1", "OrBlk2", "OrBlk 2")
    if ob2:
        rename_p2[ob2] = "BP OrBlk2"
    pool2 = pool2.rename(columns=rename_p2)

    for col in _BILL_TO_B_SIDE:
        if col not in pool2.columns:
            pool2[col] = ""

    pool1_j = pool1.drop(columns=[c for c in _BILL_TO_B_SIDE if c in pool1.columns], errors="ignore")
    pool1_j["_tax"] = _ns(pool1_j["Tax Number 1"]) if "Tax Number 1" in pool1_j.columns else ""
    pool2["_tax"] = _ns(pool2["BP Tax Number 1"])

    p2_cols = ["_tax"] + [c for c in _BILL_TO_B_SIDE if c in pool2.columns]
    joined = pool1_j.merge(pool2[p2_cols], on="_tax", how="inner")
    joined = joined.drop(columns=["_tax"], errors="ignore")

    if "BP" in joined.columns:
        joined = joined.rename(columns={"BP": "BP now"})

    for col in _BILL_TO_B_SIDE + ["BP SO"]:
        if col not in joined.columns:
            joined[col] = ""

    # BP status — как IIf в Access
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
    joined = joined.sort_values(sort_cols, kind="stable").drop(columns=["_st"], errors="ignore")
    return _prep_bill_to_sheet(joined.reset_index(drop=True))


def _bill_to_rows_access(bt: pd.DataFrame, so_key: str) -> pd.DataFrame:
    """Лист SOrg: только j.[SO] = SOrg (Access не фильтрует BP SO в WHERE)."""
    if bt.empty or not so_key:
        return pd.DataFrame()
    return bt[_effective_so(bt, so_key).eq(so_key)].copy()


def _bill_to_sheet_label(so_token: str) -> str:
    so = _norm(so_token) or "SO"
    return f"{so} Привязка Bill-to по ИНН"


def save_pair_excel(
    pair_name: str,
    errors_df: pd.DataFrame,
    bill_to_df: pd.DataFrame,
    exc_df: pd.DataFrame,
    sorg_folders: list[str],
) -> None:
    pair_dir = OUTPUT_DIR / pair_name
    pair_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%d.%m.%Y")
    out_file = pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str} - Access.xlsx"

    with pd.ExcelWriter(out_file, engine="openpyxl") as w:
        used: set[str] = set()
        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_mismatch_sheet_label(so_key), used)
            used.add(title)
            sub = errors_df[_effective_so(errors_df, so_key).eq(so_key)] if not errors_df.empty else errors_df
            if not sub.empty:
                sub = sub.sort_values(
                    by=[c for c in ("Comment MD Analyst", "Customer") if c in sub.columns],
                    kind="stable",
                )
            prep = _prep_mismatch_sheet(sub)
            if prep.empty:
                pd.DataFrame({"Сообщение": [f"Нет несоответствий для SO {so_key}"]}).to_excel(
                    w, sheet_name=title, index=False
                )
            else:
                prep.to_excel(w, sheet_name=title, index=False)

        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_bill_to_sheet_label(so_key), used)
            used.add(title)
            grp_raw = _bill_to_rows_access(bill_to_df, so_key)
            grp = _prep_bill_to_sheet(grp_raw) if not grp_raw.empty else pd.DataFrame(columns=BILL_TO_COLS)
            if grp.empty:
                pd.DataFrame(columns=BILL_TO_COLS).to_excel(w, sheet_name=title, index=False)
            else:
                grp.to_excel(w, sheet_name=title, index=False)

        if not exc_df.empty:
            exc_df.to_excel(w, sheet_name="Exception", index=False)
        else:
            pd.DataFrame({"Сообщение": ["Нет записей-исключений"]}).to_excel(w, sheet_name="Exception", index=False)

    print(f"[new_access] сохранено: {out_file}", flush=True)


def process_sorg(folder: str, exc_keys: set[tuple[str, str]]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Access-цепочка для одного SOrg:
    Base → DelDup → Exception → BP/PY/ZY → фильтры → Checks → Errors / Bill-to
    """
    base_raw, bp_df, py_df, zy_df = load_sorg(folder)
    if base_raw.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    base = apply_exception_access(base_raw, exc_keys, folder)
    if base.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    merged = merge_all_partners_access(base, bp_df, py_df, zy_df)
    filtered = apply_filters_access(merged)
    if filtered.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    checked = add_checks_access(filtered)
    errors = checked[checked["Comment MD Analyst"].fillna("").astype(str).str.strip() != ""].copy()
    if not errors.empty:
        errors = errors.drop_duplicates(subset=["Customer"], keep="first")
        # Эталон Access (3801): 110 строк; лишние у нас — с пустым BP Name (нет строки в Base_DelDup).
        errors = errors[errors["BP Name"].fillna("").astype(str).str.strip().ne("")].copy()
    bill_to = build_bill_to_access(checked, base_raw)

    if os.environ.get("REPORTS_DEBUG"):
        dup = checked.groupby("Customer").size().sort_values(ascending=False)
        print(
            f"[new_access] DEBUG SO {folder}: base_raw={len(base_raw)} base_exc={len(base)} "
            f"merged={len(merged)} filtered={len(filtered)} errors={len(errors)} "
            f"dup_max={dup.max() if len(dup) else 0}",
            flush=True,
        )

    return errors, bill_to, base_raw


def _merge_sorg_results(
    pair_name: str,
    folders: list[str],
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]],
    exc_df: pd.DataFrame,
    *,
    failed_folders: list[str] | None = None,
) -> bool:
    errors_parts: list[pd.DataFrame] = []
    bill_parts: list[pd.DataFrame] = []
    order = {f: i for i, f in enumerate(folders)}
    done = {f for f, _ in results}
    for folder, (errors, bill_to, base) in sorted(results, key=lambda x: order.get(x[0], 0)):
        if not base.empty:
            print(
                f"[new_access] SO {folder}: ошибок {len(errors)}, Bill-to {len(bill_to)}",
                flush=True,
            )
        if not errors.empty:
            errors_parts.append(errors)
        if not bill_to.empty:
            bill_parts.append(bill_to)

    errors_all = pd.concat(errors_parts, ignore_index=True) if errors_parts else pd.DataFrame()
    if not errors_all.empty:
        errors_all = _attach_comment_om(errors_all, exc_df)
    bill_all = pd.concat(bill_parts, ignore_index=True) if bill_parts else pd.DataFrame()
    save_pair_excel(pair_name, errors_all, bill_all, exc_df, folders)

    missing = failed_folders or [f for f in folders if f not in done]
    if missing:
        print(
            f"[new_access] сохранено (частично): пара {pair_name} — без SO {', '.join(missing)}",
            flush=True,
        )
    return not missing


def process_pair(pair_name: str, folders: list[str], exc_df: pd.DataFrame) -> bool:
    """Последовательная обработка (для отладки и --no-parallel)."""
    exc_keys = exception_keys(exc_df)
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]] = []
    failed: list[str] = []
    iterator = folders
    if tqdm is not None:
        iterator = tqdm(folders, desc=pair_name, unit="SOrg", leave=True)
    for folder in iterator:
        msg = f"SO {folder}: чтение и проверки…"
        if tqdm is not None and hasattr(iterator, "set_postfix_str"):
            iterator.set_postfix_str(msg)
        else:
            print(f"[new_access] {pair_name}: {msg}", flush=True)
        try:
            results.append((folder, process_sorg(folder, exc_keys)))
        except Exception as exc:
            failed.append(folder)
            print(f"[new_access] SO {folder}: ОШИБКА — {exc}", flush=True)
            traceback.print_exc()
    if not results:
        print(f"[new_access] пара {pair_name}: нет успешных SO — отчёт не сохранён", flush=True)
        return False
    return _merge_sorg_results(pair_name, folders, results, exc_df, failed_folders=failed)


async def process_pair_async(pair_name: str, folders: list[str], exc_df: pd.DataFrame) -> bool:
    exc_keys = exception_keys(exc_df)
    workers = worker_count(len(folders), default_cap=2)

    if not parallel_enabled() or workers <= 1 or len(folders) <= 1:
        return process_pair(pair_name, folders, exc_df)

    print(
        f"[new_access] {pair_name}: параллельно {len(folders)} SOrg (workers={workers})",
        flush=True,
    )

    async def _one(folder: str) -> tuple[str, object]:
        print(f"[new_access] {pair_name}: SO {folder}: чтение и проверки…", flush=True)
        try:
            res = await async_io(process_sorg, folder, exc_keys)
            return folder, res
        except Exception as exc:
            print(f"[new_access] SO {folder}: ОШИБКА — {exc}", flush=True)
            traceback.print_exc()
            return folder, exc

    raw = await gather_limited([_one(f) for f in folders], limit=workers)
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]] = []
    failed: list[str] = []
    for folder, item in raw:
        if isinstance(item, Exception):
            failed.append(folder)
        else:
            results.append((folder, item))
    if not results:
        print(f"[new_access] пара {pair_name}: нет успешных SO — отчёт не сохранён", flush=True)
        return False
    return _merge_sorg_results(pair_name, folders, results, exc_df, failed_folders=failed)


async def _run_all_pairs(jobs: list[tuple[str, list[str]]], exc_df: pd.DataFrame) -> int:
    """Обрабатывает пары по очереди; каждая пара сохраняется сразу после своих SOrg."""
    errors = 0
    for pair_name, folders in jobs:
        print(f"[new_access] === пара {pair_name} ===", flush=True)
        try:
            ok = await process_pair_async(pair_name, folders, exc_df)
            if not ok:
                errors += 1
        except Exception as exc:
            errors += 1
            print(f"[new_access] пара {pair_name}: критическая ошибка — {exc}", flush=True)
            traceback.print_exc()
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="PF BP-PY-ZY — логика Access, отчёт парами")
    parser.add_argument("--mode", choices=["pairs"], default="pairs")
    parser.add_argument(
        "--workers",
        type=int,
        default=0,
        help="Параллельных SOrg в паре (0 = авто, 1 = последовательно)",
    )
    parser.add_argument("--no-parallel", action="store_true", help="Отключить параллельное чтение и обработку")
    args = parser.parse_args()

    if args.no_parallel:
        os.environ["REPORTS_PARALLEL"] = "0"
    elif args.workers > 0:
        os.environ["REPORTS_WORKERS"] = str(args.workers)

    paths = load_runtime_paths_dict()
    script_root = Path(__file__).resolve().parent
    local_data = (script_root / "data").resolve()
    data_dir = paths["data_dir"].resolve()
    print(f"[new_access] data_dir:  {paths['data_dir']}", flush=True)
    print(f"[new_access] base_dir:  {paths['base_dir']}", flush=True)
    print(f"[new_access] result:    {paths['output_dir']}", flush=True)
    if data_dir != local_data and script_root not in data_dir.parents:
        print(
            f"[new_access] ВНИМАНИЕ: data_dir не рядом со скриптом ({script_root}).\n"
            f"  Проверьте runtime_paths.json — возможно читаете/пишете не ту папку data.",
            flush=True,
        )

    if not BASE_DIR.exists():
        print(f"[new_access] Нет каталога: {BASE_DIR}", flush=True)
        return 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    exc_df = collect_and_persist_global_exception(BASE_DIR, OUTPUT_DIR)

    jobs = [(name, cfg["folders"]) for name, cfg in PAIRS.items()]
    par = "вкл" if parallel_enabled() else "выкл"
    w = worker_count(2, default_cap=2)
    print(f"[new_access] пары: {', '.join(j[0] for j in jobs)} | parallel={par} workers≈{w}", flush=True)

    exit_code = 0
    try:
        pair_errors = asyncio.run(_run_all_pairs(jobs, exc_df))
        if pair_errors:
            exit_code = 1
            print(f"[new_access] готово с ошибками в {pair_errors} парах", flush=True)
        else:
            print("[new_access] готово", flush=True)
    except Exception:
        traceback.print_exc()
        exit_code = 1
    finally:
        shutdown_executor()
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
