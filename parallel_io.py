# -*- coding: utf-8 -*-
"""Параллельное чтение Excel и asyncio-обёртки для тяжёлых sync-функций."""
from __future__ import annotations

import asyncio
import os
from concurrent.futures import ThreadPoolExecutor
from typing import Callable, TypeVar

T = TypeVar("T")

_EXECUTOR: ThreadPoolExecutor | None = None


def parallel_enabled() -> bool:
    return os.environ.get("REPORTS_PARALLEL", "1").strip().lower() not in ("0", "false", "no")


def worker_count(tasks: int, *, default_cap: int = 4) -> int:
    raw = os.environ.get("REPORTS_WORKERS", "").strip()
    if raw.isdigit():
        return max(1, min(int(raw), max(1, tasks)))
    return max(1, min(tasks, default_cap))


def _executor() -> ThreadPoolExecutor:
    global _EXECUTOR
    if _EXECUTOR is None:
        cap = worker_count(8, default_cap=8)
        _EXECUTOR = ThreadPoolExecutor(max_workers=cap, thread_name_prefix="pf_checks")
    return _EXECUTOR


def shutdown_executor() -> None:
    global _EXECUTOR
    if _EXECUTOR is not None:
        _EXECUTOR.shutdown(wait=True, cancel_futures=False)
        _EXECUTOR = None


def run_io(fn: Callable[..., T], /, *args, **kwargs) -> T:
    return fn(*args, **kwargs)


def map_io(func: Callable[[T], T], items: list[T], *, max_workers: int | None = None) -> list:
    if not items:
        return []
    if not parallel_enabled() or len(items) == 1:
        return [func(x) for x in items]
    workers = max_workers or worker_count(len(items))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="pf_map") as pool:
        return list(pool.map(func, items))


async def async_io(fn: Callable[..., T], /, *args, **kwargs) -> T:
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_executor(), lambda: fn(*args, **kwargs))


async def gather_limited(
    coros: list,
    *,
    limit: int,
) -> list:
    if not coros:
        return []
    sem = asyncio.Semaphore(max(1, limit))

    async def _wrap(coro):
        async with sem:
            return await coro

    return await asyncio.gather(*[_wrap(c) for c in coros])
