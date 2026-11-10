# -*- coding: utf-8 -*-
"""周数据采集：磁力金牛花费/成交 + 生意通退款率，产出预览

预览只读不写表；用户确认后再调用 merge_roi.run_merge(items=...) 落盘。
"""
from __future__ import annotations

import json
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

import export_full
import export_syt
import kx_core as kx
import merge_roi

PREVIEW_DIR = kx.DATA_DIR / "week_preview"


def _log_default(msg: str) -> None:
    print(msg, flush=True)


def _parse(d: str) -> date:
    return datetime.strptime(d, "%Y-%m-%d").date()


def _export_ok(start: str, end: str) -> Path | None:
    """已有导出且 summary 的日期区间匹配才算可复用"""
    d = kx.DATA_DIR / "exports" / f"{start}_{end}"
    csv = d / "products.csv"
    if not csv.exists():
        return None
    summary = d / "summary.json"
    if summary.exists():
        try:
            j = json.loads(summary.read_text(encoding="utf-8"))
            if j.get("start") != start or j.get("end") != end:
                return None
        except Exception:  # noqa: BLE001
            return None
    return csv


SYT_NAME_RE = re.compile(r"商品列表(\d{8})_(\d{8})_")


def _syt_file(day: date, tolerance_days: int = 3) -> Path | None:
    """找最接近指定起始日的生意通文件（平台窗口可能与程序口径差 1 天）"""
    d = kx.DATA_DIR / "syt"
    if not d.exists():
        return None
    best = None
    best_gap = tolerance_days + 1
    for p in d.glob("商品列表*_*.xlsx"):
        m = SYT_NAME_RE.match(p.name)
        if not m:
            continue
        try:
            start = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        gap = abs((start - day).days)
        if gap < best_gap or (gap == best_gap and best is not None
                              and p.stat().st_mtime > best.stat().st_mtime):
            best, best_gap = p, gap
    return best


def _syt_range(p: Path) -> tuple[date, date] | None:
    m = SYT_NAME_RE.match(p.name)
    if not m:
        return None
    try:
        return (datetime.strptime(m.group(1), "%Y%m%d").date(),
                datetime.strptime(m.group(2), "%Y%m%d").date())
    except ValueError:
        return None


def _read_items(cfg: dict) -> list[tuple[str, str]]:
    """从核算表最后一块读商品顺序"""
    raw = str(cfg.get("roi_book_path") or "").strip()
    if not raw:
        raise RuntimeError("未配置核算表路径 roi_book_path")
    book = kx.resolve_path(raw)
    if not book.exists():
        raise FileNotFoundError(f"核算表不存在: {book}")
    wb = load_workbook(book)
    try:
        ws = wb[merge_roi.SHEET_PROFIT]
        blocks = merge_roi.detect_periods(ws)
        if not blocks:
            raise RuntimeError("盈利表未识别到周期块")
        last = blocks[-1]
        out = []
        for i in range(merge_roi.ITEMS_PER_PERIOD):
            r = last["first"] + i
            out.append((merge_roi._cell_id(ws.cell(row=r, column=2).value),
                        str(ws.cell(row=r, column=3).value or "").strip()))
        return out
    finally:
        wb.close()


def _export_summary(start: str, end: str) -> dict:
    p = kx.DATA_DIR / "exports" / f"{start}_{end}" / "summary.json"
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}


def _export_failures(start: str, end: str, range_key: str) -> list[dict]:
    out = []
    for f in _export_summary(start, end).get("failures") or []:
        out.append({
            "range": range_key,
            "account_id": str(f.get("account_id") or ""),
            "account_name": str(f.get("account_name") or ""),
            "error": str(f.get("error") or "")[:300],
        })
    return out


def check_cache(start: str, end: str, range_start: str, range_end: str) -> dict:
    """检查四项数据是否已存在，供前端提示用户是否复用"""
    d_start = _parse(start)
    d_rs = _parse(range_start)
    out: dict = {}
    for key, s, e in (("week_csv", start, end), ("range_csv", range_start, range_end)):
        p = kx.DATA_DIR / "exports" / f"{s}_{e}" / "products.csv"
        out[key] = ({"exists": True, "path": str(p),
                     "mtime": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")}
                    if p.exists() else {"exists": False})
    for key, day in (("syt_week", d_start), ("syt_range", d_rs)):
        p = _syt_file(day)
        out[key] = ({"exists": True, "path": str(p), "name": p.name,
                     "mtime": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")}
                    if p else {"exists": False})
    out["any"] = any(v.get("exists") for v in out.values() if isinstance(v, dict))
    return out


def collect(cfg: dict, start: str, end: str, range_start: str, range_end: str,
            reuse: bool = True, logger=_log_default) -> dict:
    d_start, d_end = _parse(start), _parse(end)
    d_rs, d_re = _parse(range_start), _parse(range_end)

    # 1) 磁力金牛 7 日段（区间接口 groupType=1，按天返回，与逐日结果一致）
    week_csv = _export_ok(start, end) if reuse else None
    if week_csv:
        logger(f"[week] 复用已有导出: {week_csv}")
    else:
        logger(f"[week] 取磁力金牛 {start}~{end}（区间）")
        export_full.run_export(cfg, start, end, None, logger=logger, range_mode=True)
        week_csv = _export_ok(start, end)
    if not week_csv:
        raise RuntimeError("磁力金牛 7 日数据导出失败")

    # 2) 磁力金牛 30 日段（整段一次区间查询，不再由 7 日 + 历史段拼接）
    range_csv = _export_ok(range_start, range_end) if reuse else None
    if range_csv:
        logger(f"[week] 复用已有导出: {range_csv}")
    else:
        logger(f"[week] 取磁力金牛 {range_start}~{range_end}（区间）")
        export_full.run_export(cfg, range_start, range_end, None, logger=logger, range_mode=True)
        range_csv = _export_ok(range_start, range_end)
    if not range_csv:
        raise RuntimeError("磁力金牛 30 日数据导出失败")

    # 3) 生意通
    syt_week = _syt_file(d_start) if reuse else None
    syt_range = _syt_file(d_rs) if reuse else None
    if syt_week and syt_range:
        logger(f"[week] 复用生意通文件: {syt_week.name} / {syt_range.name}")
    else:
        logger("[week] 下载生意通数据（上一周 + 近30日）")
        res = export_syt.run_export(cfg, logger=logger,
                                    week_start=d_start, week_end=d_end,
                                    range_start=d_rs, range_end=d_re)
        if res.get("errors"):
            raise RuntimeError(f"生意通下载失败: {res['errors'][0].get('error')}")
        out_dir = Path(res.get("out_dir") or (kx.DATA_DIR / "syt"))
        by_mode = {f.get("mode"): out_dir / str(f.get("file"))
                   for f in res.get("files") or [] if f.get("file")}
        syt_week = by_mode.get("week") or _syt_file(d_start)
        syt_range = by_mode.get("range") or _syt_file(d_rs)
    if not syt_week or not syt_range:
        have = sorted(p.name for p in (kx.DATA_DIR / "syt").glob("商品列表*.xlsx"))[-3:]
        raise RuntimeError(f"生意通文件缺失（周={d_start}，30日={d_rs}；"
                           f"现有文件：{have or '无'}）")

    # 生意通窗口可能与程序口径差几天（如「近30日」截至昨天），日志提示
    for tag, p, expect in (("上一周", syt_week, (d_start, d_end)),
                           ("近30日", syt_range, (d_rs, d_re))):
        rng = _syt_range(p)
        if rng and rng != expect:
            logger(f"[week] 注意：生意通{tag}窗口 {rng[0]}~{rng[1]} 与金牛 "
                   f"{expect[0]}~{expect[1]} 不一致，退款率按生意通口径")

    # 4) 汇总
    products = _read_items(cfg)
    week = merge_roi.read_products(week_csv)
    range_rows = merge_roi.read_products(range_csv)

    targets = {pid for pid, _n in products}
    week_rate = merge_roi.read_syt_rates(syt_week, targets)
    range_rate = merge_roi.read_syt_rates(syt_range, targets)
    missing = targets - set(week_rate) | targets - set(range_rate)
    if missing:
        raise RuntimeError(f"生意通缺少商品退款率: {sorted(missing)}")

    items = []
    for pid, name in products:
        w = week.get(pid, [0.0, 0.0])
        r30 = range_rows.get(pid, [0.0, 0.0])
        items.append({
            "product_id": pid,
            "product_name": name,
            "cost7": round(w[0], 2),
            "cost30": round(r30[0], 2),
            "gmv7": round(w[1], 2),
            "gmv30": round(r30[1], 2),
            "rate7": week_rate[pid],
            "rate30": range_rate[pid],
        })

    failures = _export_failures(start, end, "7") + _export_failures(range_start, range_end, "30")

    data = {
        "start": start, "end": end,
        "range_start": range_start, "range_end": range_end,
        "label7": merge_roi.date_label(d_start, d_end),
        "label30": merge_roi.date_label(d_rs, d_re),
        "items": items,
        "failures": failures,
        "totals": {
            "cost7": round(sum(x["cost7"] for x in items), 2),
            "gmv7": round(sum(x["gmv7"] for x in items), 2),
            "cost30": round(sum(x["cost30"] for x in items), 2),
            "gmv30": round(sum(x["gmv30"] for x in items), 2),
        },
        "files": {
            "week_csv": str(week_csv), "range_csv": str(range_csv),
            "syt_week": str(syt_week), "syt_range": str(syt_range),
        },
        "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_preview(data)
    logger(f"[week] 预览就绪：{len(items)} 个商品，7日消耗 {data['totals']['cost7']}，"
           f"30日消耗 {data['totals']['cost30']}")
    if failures:
        logger(f"[week] 有 {len(failures)} 个账户区间数据失败，合计可能偏小，可在预览里重试")
    return data


def retry(cfg: dict, start: str, end: str, range_start: str, range_end: str,
          range_key: str, account_ids: list[str], logger=_log_default) -> dict:
    """手动重试：只重抓指定区间里失败的账户，并合并回已有导出"""
    seg = {"7": (start, end), "30": (range_start, range_end)}
    if range_key not in seg:
        raise ValueError("range 只能是 7 或 30")
    s, e = seg[range_key]
    return export_full.retry_accounts(cfg, s, e, account_ids, logger=logger)


def preview_path(start: str, end: str) -> Path:
    return PREVIEW_DIR / f"{start}_{end}.json"


def save_preview(data: dict) -> Path:
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    p = preview_path(data["start"], data["end"])
    p.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return p


def load_preview(start: str, end: str) -> dict | None:
    p = preview_path(start, end)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return None
