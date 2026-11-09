# -*- coding: utf-8 -*-
"""周数据采集：磁力金牛花费/成交 + 生意通退款率，产出预览

预览只读不写表；用户确认后再调用 merge_roi.run_merge(items=...) 落盘。
"""
from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

import export_full
import export_syt
import kx_core as kx
import merge_excel
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


def _syt_file(day: date) -> Path | None:
    d = kx.DATA_DIR / "syt"
    if not d.exists():
        return None
    cands = sorted(d.glob(f"商品列表{day:%Y%m%d}_*.xlsx"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


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


def read_month_cost(path: Path | None, pids: list[str], start: str, end: str) -> dict:
    """从月度表「消耗成交」读指定商品区间消耗，返回 value/missing/days"""
    if not path or not Path(path).exists():
        return {pid: {"value": 0.0, "missing": [], "days": 0} for pid in pids}
    wb = load_workbook(path, data_only=True)
    try:
        sheet = merge_excel.SHEET if merge_excel.SHEET in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet]
        _total_start, cost_cols, total_row = merge_excel._layout(ws)
        pid_col = {}
        for col in cost_cols:
            pid = merge_excel._cell_id(ws.cell(row=2, column=col).value)
            if pid:
                pid_col[pid] = col
        d0, d1 = _parse(start), _parse(end)
        days = list(range(d0.day, d1.day + 1))
        out = {}
        for pid in pids:
            col = pid_col.get(pid)
            if not col:
                out[pid] = {"value": 0.0, "missing": [
                    f"{d0.year}-{d0.month:02d}-{day:02d}" for day in days], "days": len(days)}
                continue
            s = 0.0
            missing = []
            for day in days:
                r = day + merge_excel.FIRST_DATA_ROW - 1
                v = ws.cell(row=r, column=col).value if r <= total_row - 1 else None
                if isinstance(v, (int, float)):
                    s += float(v)
                else:
                    missing.append(f"{d0.year}-{d0.month:02d}-{day:02d}")
            out[pid] = {"value": round(s, 2), "missing": missing, "days": len(days)}
        return out
    finally:
        wb.close()


def check_cache(start: str, end: str, range_start: str, range_end: str) -> dict:
    """检查四项数据是否已存在，供前端提示用户是否复用"""
    d_start = _parse(start)
    d_rs = _parse(range_start)
    prev_end = (d_start - timedelta(days=1)).isoformat()
    out: dict = {}
    for key, s, e in (("week_csv", start, end), ("prev_csv", range_start, prev_end)):
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
    prev_end = (d_start - timedelta(days=1)).isoformat()

    # 1) 磁力金牛 7 日段（逐日，区间 groupType=1 实测数据重复/遗漏，不可用）
    week_csv = _export_ok(start, end) if reuse else None
    if week_csv:
        logger(f"[week] 复用已有导出: {week_csv}")
    else:
        logger(f"[week] 取磁力金牛 {start}~{end}")
        export_full.run_export(cfg, start, end, None, logger=logger)
        week_csv = _export_ok(start, end)
    if not week_csv:
        raise RuntimeError("磁力金牛 7 日数据导出失败")

    # 2) 磁力金牛 30 日历史段
    prev_csv = _export_ok(range_start, prev_end) if reuse else None
    if prev_csv:
        logger(f"[week] 复用已有导出: {prev_csv}")
    else:
        logger(f"[week] 取磁力金牛 {range_start}~{prev_end}")
        export_full.run_export(cfg, range_start, prev_end, None, logger=logger)
        prev_csv = _export_ok(range_start, prev_end)
    if not prev_csv:
        raise RuntimeError("磁力金牛 30 日历史数据导出失败")

    # 3) 生意通
    syt_week = _syt_file(d_start) if reuse else None
    syt_range = _syt_file(d_rs) if reuse else None
    if syt_week and syt_range:
        logger(f"[week] 复用生意通文件: {syt_week.name} / {syt_range.name}")
    else:
        logger("[week] 下载生意通数据（上一周 + 近30日）")
        res = export_syt.run_export(cfg, logger=logger)
        if res.get("errors"):
            raise RuntimeError(f"生意通下载失败: {res['errors'][0].get('error')}")
        syt_week = _syt_file(d_start)
        syt_range = _syt_file(d_rs)
    if not syt_week or not syt_range:
        raise RuntimeError("生意通文件缺失")

    # 4) 汇总
    products = _read_items(cfg)
    week = merge_roi.read_products(week_csv)
    prev = merge_roi.read_products(prev_csv)

    targets = {pid for pid, _n in products}
    week_rate = merge_roi.read_syt_rates(syt_week, targets)
    range_rate = merge_roi.read_syt_rates(syt_range, targets)
    missing = targets - set(week_rate) | targets - set(range_rate)
    if missing:
        raise RuntimeError(f"生意通缺少商品退款率: {sorted(missing)}")

    source = str(cfg.get("week_cost_source") or "both").strip().lower()
    if source not in ("month", "niu", "both"):
        source = "both"
    items = []
    month_path = kx.resolve_path(cfg["excel_path"]) if cfg.get("excel_path") else None
    pids = [p for p, _n in products]
    m7 = read_month_cost(month_path, pids, start, end)
    m30 = read_month_cost(month_path, pids, range_start, range_end)
    for pid, name in products:
        w = week.get(pid, [0.0, 0.0])
        p = prev.get(pid, [0.0, 0.0])
        niu7 = round(w[0], 2)
        niu30 = round(w[0] + p[0], 2)
        d7 = m7.get(pid) or {"value": 0.0, "missing": [], "days": 0}
        d30 = m30.get(pid) or {"value": 0.0, "missing": [], "days": 0}
        items.append({
            "product_id": pid,
            "product_name": name,
            "cost7": d7["value"] if source == "month" else niu7,
            "cost7_niu": niu7,
            "cost7_month": d7["value"],
            "cost30_niu": niu30,
            "cost30_month": d30["value"],
            "missing7": d7["missing"],
            "missing30": d30["missing"],
            "incomplete7": bool(d7["missing"]),
            "incomplete30": bool(d30["missing"]),
            "gmv7": round(w[1], 2),
            "rate7": week_rate[pid],
            "cost30": d30["value"] if source == "month" else niu30,
            "gmv30": round(w[1] + p[1], 2),
            "rate30": range_rate[pid],
        })

    data = {
        "start": start, "end": end,
        "range_start": range_start, "range_end": range_end,
        "cost_source": source,
        "label7": merge_roi.date_label(d_start, d_end),
        "label30": merge_roi.date_label(d_rs, d_re),
        "items": items,
        "totals": {
            "cost7": round(sum(x["cost7"] for x in items), 2),
            "gmv7": round(sum(x["gmv7"] for x in items), 2),
            "cost30": round(sum(x["cost30"] for x in items), 2),
            "gmv30": round(sum(x["gmv30"] for x in items), 2),
        },
        "files": {
            "week_csv": str(week_csv), "prev_csv": str(prev_csv),
            "syt_week": str(syt_week), "syt_range": str(syt_range),
        },
        "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_preview(data)
    logger(f"[week] 预览就绪：{len(items)} 个商品，7日花费 {data['totals']['cost7']}，"
           f"30日花费 {data['totals']['cost30']}")
    return data


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
