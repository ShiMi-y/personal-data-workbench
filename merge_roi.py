# -*- coding: utf-8 -*-
"""退款率 / 盈利 两张表的周期追加

做的事：
  1. 「盈利」表末尾追加一个新周期（9 个商品 + 总计行）
     7 日块 A:H，30 日块 J:Q，公式沿用上一周期
  2. 「退款率」表在 G 列前插入 4 列（新块占 G:J，旧块整体右移）
     G/I 引用盈利表新周期，H/J 填生意通退款率
  3. 修上一周期总计行的 SUM 范围（原表漏算最后一行）

红线：
  · 不改表头 1-3 行以外的历史块内容
  · 插入列后 `1` 表的 AY/BB 仍指向 H/J，自动切到新周期
  · 写入前备份，写入后逐项自检
"""
from __future__ import annotations

import csv
import re
import shutil
import sys
from collections import defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import kx_core as kx

SHEET_RATE = "退款率"
SHEET_PROFIT = "盈利"
GROUP_WIDTH = 4          # 退款率表每周期 4 列
INSERT_AT = 7            # 插在 G 列之前
ITEMS_PER_PERIOD = 9     # 每个周期 9 个商品


def log_default(msg: str) -> None:
    print(msg, flush=True)


def date_label(d1: date, d2: date) -> str:
    return f"{d1.month}.{d1.day}-{d2.month}.{d2.day}"


def _parse(d: str) -> date:
    return datetime.strptime(d, "%Y-%m-%d").date()


def _cell_id(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _copy_style(dst, src) -> None:
    if not src.has_style:
        return
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)
    dst.alignment = copy(src.alignment)


def detect_periods(ws) -> list[dict]:
    """盈利表：按 A 列连续相同值切周期块"""
    blocks: list[dict] = []
    cur, start = None, None
    for r in range(1, int(ws.max_row or 0) + 1):
        v = ws.cell(row=r, column=1).value
        s = str(v).strip() if v is not None else ""
        if not s:
            continue
        if cur is None:
            cur, start = s, r
        elif s != cur:
            blocks.append({"label": cur, "first": start, "total": r - 1})
            cur, start = s, r
    if cur is not None:
        blocks.append({"label": cur, "first": start, "total": int(ws.max_row or 0)})
    return blocks


def read_products(path: Path) -> dict[str, list[float]]:
    out: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0])
    with path.open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            pid = str(r.get("product_id") or "").strip()
            if not pid:
                continue
            out[pid][0] += float(r.get("cost") or 0)
            out[pid][1] += float(r.get("gmv") or 0)
    return out


def read_syt_rates(path: Path, targets: set[str]) -> dict[str, float]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    out: dict[str, float] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        pid = str(row[0] or "").strip()
        if pid not in targets:
            continue
        s = str(row[3] or "").strip().replace("%", "")
        try:
            out[pid] = round(float(s) / 100, 6)
        except ValueError:
            pass
    wb.close()
    return out


def _shift_columns(ws, last_col: int, insert_at: int, width: int) -> None:
    """把 insert_at..last_col 整体右移 width 列（从右往左搬，避免覆盖）"""
    max_row = int(ws.max_row or 0)
    for c in range(last_col, insert_at - 1, -1):
        dst_c = c + width
        for r in range(1, max_row + 1):
            src = ws.cell(row=r, column=c)
            dst = ws.cell(row=r, column=dst_c)
            dst.value = src.value
            _copy_style(dst, src)
        w = ws.column_dimensions[get_column_letter(c)].width
        if w:
            ws.column_dimensions[get_column_letter(dst_c)].width = w


def _clear_columns(ws, insert_at: int, width: int) -> None:
    max_row = int(ws.max_row or 0)
    for c in range(insert_at, insert_at + width):
        for r in range(1, max_row + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            _copy_style(cell, ws.cell(row=r, column=c + width))
        src_w = ws.column_dimensions[get_column_letter(c + width)].width
        if src_w:
            ws.column_dimensions[get_column_letter(c)].width = src_w


def _prune_backups(bdir: Path, stem: str, suffix: str, limit: int, logger) -> None:
    """备份数量超过上限时删除最旧的（limit < 0 表示不限制）"""
    if limit < 0:
        return
    files = sorted(bdir.glob(f"{stem}_backup_*{suffix}"), key=lambda p: p.name)
    extra = len(files) - limit
    if extra <= 0:
        return
    removed = 0
    for p in files[:extra]:
        try:
            p.unlink()
            removed += 1
        except Exception:  # noqa: BLE001
            pass
    if removed:
        logger(f"[roi] 备份超出上限 {limit} 个，已删除最旧 {removed} 个")


def run_merge(cfg: dict, start: str, end: str, range_start: str, range_end: str,
              book_path: Path | None = None, week_csv: Path | None = None,
              range_csv: Path | None = None, syt_week: Path | None = None,
              syt_range: Path | None = None, items: list[dict] | None = None,
              backup: bool = True, save: bool = True, overwrite: bool = False,
              logger=log_default) -> dict:
    d_start, d_end = _parse(start), _parse(end)
    d_rs, d_re = _parse(range_start), _parse(range_end)

    book = Path(book_path) if book_path else None
    if book is None:
        raw = str(cfg.get("roi_book_path") or "").strip()
        if not raw:
            raise RuntimeError("未配置核算表路径 roi_book_path")
        book = kx.resolve_path(raw)

    week_csv = Path(week_csv) if week_csv else (kx.DATA_DIR / "exports" / f"{start}_{end}" / "products.csv")
    range_csv = Path(range_csv) if range_csv else (
        kx.DATA_DIR / "exports" / f"{range_start}_{range_end}" / "products.csv")
    syt_dir = kx.DATA_DIR / "syt"

    if syt_week is None or syt_range is None:
        cands = sorted(syt_dir.glob("商品列表*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if syt_week is None:
            syt_week = next((p for p in cands if p.name.startswith(
                f"商品列表{d_start:%Y%m%d}_")), None)
        if syt_range is None:
            syt_range = next((p for p in cands if p.name.startswith(
                f"商品列表{d_rs:%Y%m%d}_")), None)

    if not book or not Path(book).exists():
        raise FileNotFoundError(f"核算表不存在: {book}")
    if not items:
        for label, p in (("本周 products.csv", week_csv), ("30日 products.csv", range_csv),
                         ("生意通周文件", syt_week), ("生意通30日文件", syt_range)):
            if p is None or not Path(p).exists():
                raise FileNotFoundError(f"{label} 不存在: {p}")

    wb = load_workbook(book)
    if SHEET_RATE not in wb.sheetnames or SHEET_PROFIT not in wb.sheetnames:
        raise RuntimeError(f"工作表缺失，实际: {wb.sheetnames}")
    ws_rate, ws_profit = wb[SHEET_RATE], wb[SHEET_PROFIT]

    blocks = detect_periods(ws_profit)
    if not blocks:
        raise RuntimeError("盈利表未识别到周期块")
    last = blocks[-1]
    if str(ws_profit.cell(row=last["total"], column=3).value or "").strip() != "总计":
        raise RuntimeError(f"盈利表最后一块 {last['first']}-{last['total']} 末行不是总计行")
    if last["total"] - last["first"] != ITEMS_PER_PERIOD:
        raise RuntimeError(f"最后一块行数异常：{last['first']}-{last['total']}")

    if items:
        products = [(str(it.get("product_id") or ""), str(it.get("product_name") or ""))
                    for it in items]
    else:
        products = []
        for i in range(ITEMS_PER_PERIOD):
            r = last["first"] + i
            products.append((_cell_id(ws_profit.cell(row=r, column=2).value),
                             str(ws_profit.cell(row=r, column=3).value or "").strip()))
    if len(products) != ITEMS_PER_PERIOD or any(not pid for pid, _n in products):
        raise RuntimeError("商品顺序异常，无法确定 9 个商品")

    if items:
        week = {str(it["product_id"]): [float(it["cost7"]), float(it["gmv7"])] for it in items}
        thirty = {str(it["product_id"]): [float(it["cost30"]), float(it["gmv30"])] for it in items}
        week_rate = {str(it["product_id"]): float(it["rate7"]) for it in items}
        range_rate = {str(it["product_id"]): float(it["rate30"]) for it in items}
    else:
        week = read_products(week_csv)
        # 30 日 CSV 覆盖整个区间（含 7 日窗口），直接汇总即可，不能再加周数据
        range_prod = read_products(range_csv)
        thirty = {pid: list(range_prod.get(pid, [0.0, 0.0]))
                  for pid, _n in products}
        targets = {pid for pid, _n in products}
        week_rate = read_syt_rates(syt_week, targets)
        range_rate = read_syt_rates(syt_range, targets)
        missing = targets - set(week_rate) | targets - set(range_rate)
        if missing:
            raise RuntimeError(f"生意通文件缺少商品退款率: {sorted(missing)}")

    label7 = date_label(d_start, d_end)
    label30 = date_label(d_rs, d_re)
    logger(f"[roi] 周期：7日 {label7}，30日 {label30}")

    dup = [b for b in blocks if b["label"] == label7]
    rate_title = str(ws_rate.cell(row=1, column=INSERT_AT).value or "")
    rate_exists = label7 in rate_title
    if (dup or rate_exists) and not overwrite:
        where = (f"第 {dup[-1]['first']}-{dup[-1]['total']} 行" if dup
                 else f"G1={rate_title}")
        raise RuntimeError(f"核算表已存在周期「{label7}」（{where}），是否覆盖？")
    if dup and dup[-1] is not blocks[-1]:
        raise RuntimeError(f"周期「{label7}」不是最后一个周期，无法覆盖，请手工处理")

    if backup:
        bdir = kx.DATA_DIR / "backup"
        bdir.mkdir(parents=True, exist_ok=True)
        bak = bdir / f"{book.stem}_backup_{datetime.now():%Y%m%d-%H%M%S}{datetime.now().microsecond // 1000:03d}{book.suffix}"
        shutil.copy2(book, bak)
        logger(f"[roi] 已备份: {bak.name}")
        try:
            limit = int(cfg.get("backup_limit", 10))
        except (TypeError, ValueError):
            limit = 10
        _prune_backups(bdir, book.stem, book.suffix, limit, logger)

    # ---------------- 盈利表：追加或覆盖周期 ----------------
    if dup:
        new_first = dup[-1]["first"]
        new_total = dup[-1]["total"]
        logger(f"[roi] 覆盖已有周期 {new_first}-{new_total} 行")
    else:
        new_first = last["total"] + 1
        new_total = new_first + ITEMS_PER_PERIOD
    for i, (pid, name) in enumerate(products):
        r = new_first + i
        ref = last["first"] + i
        w = week.get(pid, [0.0, 0.0])
        t = thirty[pid]
        vals = {
            1: label7, 2: pid, 3: name,
            4: round(w[0], 2), 5: round(w[1], 2),
            6: f"='1'!BA{4 + i}", 7: f"=E{r}/D{r}", 8: f"=E{r}/F{r}-D{r}",
            10: label30, 11: pid, 12: name,
            13: round(t[0], 2), 14: round(t[1], 2),
            15: f"='1'!BD{4 + i}", 16: f"=N{r}/M{r}", 17: f"=N{r}/O{r}-M{r}",
        }
        for c in range(1, 18):
            cell = ws_profit.cell(row=r, column=c)
            cell.value = vals.get(c)
            _copy_style(cell, ws_profit.cell(row=ref, column=c))

    tot = {
        1: label7, 3: "总计",
        4: f"=SUM(D{new_first}:D{new_first + ITEMS_PER_PERIOD - 1})",
        5: f"=SUM(E{new_first}:E{new_first + ITEMS_PER_PERIOD - 1})",
        7: f"=E{new_total}/D{new_total}",
        8: f"=SUM(H{new_first}:H{new_first + ITEMS_PER_PERIOD - 1})",
        10: label30, 12: "总计",
        13: f"=SUM(M{new_first}:M{new_first + ITEMS_PER_PERIOD - 1})",
        14: f"=SUM(N{new_first}:N{new_first + ITEMS_PER_PERIOD - 1})",
        16: f"=N{new_total}/M{new_total}",
        17: f"=SUM(Q{new_first}:Q{new_first + ITEMS_PER_PERIOD - 1})",
    }
    for c in range(1, 18):
        cell = ws_profit.cell(row=new_total, column=c)
        cell.value = tot.get(c)
        _copy_style(cell, ws_profit.cell(row=last["total"], column=c))
    logger(f"[roi] 盈利表已{'覆盖' if dup else '追加'} {new_first}-{new_total} 行")

    # ---------------- 修上一周期总计行 ----------------
    fixed = []
    prev_block = blocks[-2] if (dup and len(blocks) >= 2) else last
    sum_re = re.compile(r"^=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$")
    for c in range(1, int(ws_profit.max_column or 0) + 1):
        cell = ws_profit.cell(row=prev_block["total"], column=c)
        m = sum_re.match(str(cell.value or ""))
        if not m:
            continue
        col = m.group(1)
        cell.value = f"=SUM({col}{prev_block['first']}:{col}{prev_block['total'] - 1})"
        fixed.append(f"{col}{prev_block['total']}")
    if fixed:
        logger(f"[roi] 已修上一周期总计行: {', '.join(fixed)}")

    # ---------------- 退款率表：插 4 列 ----------------
    last_col = int(ws_rate.max_column or 0)
    max_row = int(ws_rate.max_row or 0)
    if last_col < INSERT_AT:
        raise RuntimeError(f"退款率表列数异常: {last_col}")
    headers = [f"7日销售额（{label7}）", f"7日退款率（{label7}）",
               f"30日销售额（{label30}）", f"30日退款率（{label30}）"]
    if not rate_exists:
        _shift_columns(ws_rate, last_col, INSERT_AT, GROUP_WIDTH)
        _clear_columns(ws_rate, INSERT_AT, GROUP_WIDTH)
        for i, h in enumerate(headers):
            cell = ws_rate.cell(row=1, column=INSERT_AT + i)
            cell.value = h
            _copy_style(cell, ws_rate.cell(row=1, column=INSERT_AT + i + GROUP_WIDTH))
    else:
        logger("[roi] 退款率表已存在该周期，直接覆盖 G:J")

    for i, (pid, _name) in enumerate(products):
        r = 2 + i
        vals = {
            INSERT_AT: f"=盈利!E{new_first + i}",
            INSERT_AT + 1: week_rate[pid],
            INSERT_AT + 2: f"=盈利!N{new_first + i}",
            INSERT_AT + 3: range_rate[pid],
        }
        for c, v in vals.items():
            cell = ws_rate.cell(row=r, column=c)
            cell.value = v
            _copy_style(cell, ws_rate.cell(row=r, column=c + GROUP_WIDTH))
    logger(f"[roi] 退款率表新块 {get_column_letter(INSERT_AT)}:"
           f"{get_column_letter(INSERT_AT + GROUP_WIDTH - 1)} 已写入")

    # ---------------- 自检 ----------------
    checks = []

    def add(name, ok, detail=""):
        checks.append({"name": name, "ok": bool(ok), "detail": detail})

    add("盈利表新周期行数", new_total - new_first + 1 == ITEMS_PER_PERIOD + 1,
        f"{new_first}-{new_total}")
    add("盈利表总计行 SUM 覆盖 9 行",
        all(str(ws_profit.cell(row=new_total, column=c).value or "").startswith(
            f"=SUM({get_column_letter(c)}{new_first}") for c in (4, 5, 8, 13, 14, 17)))
    add("上一周期总计行已修", bool(fixed), ", ".join(fixed))
    expect_cols = last_col if rate_exists else last_col + GROUP_WIDTH
    add("退款率表列数正确", int(ws_rate.max_column or 0) == expect_cols,
        f"{last_col} -> {ws_rate.max_column}")
    add("退款率表新块标题", all(
        str(ws_rate.cell(row=1, column=INSERT_AT + i).value or "") == h
        for i, h in enumerate(headers)))
    add("1 表仍引用 H/J",
        str(wb["1"]["AY4"].value or "") == "=退款率!H2"
        and str(wb["1"]["BB4"].value or "") == "=退款率!J2")

    ok = all(c["ok"] for c in checks)
    saved = False
    if ok and save:
        try:
            wb.save(book)
            saved = True
            logger(f"[roi] 已保存: {book}")
        except PermissionError:
            logger("[roi] 保存失败：文件被占用，请先关闭 WPS/Excel")
    elif ok:
        logger("[roi] 仅校验（未保存）")
    elif not save:
        logger("[roi] 自检未通过，未保存")

    return {
        "book": str(book),
        "label7": label7,
        "label30": label30,
        "new_first": new_first,
        "new_total": new_total,
        "items": len(products),
        "total_cost7": round(sum(week.get(p, [0, 0])[0] for p, _n in products), 2),
        "total_gmv7": round(sum(week.get(p, [0, 0])[1] for p, _n in products), 2),
        "total_cost30": round(sum(thirty[p][0] for p, _n in products), 2),
        "total_gmv30": round(sum(thirty[p][1] for p, _n in products), 2),
        "checks": checks,
        "saved": saved,
        "overwritten": bool(dup),
        "files": {"week_csv": str(week_csv) if week_csv else "",
                  "range_csv": str(range_csv) if range_csv else "",
                  "syt_week": str(syt_week) if syt_week else "",
                  "syt_range": str(syt_range) if syt_range else ""},
    }


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True, help="7日开始 YYYY-MM-DD")
    ap.add_argument("--end", required=True, help="7日结束 YYYY-MM-DD")
    ap.add_argument("--range-start", required=True, help="30日开始 YYYY-MM-DD")
    ap.add_argument("--range-end", required=True, help="30日结束 YYYY-MM-DD")
    ap.add_argument("--book")
    ap.add_argument("--account")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    cfg = kx.load_config()
    if args.account:
        acct = kx.find_login_account(cfg, args.account)
        if not acct:
            print(f"未找到账号: {args.account}")
            return 2
        cfg = kx.account_cfg(cfg, acct)
    else:
        cfg = kx.default_account_cfg(cfg)

    try:
        res = run_merge(cfg, args.start, args.end, args.range_start, args.range_end,
                        book_path=Path(args.book) if args.book else None,
                        save=not args.dry_run)
    except Exception as e:  # noqa: BLE001
        print(f"[roi] 失败: {type(e).__name__}: {e}")
        return 1
    print("\n===== 校验清单 =====")
    for i, c in enumerate(res["checks"], 1):
        print(f"{i}. [{'PASS' if c['ok'] else 'FAIL'}] {c['name']}" + (f"  {c['detail']}" if c["detail"] else ""))
    print(f"===== {'已保存' if res['saved'] else '未保存'} =====")
    return 0 if all(c["ok"] for c in res["checks"]) else 1


if __name__ == "__main__":
    sys.exit(main())
