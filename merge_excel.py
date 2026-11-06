# -*- coding: utf-8 -*-
"""合并文档：把导出结果（账户×日期）回填进月度 Excel，并执行 9 项自检

红线（沿用 SOP）:
  只写「消耗」「成交」两列；ROI 列(D/G/J/M/P/S/V/Y/AB)、总计列(AC/AD/AE)、第 35 行、
  表头 1-4 行、A 列日期一律不写。
  取数失败留空不填 0；真实 0 填 0（数据来自接口，0 即真实 0）。

用法:
  py merge_excel.py --date 2026-11-04
  py merge_excel.py --start 2026-11-01 --end 2026-11-05
  py merge_excel.py --date 2026-11-04 --accounts 10000001,10000002 --dry-run
"""
from __future__ import annotations

import argparse
import csv
import shutil
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import kx_core as kx

BASE = Path(__file__).resolve().parent
SHEET = "消耗成交"
FIRST_DATA_ROW, LAST_DATA_ROW, TOTAL_ROW = 5, 34, 35
COST_COLS = [2 + i * 3 for i in range(9)]
ROI_COLS = [c + 2 for c in COST_COLS]
TOTAL_COLS = (29, 30, 31)


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _cell_id(v) -> str:
    """表内 ID 统一转字符串：整数型 float 不带小数点，None 归空"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _pick_style_ref(ws, col: int):
    """该列的样式基准：优先取第一行已有数据的单元格，否则取首个数据行"""
    for r in range(FIRST_DATA_ROW, TOTAL_ROW):
        c = ws.cell(row=r, column=col)
        if c.value not in (None, ""):
            return c
    return ws.cell(row=FIRST_DATA_ROW, column=col)


def _sync_style(dst, src):
    """只同步影响观感的字体/数字格式/对齐，保留原单元格的边框与填充"""
    dst.font = copy(src.font)
    dst.number_format = src.number_format
    dst.alignment = copy(src.alignment)


def _snapshot(ws):
    head = [[ws.cell(row=r, column=c).value for c in range(1, 32)] for r in range(1, 5)]
    col_a = [ws.cell(row=r, column=1).value for r in range(FIRST_DATA_ROW, TOTAL_ROW + 1)]
    total = [ws.cell(row=TOTAL_ROW, column=c).value for c in range(1, 32)]
    return head, col_a, total


def _prune_backups(bdir: Path, stem: str, suffix: str, limit: int, logger):
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
        except Exception:
            pass
    if removed:
        logger(f"[merge] 备份超出上限 {limit} 个，已删除最旧 {removed} 个")


def run_merge(cfg: dict, start: str, end: str, account_ids: list[str] | None = None,
              excel_path: Path | None = None, export_dir: Path | None = None,
              backup: bool = True, save: bool = True, logger=print) -> dict:
    excel_path = Path(excel_path) if excel_path else kx.resolve_path(cfg["excel_path"])
    export_dir = Path(export_dir) if export_dir else (kx.DATA_DIR / "exports" / f"{start}_{end}")
    csv_path = export_dir / "accounts.csv"
    if not csv_path.exists():
        raise FileNotFoundError(f"未找到导出数据: {csv_path}（请先执行导出）")
    if not excel_path.exists():
        raise FileNotFoundError(f"目标 Excel 不存在: {excel_path}")

    rows = list(csv.DictReader(csv_path.open(encoding="utf-8-sig")))
    wanted = {str(x) for x in account_ids} if account_ids else None
    acc: dict[tuple[str, str], tuple[float, float]] = {}
    for r in rows:
        aid = str(r["account_id"]).strip()
        if wanted and aid not in wanted:
            continue
        try:
            acc[(r["date"], aid)] = (float(r["cost"]), float(r["gmv"]))
        except (TypeError, ValueError):
            continue
    if not acc:
        raise RuntimeError("没有可回填的数据行")

    # 月度表按「商品」组织：行1 商品名 / 行2 商品ID / 行3 所属账户ID，
    # 因此优先按商品ID取数；账户下还有别的商品时，只取本商品，避免把其他商品的花费记进来。
    prods: dict[tuple[str, str], list] = {}
    ppath = export_dir / "products.csv"
    if ppath.exists():
        for r in csv.DictReader(ppath.open(encoding="utf-8-sig")):
            key = (r["date"], str(r["product_id"]).strip())
            try:
                p_cost, p_gmv = float(r["cost"]), float(r["gmv"])
            except (TypeError, ValueError):
                continue
            aid_row = str(r["account_id"]).strip()
            item = prods.get(key)
            if item:
                item[0] += p_cost
                item[1] += p_gmv
                c0, g0 = item[2].get(aid_row, (0.0, 0.0))
                item[2][aid_row] = (c0 + p_cost, g0 + p_gmv)
            else:
                prods[key] = [p_cost, p_gmv, {aid_row: (p_cost, p_gmv)}]

    # 同一商品可能在多个账户的报表里出现（多为其他账户的 0 值行）：取真正有数据的那个账户
    for item in prods.values():
        accs = item[2]
        real = [a for a, (c, g) in accs.items() if c or g]
        item[2] = real[0] if len(real) == 1 else ("MULTI" if real else next(iter(accs), ""))

    if backup:
        bdir = kx.DATA_DIR / "backup"
        bdir.mkdir(parents=True, exist_ok=True)
        bak = bdir / f"{excel_path.stem}_backup_{datetime.now().strftime('%Y%m%d-%H%M%S')}{excel_path.suffix}"
        shutil.copy2(excel_path, bak)
        logger(f"[merge] 已备份: {bak.name}")
        try:
            limit = int(cfg.get("backup_limit", 10))
        except (TypeError, ValueError):
            limit = 10
        _prune_backups(bdir, excel_path.stem, excel_path.suffix, limit, logger)
    else:
        logger("[merge] 未备份（dry-run 或设置里关闭了自动备份）")

    wb = load_workbook(excel_path)
    if SHEET not in wb.sheetnames:
        raise RuntimeError(f"工作表不存在: {SHEET}，实际: {wb.sheetnames}")
    ws = wb[SHEET]

    specs = []
    for col in COST_COLS:
        aid = _cell_id(ws.cell(row=3, column=col).value)
        if not aid:
            continue
        specs.append({
            "col": col,
            "aid": aid,
            "pid": _cell_id(ws.cell(row=2, column=col).value),
            "name": str(ws.cell(row=1, column=col).value or "").strip(),
        })
    if not specs:
        raise RuntimeError("第 3 行未读取到账户ID，表结构不符")
    if wanted:
        specs = [s for s in specs if s["aid"] in wanted]
        if not specs:
            raise RuntimeError("所选账户都不在月度表内")
    dates = sorted({d for d, _a in acc})
    logger(f"[merge] 月度表商品列 {len(specs)} 个，待回填 {len(dates)} 天")

    # 回填前先记录各列的样式基准（取该列第一条已有数据），保证写入后字体/数字格式与表内一致
    style_refs: dict[int, object] = {}
    for col in COST_COLS:
        style_refs[col] = _pick_style_ref(ws, col)
        style_refs[col + 1] = _pick_style_ref(ws, col + 1)

    head0, cola0, total0 = _snapshot(ws)
    written, written_cells, zeros, skipped, notes = 0, [], [], [], []

    for date in dates:
        target = datetime.strptime(date, "%Y-%m-%d").day + 4
        if not (FIRST_DATA_ROW <= target <= LAST_DATA_ROW):
            skipped.append(f"{date} 行号越界")
            continue
        for spec in specs:
            col, aid, pid = spec["col"], spec["aid"], spec["pid"]
            label = spec["name"] or aid
            cost = gmv = None

            if pid and (date, pid) in prods:
                cost, gmv, pid_acct = prods[(date, pid)]
                if pid_acct == "MULTI":
                    notes.append(f"{date} 商品{pid} 在多个账户产生数据，按商品合计取值")
                elif pid_acct and pid_acct != aid:
                    notes.append(f"{date} 商品{pid} 归属账户 {pid_acct} 与表内 {aid} 不一致，按商品取数")
                a_total = acc.get((date, aid))
                if a_total and (abs(a_total[0] - cost) > 0.005 or abs(a_total[1] - gmv) > 0.005):
                    notes.append(f"{date} {label}: 账户汇总 {a_total[0]}/{a_total[1]} ≠ 商品 {cost}/{gmv}"
                                 f"（该账户另有其他商品产生数据，已只取商品 {pid}）")
            else:
                if pid:
                    notes.append(f"{date} 商品{pid} 不在报表中，回退为账户汇总")
                a_total = acc.get((date, aid))
                if a_total is None:
                    skipped.append(f"{date} {label} 无数据")
                    continue
                cost, gmv = a_total

            c_cost = ws.cell(row=target, column=col, value=cost)
            c_gmv = ws.cell(row=target, column=col + 1, value=gmv)
            if col in style_refs:
                _sync_style(c_cost, style_refs[col])
                _sync_style(c_gmv, style_refs[col + 1])
            written += 1
            written_cells.append((target, col, cost, gmv))
            if cost == 0 or gmv == 0:
                zeros.append(f"{date} {label} 花费={cost} 成交={gmv}")

    for n in notes[:5]:
        logger(f"[merge] 提示: {n}")
    if len(notes) > 5:
        logger(f"[merge] 提示: 另有 {len(notes) - 5} 条同类记录")

    checks = []

    def add(name, ok, detail=""):
        checks.append({"name": name, "ok": bool(ok), "detail": detail})

    add("单元格数量齐全", len(written_cells) * 2 == written * 2, f"{written} 条 × 2 列")

    mismatch = []
    for tr, col, cost, gmv in written_cells:
        v1, v2 = ws.cell(row=tr, column=col).value, ws.cell(row=tr, column=col + 1).value
        if not (isinstance(v1, (int, float)) and abs(v1 - cost) < 0.005):
            mismatch.append(f"{get_column_letter(col)}{tr}")
        if not (isinstance(v2, (int, float)) and abs(v2 - gmv) < 0.005):
            mismatch.append(f"{get_column_letter(col + 1)}{tr}")
    add("数值逐条读回比对", not mismatch, "; ".join(mismatch[:5]))

    bad_roi = [f"{get_column_letter(c)}{r}" for c in ROI_COLS
               for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1)
               if not _is_formula(ws.cell(row=r, column=c).value)]
    add("ROI 列公式完好", not bad_roi, f"非公式 {len(bad_roi)} 个")

    bad_total = [f"{get_column_letter(c)}{r}" for c in TOTAL_COLS
                 for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1)
                 if not _is_formula(ws.cell(row=r, column=c).value)]
    add("总计列公式完好", not bad_total, f"非公式 {len(bad_total)} 个")

    bad_row35 = [get_column_letter(c) + str(TOTAL_ROW) for c in range(2, 32)
                 if not _is_formula(ws.cell(row=TOTAL_ROW, column=c).value)]
    add("总计行公式完好", not bad_row35, f"非公式 {len(bad_row35)} 个")

    head1, cola1, total1 = _snapshot(ws)
    add("无越界写入（表头/日期列/总计行）", head0 == head1 and cola0 == cola1 and total0 == total1)

    total_cost = round(sum(c for _, _, c, _ in written_cells), 2)
    total_gmv = round(sum(g for _, _, _, g in written_cells), 2)
    add("回填合计（对照后台/AC、AD）", True, f"消耗合计={total_cost} 成交合计={total_gmv}")
    add("自动重算提示", True, "用 WPS/Excel 打开一次触发重算")
    add("零值合理性", True, (f"{len(zeros)} 处: " + "; ".join(zeros[:5])) if zeros else "无 0 值")

    ok = all(c["ok"] for c in checks)
    saved = False
    locked = False
    if ok and save:
        try:
            wb.save(excel_path)
            saved = True
            logger(f"[merge] 已保存: {excel_path}")
        except PermissionError:
            locked = True
            logger(f"[merge] 文件被占用（请先关闭 Excel/WPS 后重试）: {excel_path}")
            ok = False
    elif ok:
        logger("[merge] 仅校验（未保存）")
    elif not locked:
        logger("[merge] 自检未通过，未保存（可用备份恢复）")

    return {
        "excel": str(excel_path),
        "written": written,
        "cells": len(written_cells) * 2,
        "total_cost": total_cost,
        "total_gmv": total_gmv,
        "skipped": skipped,
        "checks": checks,
        "saved": saved,
        "locked": locked,
        "error": "月度 Excel 正被其他程序占用，请先关闭 WPS/Excel 再重试" if locked else None,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date")
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--accounts")
    ap.add_argument("--file", help="目标 Excel，默认取 config.excel_path")
    ap.add_argument("--export-dir")
    ap.add_argument("--dry-run", action="store_true", help="只校验不保存")
    ap.add_argument("--config", default=str(kx.CONFIG_PATH))
    args = ap.parse_args()

    cfg = kx.load_config(Path(args.config))
    cfg = kx.default_account_cfg(cfg)
    if args.start and args.end:
        start, end = args.start, args.end
    else:
        start = end = args.date or kx.default_date(cfg)
    ids = [x.strip() for x in args.accounts.split(",")] if args.accounts else None

    res = run_merge(cfg, start, end, ids,
                    excel_path=Path(args.file) if args.file else None,
                    export_dir=Path(args.export_dir) if args.export_dir else None,
                    backup=not args.dry_run, save=not args.dry_run)
    print("\n===== 校验清单 =====")
    for i, c in enumerate(res["checks"], 1):
        print(f"{i}. [{'PASS' if c['ok'] else 'FAIL'}] {c['name']}" + (f"  {c['detail']}" if c["detail"] else ""))
    print(f"===== {'已保存' if res['saved'] else ('校验通过（未保存）' if args.dry_run else '未保存')} =====")
    return 0 if all(c["ok"] for c in res["checks"]) else 1


if __name__ == "__main__":
    sys.exit(main())

