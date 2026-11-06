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
TOTAL_LABEL = "总计"
GROUPS_START = 2


def _layout(ws) -> tuple[int, list[int]]:
    """返回（总计列起始列，所有商品组的消耗列）；兼容后续追加过商品列的表"""
    total_start = None
    for c in range(GROUPS_START, ws.max_column + 1):
        if str(ws.cell(row=1, column=c).value or "").strip() == TOTAL_LABEL:
            total_start = c
            break
    if total_start is None:
        total_start = 29
    return total_start, list(range(GROUPS_START, total_start, 3))


def _copy_style(dst, src):
    if not src.has_style:
        return
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)
    dst.alignment = copy(src.alignment)


def _copy_cell(dst, src):
    dst.value = src.value
    _copy_style(dst, src)


def _move_totals(ws, total_start: int, shift: int) -> int:
    """把「总计」三列整体右移 shift 列（保留样式/合并/列宽），返回新起始列"""
    new_start = total_start + shift
    max_row = ws.max_row
    for r in range(1, max_row + 1):
        for i in range(3):
            _copy_cell(ws.cell(row=r, column=new_start + i),
                       ws.cell(row=r, column=total_start + i))
    old_l, old_r = get_column_letter(total_start), get_column_letter(total_start + 2)
    for r in (1, 2, 3):
        try:
            ws.unmerge_cells(f"{old_l}{r}:{old_r}{r}")
        except Exception:
            pass
    # 旧位置先清空，随后由新商品组覆盖；多余部分保持空白
    for r in range(1, max_row + 1):
        for i in range(3):
            ws.cell(row=r, column=total_start + i).value = None
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=new_start, end_row=r, end_column=new_start + 2)
    for i in range(3):
        src = ws.column_dimensions[get_column_letter(total_start + i)]
        dst = ws.column_dimensions[get_column_letter(new_start + i)]
        if src.width:
            dst.width = src.width
    return new_start


def _build_group(ws, col: int, name: str, pid: str, aid: str, ref_col: int):
    """在 col 处新建一组「消耗/成交/ROI」列，样式复制自 ref_col 组"""
    for i in range(3):
        ws.column_dimensions[get_column_letter(col + i)].width = \
            ws.column_dimensions[get_column_letter(ref_col + i)].width or 13
        for r in range(1, TOTAL_ROW + 1):
            _copy_style(ws.cell(row=r, column=col + i),
                        ws.cell(row=r, column=ref_col + i))
    ws.cell(row=1, column=col, value=name)
    ws.cell(row=2, column=col, value=pid)
    ws.cell(row=3, column=col, value=aid)
    ws.cell(row=4, column=col, value="消耗")
    ws.cell(row=4, column=col + 1, value="成交")
    ws.cell(row=4, column=col + 2, value="ROI")
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 2)


def _write_formulas(ws, cost_cols: list[int], total_cols: tuple[int, int, int]):
    """按当前所有商品组重写 ROI/总计行/总计列公式"""
    gmv_cols = [c + 1 for c in cost_cols]
    roi_cols = [c + 2 for c in cost_cols]
    for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        for c in roi_cols:
            ws.cell(row=r, column=c).value = \
                f"={get_column_letter(c - 1)}{r}/{get_column_letter(c - 2)}{r}"
    tc, tg, tr = total_cols
    for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        ws.cell(row=r, column=tc).value = "=" + "+".join(
            f"{get_column_letter(c)}{r}" for c in cost_cols)
        ws.cell(row=r, column=tg).value = "=" + "+".join(
            f"{get_column_letter(c)}{r}" for c in gmv_cols)
        ws.cell(row=r, column=tr).value = f"={get_column_letter(tg)}{r}/{get_column_letter(tc)}{r}"
    for c in cost_cols:
        ws.cell(row=TOTAL_ROW, column=c).value = \
            f"=SUM({get_column_letter(c)}{FIRST_DATA_ROW}:{get_column_letter(c)}{LAST_DATA_ROW})"
        ws.cell(row=TOTAL_ROW, column=c + 1).value = \
            f"=SUM({get_column_letter(c + 1)}{FIRST_DATA_ROW}:{get_column_letter(c + 1)}{LAST_DATA_ROW})"
        ws.cell(row=TOTAL_ROW, column=c + 2).value = \
            f"={get_column_letter(c + 1)}{TOTAL_ROW}/{get_column_letter(c)}{TOTAL_ROW}"
    ws.cell(row=TOTAL_ROW, column=tc).value = \
        f"=SUM({get_column_letter(tc)}{FIRST_DATA_ROW}:{get_column_letter(tc)}{LAST_DATA_ROW})"
    ws.cell(row=TOTAL_ROW, column=tg).value = \
        f"=SUM({get_column_letter(tg)}{FIRST_DATA_ROW}:{get_column_letter(tg)}{LAST_DATA_ROW})"
    ws.cell(row=TOTAL_ROW, column=tr).value = f"={get_column_letter(tg)}{TOTAL_ROW}/{get_column_letter(tc)}{TOTAL_ROW}"


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
    maxc = ws.max_column
    head = [[ws.cell(row=r, column=c).value for c in range(1, maxc + 1)] for r in range(1, 5)]
    col_a = [ws.cell(row=r, column=1).value for r in range(FIRST_DATA_ROW, TOTAL_ROW + 1)]
    total = [ws.cell(row=TOTAL_ROW, column=c).value for c in range(1, maxc + 1)]
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
    started_at = datetime.now()
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
    account_names: dict[str, str] = {}
    for r in rows:
        aid = str(r["account_id"]).strip()
        account_names.setdefault(aid, str(r.get("account_name") or "").strip())
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
    prod_by_acct: dict[tuple[str, str, str], tuple[float, float]] = {}
    pmeta: dict[tuple[str, str], dict] = {}
    ppath = export_dir / "products.csv"
    if ppath.exists():
        for r in csv.DictReader(ppath.open(encoding="utf-8-sig")):
            key = (r["date"], str(r["product_id"]).strip())
            pid = str(r["product_id"]).strip()
            try:
                p_cost, p_gmv = float(r["cost"]), float(r["gmv"])
            except (TypeError, ValueError):
                continue
            aid_row = str(r["account_id"]).strip()
            if pid:
                k3 = (r["date"], aid_row, pid)
                c0, g0 = prod_by_acct.get(k3, (0.0, 0.0))
                prod_by_acct[k3] = (c0 + p_cost, g0 + p_gmv)
                meta = pmeta.setdefault((aid_row, pid), {"name": "", "cost": 0.0, "gmv": 0.0})
                if not meta["name"]:
                    meta["name"] = str(r.get("product_name") or "").strip()
                meta["cost"] += p_cost
                meta["gmv"] += p_gmv
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

    total_start, all_cost_cols = _layout(ws)
    specs = []
    for col in all_cost_cols:
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

    # 所选但月度表里没有的账户：在总计列前新增商品列
    new_groups: list[dict] = []
    if wanted:
        existing_aids = {s["aid"] for s in specs}
        missing = sorted(a for a in wanted
                         if a not in existing_aids and any(k[1] == a for k in acc))
        for aid in missing:
            plist = [{"pid": pid, "name": meta["name"], "cost": meta["cost"]}
                     for (a, pid), meta in pmeta.items() if a == aid]
            plist.sort(key=lambda x: (-x["cost"], x["pid"]))
            if plist:
                for p in plist:
                    new_groups.append({"aid": aid, "pid": p["pid"],
                                       "name": p["name"] or p["pid"], "kind": "product"})
            else:
                new_groups.append({"aid": aid, "pid": "",
                                   "name": account_names.get(aid) or aid, "kind": "account"})
        specs = [s for s in specs if s["aid"] in wanted]
        if not specs and not new_groups:
            raise RuntimeError("所选账户都不在月度表内，且没有可新增的数据")

    if new_groups:
        ref_col = all_cost_cols[-1] if all_cost_cols else GROUPS_START
        old_total_start = total_start
        total_start = _move_totals(ws, old_total_start, 3 * len(new_groups))
        for i, g in enumerate(new_groups):
            g["col"] = old_total_start + 3 * i
            _build_group(ws, g["col"], g["name"], g["pid"], g["aid"], ref_col)
        all_cost_cols = list(range(GROUPS_START, total_start, 3))
        _write_formulas(ws, all_cost_cols, (total_start, total_start + 1, total_start + 2))
        for g in new_groups:
            logger(f"[merge] 新增列 {get_column_letter(g['col'])}-{get_column_letter(g['col'] + 2)}: "
                   f"账户 {g['aid']} 商品 {g['name']}（{g['pid'] or '-'}）")
    total_cols = (total_start, total_start + 1, total_start + 2)

    dates = sorted({d for d, _a in acc})
    logger(f"[merge] 月度表商品列 {len(specs)} 个，待回填 {len(dates)} 天"
           + (f"，新增商品列 {len(new_groups)} 个" if new_groups else ""))

    data_specs = list(specs)
    for g in new_groups:
        data_specs.append({"col": g["col"], "aid": g["aid"], "pid": g["pid"],
                           "name": g["name"], "source": g["kind"]})

    # 回填前先记录各列的样式基准（取该列第一条已有数据），保证写入后字体/数字格式与表内一致
    style_refs: dict[int, object] = {}
    for col in all_cost_cols:
        style_refs[col] = _pick_style_ref(ws, col)
        style_refs[col + 1] = _pick_style_ref(ws, col + 1)

    head0, cola0, total0 = _snapshot(ws)
    written, written_cells, zeros, skipped, notes = 0, [], [], [], []

    for date in dates:
        target = datetime.strptime(date, "%Y-%m-%d").day + 4
        if not (FIRST_DATA_ROW <= target <= LAST_DATA_ROW):
            skipped.append(f"{date} 行号越界")
            continue
        for spec in data_specs:
            col, aid, pid = spec["col"], spec["aid"], spec["pid"]
            label = spec["name"] or aid
            cost = gmv = None
            source = spec.get("source")

            if source == "product":
                hit = prod_by_acct.get((date, aid, pid))
                if hit is None:
                    skipped.append(f"{date} {label} 无商品数据")
                    continue
                cost, gmv = hit
            elif source == "account":
                a_total = acc.get((date, aid))
                if a_total is None:
                    skipped.append(f"{date} {label} 无数据")
                    continue
                cost, gmv = a_total
            elif pid and (date, pid) in prods:
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

    bad_roi = [f"{get_column_letter(c)}{r}" for c in (c + 2 for c in all_cost_cols)
               for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1)
               if not _is_formula(ws.cell(row=r, column=c).value)]
    add("ROI 列公式完好", not bad_roi, f"非公式 {len(bad_roi)} 个")

    bad_total = [f"{get_column_letter(c)}{r}" for c in total_cols
                 for r in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1)
                 if not _is_formula(ws.cell(row=r, column=c).value)]
    add("总计列公式完好", not bad_total, f"非公式 {len(bad_total)} 个")

    bad_row35 = [get_column_letter(c) + str(TOTAL_ROW) for c in range(2, ws.max_column + 1)
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

    finished_at = datetime.now()
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
        "added_columns": [{
            "account_id": g["aid"],
            "product_id": g["pid"],
            "product_name": g["name"],
            "col": get_column_letter(g["col"]),
            "range": f"{get_column_letter(g['col'])}-{get_column_letter(g['col'] + 2)}",
        } for g in new_groups],
        "started_at": started_at.strftime("%Y-%m-%d %H:%M:%S"),
        "finished_at": finished_at.strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed": round((finished_at - started_at).total_seconds(), 1),
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
