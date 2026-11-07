# -*- coding: utf-8 -*-
"""合并文档：把导出结果（账户×日期）回填进月度 Excel，并执行 9 项自检

红线（沿用 SOP）:
  只写「消耗」「成交」两列；ROI 列、总计列、总计行、表头 1-4 行、A 列日期一律不写。
  取数失败留空不填 0；真实 0 填 0（数据来自接口，0 即真实 0）。

模板兼容:
  数据行从第 5 行开始，总计行按 A 列「总计」自动识别。
  30 天模板（第 35 行总计）只支持 1~30 日；31 天模板（第 36 行总计）支持 1~31 日。
  日期越界、跨月、导出失败账户未处理时一律拒绝保存，不再“静默成功”。

用法:
  py merge_excel.py --date 2026-11-04
  py merge_excel.py --start 2026-11-01 --end 2026-11-05
  py merge_excel.py --date 2026-11-04 --accounts 10000001,10000002 --dry-run
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
from copy import copy
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import kx_core as kx

BASE = Path(__file__).resolve().parent
SHEET = "消耗成交"
FIRST_DATA_ROW = 5
TOTAL_ROW_FALLBACK = 35
TOTAL_LABEL = "总计"
GROUPS_START = 2
GROUP_WIDTH = 3


# ---------------------------------------------------------------- layout
def _layout(ws) -> tuple[int, list[int], int]:
    """返回（总计列起始列，所有商品组的消耗列，总计行）

    总计列：第 1 行值为「总计」的那一列（默认 AC=29）。
    总计行：A 列值为「总计」的那一行（30 天模板=35，31 天模板=36）。
    """
    max_col = int(ws.max_column or 0)
    total_start = None
    for c in range(GROUPS_START, max_col + 1):
        if str(ws.cell(row=1, column=c).value or "").strip() == TOTAL_LABEL:
            total_start = c
            break
    if total_start is None:
        total_start = 29

    max_row = int(ws.max_row or 0)
    total_row = None
    for r in range(FIRST_DATA_ROW, max_row + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == TOTAL_LABEL:
            total_row = r
            break
    if total_row is None or total_row <= FIRST_DATA_ROW:
        total_row = TOTAL_ROW_FALLBACK
    return total_start, list(range(GROUPS_START, total_start, GROUP_WIDTH)), total_row


def _table_month(ws, first: int, last: int) -> tuple[int | None, int] | None:
    """从 A 列日期推断月度表所属月份；识别不了返回 None。

    A 列三种写法都认：datetime/date、文本日期、Excel 日期序列号（如 46327）。
    序列号按 1900 日期系统换算，并用 20000~80000 区间排除普通数字。
    """
    for r in range(first, last + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, datetime):
            return v.year, v.month
        if isinstance(v, date):
            return v.year, v.month
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if 20000 <= float(v) <= 80000:
                d = datetime(1899, 12, 30) + timedelta(days=float(v))
                return d.year, d.month
            continue
        s = str(v or "").strip()
        if not s:
            continue
        m = re.match(r"^(\d{4})[-/.年](\d{1,2})", s)
        if m:
            return int(m.group(1)), int(m.group(2))
        m = re.match(r"^(\d{1,2})月(\d{1,2})日", s)
        if m:
            return None, int(m.group(1))
        m = re.match(r"^(\d{1,2})[-/.](\d{1,2})$", s)
        if m:
            return None, int(m.group(1))
    return None


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
    max_row = int(ws.max_row or 0)
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


def _build_group(ws, col: int, name: str, pid: str, aid: str, ref_col: int, total_row: int):
    """在 col 处新建一组「消耗/成交/ROI」列，样式复制自 ref_col 组"""
    for i in range(3):
        ws.column_dimensions[get_column_letter(col + i)].width = \
            ws.column_dimensions[get_column_letter(ref_col + i)].width or 13
        for r in range(1, total_row + 1):
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


def _write_formulas(ws, cost_cols: list[int], total_cols: tuple[int, int, int],
                    data_first: int, data_last: int, total_row: int):
    """按当前所有商品组重写 ROI/总计行/总计列公式"""
    gmv_cols = [c + 1 for c in cost_cols]
    roi_cols = [c + 2 for c in cost_cols]
    for r in range(data_first, data_last + 1):
        for c in roi_cols:
            ws.cell(row=r, column=c).value = \
                f"={get_column_letter(c - 1)}{r}/{get_column_letter(c - 2)}{r}"
    tc, tg, tr = total_cols
    for r in range(data_first, data_last + 1):
        ws.cell(row=r, column=tc).value = "=" + "+".join(
            f"{get_column_letter(c)}{r}" for c in cost_cols)
        ws.cell(row=r, column=tg).value = "=" + "+".join(
            f"{get_column_letter(c)}{r}" for c in gmv_cols)
        ws.cell(row=r, column=tr).value = f"={get_column_letter(tg)}{r}/{get_column_letter(tc)}{r}"
    for c in cost_cols:
        ws.cell(row=total_row, column=c).value = \
            f"=SUM({get_column_letter(c)}{data_first}:{get_column_letter(c)}{data_last})"
        ws.cell(row=total_row, column=c + 1).value = \
            f"=SUM({get_column_letter(c + 1)}{data_first}:{get_column_letter(c + 1)}{data_last})"
        ws.cell(row=total_row, column=c + 2).value = \
            f"={get_column_letter(c + 1)}{total_row}/{get_column_letter(c)}{total_row}"
    ws.cell(row=total_row, column=tc).value = \
        f"=SUM({get_column_letter(tc)}{data_first}:{get_column_letter(tc)}{data_last})"
    ws.cell(row=total_row, column=tg).value = \
        f"=SUM({get_column_letter(tg)}{data_first}:{get_column_letter(tg)}{data_last})"
    ws.cell(row=total_row, column=tr).value = \
        f"={get_column_letter(tg)}{total_row}/{get_column_letter(tc)}{total_row}"


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _cell_id(v) -> str:
    """表内 ID 统一转字符串：整数型 float 不带小数点，None 归空"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _pick_style_ref(ws, col: int, total_row: int):
    """该列的样式基准：优先取第一行已有数据的单元格，否则取首个数据行"""
    for r in range(FIRST_DATA_ROW, total_row):
        c = ws.cell(row=r, column=col)
        if c.value not in (None, ""):
            return c
    return ws.cell(row=FIRST_DATA_ROW, column=col)


def _sync_style(dst, src):
    """只同步影响观感的字体/数字格式/对齐，保留原单元格的边框与填充"""
    dst.font = copy(src.font)
    dst.number_format = src.number_format
    dst.alignment = copy(src.alignment)


def _snapshot(ws, total_row: int):
    maxc = int(ws.max_column or 1)
    head = [[ws.cell(row=r, column=c).value for c in range(1, maxc + 1)] for r in range(1, 5)]
    col_a = [ws.cell(row=r, column=1).value for r in range(FIRST_DATA_ROW, total_row + 1)]
    total = [ws.cell(row=total_row, column=c).value for c in range(1, maxc + 1)]
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


def _summary_failures(export_dir: Path) -> list[dict]:
    p = export_dir / "summary.json"
    if not p.exists():
        return []
    try:
        j = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return []
    return list(j.get("failures") or [])


def _read_csv_rows(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def read_product_groups(path: Path) -> list[dict]:
    """读取月度表第 1~3 行的商品组（商品名 / 商品ID / 所属账户ID）"""
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb[SHEET] if SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
        _total_start, cost_cols, _total_row = _layout(ws)
        out = []
        for col in cost_cols:
            aid = _cell_id(ws.cell(row=3, column=col).value)
            if not aid:
                continue
            out.append({
                "account_id": aid,
                "product_id": _cell_id(ws.cell(row=2, column=col).value),
                "product_name": str(ws.cell(row=1, column=col).value or "").strip(),
            })
        return out
    finally:
        wb.close()


# ---------------------------------------------------------------- merge
def run_merge(cfg: dict, start: str, end: str, account_ids: list[str] | None = None,
              excel_path: Path | None = None, export_dir: Path | None = None,
              backup: bool = True, save: bool = True,
              allow_partial: bool = False, logger=print) -> dict:
    started_at = datetime.now()
    if excel_path is None:
        raw_excel = str(cfg.get("excel_path") or "").strip()
        if not raw_excel:
            raise RuntimeError("未配置月度表路径 excel_path")
        excel_path = kx.resolve_path(raw_excel)
    else:
        excel_path = Path(excel_path)
    export_dir = Path(export_dir) if export_dir else (kx.DATA_DIR / "exports" / f"{start}_{end}")
    csv_path = export_dir / "accounts.csv"
    if not csv_path.exists():
        raise FileNotFoundError(f"未找到导出数据: {csv_path}（请先执行导出）")
    if not excel_path.exists():
        raise FileNotFoundError(f"目标 Excel 不存在: {excel_path}")

    failures = _summary_failures(export_dir)
    if failures and not allow_partial:
        ids = "、".join(str(f.get("account_id") or "?") for f in failures[:5])
        raise RuntimeError(
            f"导出存在失败账户（{len(failures)} 个：{ids}），已拒绝合并以避免写入不完整数据；"
            "请重新导出，或显式允许部分合并")
    if failures:
        logger(f"[merge] 警告: 导出存在 {len(failures)} 个失败账户，按 allow_partial 继续")

    rows = _read_csv_rows(csv_path)
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
    prod_account_days: set[tuple[str, str]] = set()
    pmeta: dict[tuple[str, str], dict] = {}
    ppath = export_dir / "products.csv"
    if ppath.exists():
        for r in _read_csv_rows(ppath):
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
                prod_account_days.add((r["date"], aid_row))
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

    dates = sorted({d for d, _a in acc})
    for d in dates:
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
            raise RuntimeError(f"日期格式不合法: {d}")
    months = sorted({d[:7] for d in dates})
    if len(months) > 1:
        raise RuntimeError(f"日期范围跨月（{'、'.join(months)}），月度表只能合并单月数据")

    wb = load_workbook(excel_path)
    if SHEET not in wb.sheetnames:
        raise RuntimeError(f"工作表不存在: {SHEET}，实际: {wb.sheetnames}")
    ws = wb[SHEET]

    total_start, all_cost_cols, total_row = _layout(ws)
    data_last = total_row - 1
    if data_last < FIRST_DATA_ROW:
        raise RuntimeError(f"月度表数据行异常：总计行 {total_row}，首个数据行 {FIRST_DATA_ROW}")

    table_month = _table_month(ws, FIRST_DATA_ROW, data_last)
    month_ok = True
    if table_month:
        t_year, t_month = table_month
        d_year, d_month = int(months[0][:4]), int(months[0][5:7])
        month_ok = (t_month == d_month and (t_year is None or t_year == d_year))
        if not month_ok:
            raise RuntimeError(
                f"待合并数据月份 {months[0]} 与月度表日期 "
                f"{t_year if t_year is not None else '????'}-{t_month:02d} 不一致，已拒绝")
    else:
        logger("[merge] 提示: 未能从 A 列识别月度表月份，跳过月份校验")

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
            # 新增列第 1 行统一用账户名称，账户名缺失时才退回商品名/商品ID
            acct_name = account_names.get(aid) or ""
            if plist:
                for p in plist:
                    new_groups.append({"aid": aid, "pid": p["pid"],
                                       "name": acct_name or p["name"] or p["pid"],
                                       "kind": "product"})
            else:
                new_groups.append({"aid": aid, "pid": "",
                                   "name": acct_name or aid, "kind": "account"})
        specs = [s for s in specs if s["aid"] in wanted]
        if not specs and not new_groups:
            raise RuntimeError("所选账户都不在月度表内，且没有可新增的数据")

    if new_groups:
        ref_col = all_cost_cols[-1] if all_cost_cols else GROUPS_START
        old_total_start = total_start
        total_start = _move_totals(ws, old_total_start, GROUP_WIDTH * len(new_groups))
        for i, g in enumerate(new_groups):
            g["col"] = old_total_start + GROUP_WIDTH * i
            _build_group(ws, g["col"], g["name"], g["pid"], g["aid"], ref_col, total_row)
        all_cost_cols = list(range(GROUPS_START, total_start, GROUP_WIDTH))
        _write_formulas(ws, all_cost_cols, (total_start, total_start + 1, total_start + 2),
                        FIRST_DATA_ROW, data_last, total_row)
        for g in new_groups:
            logger(f"[merge] 新增列 {get_column_letter(g['col'])}-{get_column_letter(g['col'] + 2)}: "
                   f"账户 {g['aid']} 商品 {g['name']}（{g['pid'] or '-'}）")
    total_cols = (total_start, total_start + 1, total_start + 2)

    logger(f"[merge] 月度表商品列 {len(specs)} 个，待回填 {len(dates)} 天"
           + (f"，新增商品列 {len(new_groups)} 个" if new_groups else ""))

    data_specs = list(specs)
    for g in new_groups:
        data_specs.append({"col": g["col"], "aid": g["aid"], "pid": g["pid"],
                           "name": g["name"], "source": g["kind"]})

    # 回填前先记录各列的样式基准（取该列第一条已有数据），保证写入后字体/数字格式与表内一致
    style_refs: dict[int, object] = {}
    for col in all_cost_cols:
        style_refs[col] = _pick_style_ref(ws, col, total_row)
        style_refs[col + 1] = _pick_style_ref(ws, col + 1, total_row)

    head0, cola0, total0 = _snapshot(ws, total_row)

    def _resolve(d: str, spec: dict):
        """返回 ((cost, gmv) | None, notes, skip_reason)"""
        col, aid, pid = spec["col"], spec["aid"], spec["pid"]
        label = spec["name"] or aid
        source = spec.get("source")
        notes: list[str] = []
        if source == "product":
            hit = prod_by_acct.get((d, aid, pid))
            if hit is None:
                return None, notes, f"{d} {label} 无商品数据"
            return hit, notes, None
        if source == "account":
            a_total = acc.get((d, aid))
            if a_total is None:
                return None, notes, f"{d} {label} 无数据"
            return a_total, notes, None

        if pid and (d, pid) in prods:
            cost, gmv, pid_acct = prods[(d, pid)]
            if pid_acct == "MULTI":
                notes.append(f"{d} 商品{pid} 在多个账户产生数据，按商品合计取值")
            elif pid_acct and pid_acct != aid:
                notes.append(f"{d} 商品{pid} 归属账户 {pid_acct} 与表内 {aid} 不一致，按商品取数")
            a_total = acc.get((d, aid))
            if a_total and (abs(a_total[0] - cost) > 0.005 or abs(a_total[1] - gmv) > 0.005):
                notes.append(f"{d} {label}: 账户汇总 {a_total[0]}/{a_total[1]} ≠ 商品 {cost}/{gmv}"
                             f"（该账户另有其他商品产生数据，已只取商品 {pid}）")
            return (cost, gmv), notes, None

        # 商品ID在报表里找不到：若该账户当天确实有别的商品数据，绝不回退账户汇总，
        # 否则「账户总额」+「其他商品额」会在总计列里被重复相加。
        if (d, aid) in prod_account_days:
            notes.append(f"{d} {label}: 商品{pid or '(空)'}不在报表中，"
                         f"且该账户存在其他商品数据，已留空以避免重复计数")
            return None, notes, f"{d} {label} 商品不在报表，已留空避免重复计数"

        a_total = acc.get((d, aid))
        if a_total is None:
            return None, notes, f"{d} {label} 无数据"
        if pid:
            notes.append(f"{d} 商品{pid} 不在报表中，回退为账户汇总")
        return a_total, notes, None

    plan: list[tuple[int, int, float, float, str, dict]] = []
    skipped: list[str] = []
    out_of_range: list[str] = []
    notes_all: list[str] = []

    for d in dates:
        target = datetime.strptime(d, "%Y-%m-%d").day + 4
        if not (FIRST_DATA_ROW <= target <= data_last):
            out_of_range.append(
                f"{d} 目标行 {target} 超出模板数据行 {FIRST_DATA_ROW}-{data_last}")
            continue
        for spec in data_specs:
            value, notes, skip = _resolve(d, spec)
            notes_all.extend(notes)
            if value is None:
                if skip:
                    skipped.append(skip)
                continue
            plan.append((target, spec["col"], value[0], value[1], d, spec))

    written_cells: list[tuple[int, int, float, float]] = []
    for target, col, cost, gmv, _d, _spec in plan:
        c_cost = ws.cell(row=target, column=col, value=cost)
        c_gmv = ws.cell(row=target, column=col + 1, value=gmv)
        if col in style_refs:
            _sync_style(c_cost, style_refs[col])
            _sync_style(c_gmv, style_refs[col + 1])
        written_cells.append((target, col, cost, gmv))
    written = len(written_cells)

    for n in notes_all[:5]:
        logger(f"[merge] 提示: {n}")
    if len(notes_all) > 5:
        logger(f"[merge] 提示: 另有 {len(notes_all) - 5} 条同类记录")
    for n in out_of_range[:5]:
        logger(f"[merge] 无法写入: {n}")
    if skipped:
        logger(f"[merge] 跳过 {len(skipped)} 条（无数据/已留空）")

    checks = []

    def add(name, ok, detail=""):
        checks.append({"name": name, "ok": bool(ok), "detail": detail})

    unique_targets = len({(t, c) for t, c, _cost, _gmv in written_cells})
    add("写入计划与实际一致",
        written == len(plan) and unique_targets == len(plan),
        f"计划 {len(plan)} 条，写入 {written} 条")

    mismatch = []
    for tr, col, cost, gmv in written_cells:
        v1, v2 = ws.cell(row=tr, column=col).value, ws.cell(row=tr, column=col + 1).value
        if not (isinstance(v1, (int, float)) and abs(v1 - cost) < 0.005):
            mismatch.append(f"{get_column_letter(col)}{tr}")
        if not (isinstance(v2, (int, float)) and abs(v2 - gmv) < 0.005):
            mismatch.append(f"{get_column_letter(col + 1)}{tr}")
    add("数值逐条读回比对", not mismatch, "; ".join(mismatch[:5]))

    bad_roi = [f"{get_column_letter(c)}{r}" for c in (c + 2 for c in all_cost_cols)
               for r in range(FIRST_DATA_ROW, data_last + 1)
               if not _is_formula(ws.cell(row=r, column=c).value)]
    add("ROI 列公式完好", not bad_roi, f"非公式 {len(bad_roi)} 个")

    bad_total = [f"{get_column_letter(c)}{r}" for c in total_cols
                 for r in range(FIRST_DATA_ROW, data_last + 1)
                 if not _is_formula(ws.cell(row=r, column=c).value)]
    add("总计列公式完好", not bad_total, f"非公式 {len(bad_total)} 个")

    bad_row = [get_column_letter(c) + str(total_row) for c in range(2, int(ws.max_column or 1) + 1)
               if not _is_formula(ws.cell(row=total_row, column=c).value)]
    add("总计行公式完好", not bad_row, f"非公式 {len(bad_row)} 个")

    head1, cola1, total1 = _snapshot(ws, total_row)
    add("无越界写入（表头/日期列/总计行）", head0 == head1 and cola0 == cola1 and total0 == total1)

    actual_cost = round(sum(float(ws.cell(row=tr, column=col).value or 0)
                            for tr, col, _c, _g in written_cells), 2)
    actual_gmv = round(sum(float(ws.cell(row=tr, column=col + 1).value or 0)
                           for tr, col, _c, _g in written_cells), 2)
    expected_cost = round(sum(item[2] for item in written_cells), 2)
    expected_gmv = round(sum(item[3] for item in written_cells), 2)
    add("回填合计与写入数据一致",
        abs(actual_cost - expected_cost) < 0.01 and abs(actual_gmv - expected_gmv) < 0.01,
        f"消耗 {actual_cost} / 成交 {actual_gmv}")

    out_detail = ("；".join(out_of_range[:2]) if out_of_range else "无越界日期")
    if table_month is None:
        table_month_text = "未识别"
    elif table_month[0] is None:
        table_month_text = f"{table_month[1]:02d}月"
    else:
        table_month_text = f"{table_month[0]}-{table_month[1]:02d}"
    add("日期行与月份校验",
        (not out_of_range) and len(months) == 1 and month_ok,
        f"日期月份={','.join(months) or '-'}；模板月份={table_month_text}；{out_detail}")

    zeros = [f"{d} {s['name'] or s['aid']} 花费={c} 成交={g}"
             for _t, _c, c, g, d, s in plan if c == 0 or g == 0]
    add("至少写入 1 条",
        written > 0,
        f"写入 {written} 条；零值 {len(zeros)} 处" + (f"（{'; '.join(zeros[:3])}）" if zeros else ""))

    ok = all(c["ok"] for c in checks)
    saved = False
    locked = False
    if ok and save:
        try:
            wb.save(excel_path)
            saved = True
            logger(f"[merge] 已保存: {excel_path}")
            logger("[merge] 提示: 保存后用 WPS/Excel 打开一次触发公式重算")
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
        "planned": len(plan),
        "cells": len(written_cells) * 2,
        "total_cost": actual_cost,
        "total_gmv": actual_gmv,
        "skipped": skipped,
        "out_of_range": out_of_range,
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
    ap.add_argument("--allow-partial", action="store_true",
                    help="允许在导出存在失败账户时仍合并（默认拒绝）")
    ap.add_argument("--config", default=str(kx.CONFIG_PATH))
    args = ap.parse_args()

    cfg = kx.load_config(Path(args.config))
    cfg = kx.default_account_cfg(cfg)
    if args.start and args.end:
        start, end = args.start, args.end
    else:
        start = end = args.date or kx.default_date(cfg)
    ids = [x.strip() for x in args.accounts.split(",")] if args.accounts else None

    try:
        res = run_merge(cfg, start, end, ids,
                        excel_path=Path(args.file) if args.file else None,
                        export_dir=Path(args.export_dir) if args.export_dir else None,
                        backup=not args.dry_run, save=not args.dry_run,
                        allow_partial=args.allow_partial)
    except Exception as e:  # noqa: BLE001
        print(f"[merge] 失败: {type(e).__name__}: {e}")
        return 1
    print("\n===== 校验清单 =====")
    for i, c in enumerate(res["checks"], 1):
        print(f"{i}. [{'PASS' if c['ok'] else 'FAIL'}] {c['name']}" + (f"  {c['detail']}" if c["detail"] else ""))
    if res.get("out_of_range"):
        print("[merge] 无法写入的日期: " + "；".join(res["out_of_range"][:5]))
    if res.get("skipped"):
        print(f"[merge] 跳过 {len(res['skipped'])} 条（无数据/已留空）")
    print(f"===== {'已保存' if res['saved'] else ('校验通过（未保存）' if args.dry_run else '未保存')} =====")
    return 0 if all(c["ok"] for c in res["checks"]) else 1


if __name__ == "__main__":
    sys.exit(main())
