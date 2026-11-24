# -*- coding: utf-8 -*-
"""内置月度表模板：带格式生成整月工作表，跨月自动新建 sheet。

布局与历史月度表保持一致：
  行1 商品名 / 行2 商品ID / 行3 账户ID / 行4 表头 / 行5 起为日期
  商品组按「消耗 / 成交 / ROI」三列一组，模板预留 1 个空组作为格式母版，
  首次导入商品时由 merge_excel 按该组样式自动扩展。
  末尾 3 列为总计列，日期下方一行为总计行。
"""
from __future__ import annotations

import calendar
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BOOK_NAME = "快手消耗成交.xlsx"
SHEET_PREFIX = "消耗成交"
GROUPS_START = 2
GROUP_WIDTH = 3
FIRST_DATA_ROW = 5
HEADER_ROW = 4

FONT_NAME = "微软雅黑"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
_THIN = Side(style="thin", color="FFB0B0B0")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
CENTER = Alignment(horizontal="center", vertical="center")
ROI_FORMAT = "0.00_"


def sheet_name(year: int, month: int) -> str:
    return f"{year}年{month}月"


def detect_month(ws) -> tuple[int | None, int] | None:
    """从 A 列（第 5 行起）识别工作表所属年月"""
    for r in range(FIRST_DATA_ROW, int(ws.max_row or 0) + 1):
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
        m = re.match(r"^(\d{1,2})[-/.](\d{1,2})$", s)
        if m:
            return None, int(m.group(1))
    return None


def pick_sheet_name(wb) -> str:
    """取最新月份的月度 sheet；兼容「2026年11月」与旧「消耗成交-2611」两种命名"""
    names = [n for n in wb.sheetnames if re.match(r"^\d{4}年\d{1,2}月$", n)
             or n == SHEET_PREFIX or n.startswith(SHEET_PREFIX + "-")]
    if not names:
        return wb.sheetnames[0]

    def key(n: str):
        m = re.match(r"^(\d{4})年(\d{1,2})月$", n)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        m = re.search(r"-(\d{2})(\d{2})$", n)
        return (2000 + int(m.group(1)), int(m.group(2))) if m else (0, 0)

    return max(names, key=key)


def create_sheet(wb, year: int, month: int):
    """新建一个带格式的整月工作表（已存在则直接返回）"""
    name = sheet_name(year, month)
    if name in wb.sheetnames:
        return wb[name]

    ws = wb.create_sheet(name)
    days = calendar.monthrange(year, month)[1]
    total_row = FIRST_DATA_ROW + days
    total_col = GROUPS_START + GROUP_WIDTH  # 商品组 B:D -> 总计列 E
    last_col = total_col + 2

    # 表头文字
    ws.cell(row=1, column=1, value="商品名")
    ws.cell(row=2, column=1, value="商品ID")
    ws.cell(row=3, column=1, value="账户ID")
    ws.cell(row=HEADER_ROW, column=1, value="日期")
    for i, label in enumerate(("消耗", "成交", "ROI")):
        ws.cell(row=HEADER_ROW, column=GROUPS_START + i, value=label)
        ws.cell(row=HEADER_ROW, column=total_col + i, value=label)
    ws.cell(row=1, column=total_col, value="总计")

    # 商品组 / 总计组 1-3 行合并
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=GROUPS_START,
                       end_row=r, end_column=GROUPS_START + GROUP_WIDTH - 1)
        ws.merge_cells(start_row=r, start_column=total_col,
                       end_row=r, end_column=last_col)

    # 日期与整行公式
    for i in range(days):
        r = FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=date(year, month, i + 1))
        b, c, d = get_column_letter(GROUPS_START), get_column_letter(GROUPS_START + 1), get_column_letter(GROUPS_START + 2)
        e, f, g = get_column_letter(total_col), get_column_letter(total_col + 1), get_column_letter(total_col + 2)
        ws.cell(row=r, column=total_col, value=f"={b}{r}")
        ws.cell(row=r, column=total_col + 1, value=f"={c}{r}")
        ws.cell(row=r, column=total_col + 2, value=f"={f}{r}/{e}{r}")

    # 总计行
    last = total_row - 1
    ws.cell(row=total_row, column=1, value="总计")
    for col in (GROUPS_START, GROUPS_START + 1, total_col, total_col + 1):
        L = get_column_letter(col)
        ws.cell(row=total_row, column=col, value=f"=SUM({L}{FIRST_DATA_ROW}:{L}{last})")
    for col in (GROUPS_START + 2, total_col + 2):
        L = get_column_letter(col)
        prev = get_column_letter(col - 1)
        prev2 = get_column_letter(col - 2)
        ws.cell(row=total_row, column=col, value=f"={prev}{total_row}/{prev2}{total_row}")

    # 列宽 / 行高
    widths = {"A": 9.0, "B": 9.5, "C": 13.0, "D": 13.0, "E": 9.0, "F": 13.0, "G": 13.0}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 26.0
    ws.row_dimensions[2].height = 18.0
    ws.row_dimensions[3].height = 18.0
    ws.row_dimensions[HEADER_ROW].height = 20.0
    for r in range(FIRST_DATA_ROW, total_row + 1):
        ws.row_dimensions[r].height = 16.5

    # 字体 / 边框 / 填充
    for r in range(1, total_row + 1):
        for c in range(1, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = CENTER
            if r <= HEADER_ROW:
                cell.fill = HEADER_FILL
                cell.font = Font(name=FONT_NAME, size=10 if c == 1 else 9, bold=True)
            else:
                cell.font = Font(name=FONT_NAME, size=10)
    for r in range(FIRST_DATA_ROW, total_row + 1):
        for c in (GROUPS_START + 2, total_col + 2):
            ws.cell(row=r, column=c).number_format = ROI_FORMAT
    for i in range(days):
        ws.cell(row=FIRST_DATA_ROW + i, column=1).number_format = "yyyy-mm-dd"

    ws.freeze_panes = "B5"
    return ws


def ensure_month_sheet(wb, year: int, month: int) -> tuple[object, bool]:
    """取指定月份 sheet，不存在则按模板新建；返回 (worksheet, 是否新建)"""
    name = sheet_name(year, month)
    if name in wb.sheetnames:
        return wb[name], False
    # 兼容旧版固定 sheet：月份一致时直接重命名，保留历史数据
    if SHEET_PREFIX in wb.sheetnames:
        ym = detect_month(wb[SHEET_PREFIX])
        if ym and ym[0] == year and ym[1] == month:
            wb[SHEET_PREFIX].title = name
            return wb[name], False
    return create_sheet(wb, year, month), True


def _rename_legacy_sheets(path: Path, logger=print) -> None:
    """把旧版 sheet 名（消耗成交 / 消耗成交-YYMM）统一改成「YYYY年M月」"""
    try:
        wb = load_workbook(path)
    except Exception:  # noqa: BLE001
        return
    changed = False
    try:
        names = list(wb.sheetnames)
        for name in names:
            target = None
            if name == SHEET_PREFIX:
                ym = detect_month(wb[name])
                if ym and ym[0]:
                    target = sheet_name(ym[0], ym[1])
            else:
                m = re.match(r"^" + re.escape(SHEET_PREFIX) + r"-(\d{2})(\d{2})$", name)
                if m:
                    target = sheet_name(2000 + int(m.group(1)), int(m.group(2)))
            if target and target != name and target not in wb.sheetnames:
                wb[name].title = target
                changed = True
                logger(f"[template] 旧 sheet「{name}」已重命名为「{target}」")
        if changed:
            wb.save(path)
    finally:
        wb.close()


def ensure_book(path: Path, legacy: Path | None = None, logger=print) -> bool:
    """确保月度表文件存在；不存在时用内置模板创建，或从旧月度表迁移。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        _rename_legacy_sheets(path, logger)
        return False

    legacy = Path(legacy) if legacy else None
    if legacy and legacy.exists() and legacy.resolve() != path.resolve():
        shutil.copy2(legacy, path)
        wb = load_workbook(path)
        try:
            names = list(wb.sheetnames)
            target = next((n for n in names if n == SHEET_PREFIX or n.startswith(SHEET_PREFIX + "-")), names[0])
            ym = detect_month(wb[target])
            if ym and ym[0]:
                new_name = sheet_name(ym[0], ym[1])
                if target != new_name and new_name not in wb.sheetnames:
                    wb[target].title = new_name
            wb.save(path)
        finally:
            wb.close()
        logger(f"[template] 已从旧月度表迁移: {legacy} -> {path}")
        return True

    wb = Workbook()
    wb.remove(wb.active)
    create_sheet(wb, date.today().year, date.today().month)
    wb.save(path)
    wb.close()
    logger(f"[template] 已用内置模板创建月度表: {path}")
    return True
