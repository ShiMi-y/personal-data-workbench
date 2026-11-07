# -*- coding: utf-8 -*-
"""merge_excel 回归测试：用合成月度表覆盖正常/31 日/跨月/失败账户/新增列/商品回退。"""
from __future__ import annotations

import csv
import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import kx_core as kx
import merge_excel


class MergeTestBase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self._old_data_dir = kx.DATA_DIR
        kx.DATA_DIR = self.root / "data"
        kx.DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.book = self.root / "月度表.xlsx"
        self.export = self.root / "export"
        self.export.mkdir()
        self.cfg = {"excel_path": str(self.book), "auto_backup": False, "backup_limit": 10}

    def tearDown(self):
        kx.DATA_DIR = self._old_data_dir
        self.tmp.cleanup()

    def build_book(self, groups, year=2026, month=11, days=30):
        """构造标准月度表：数据行 5..(4+days)，总计行 4+days+1。"""
        total_row = 4 + days + 1
        wb = Workbook()
        ws = wb.active
        ws.title = merge_excel.SHEET
        ws.cell(row=1, column=1, value="日期")
        col = 2
        for name, pid, aid in groups:
            ws.cell(row=1, column=col, value=name)
            ws.cell(row=2, column=col, value=pid)
            ws.cell(row=3, column=col, value=aid)
            ws.cell(row=4, column=col, value="消耗")
            ws.cell(row=4, column=col + 1, value="成交")
            ws.cell(row=4, column=col + 2, value="ROI")
            for r in (1, 2, 3):
                ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 2)
            col += 3
        cost_cols = list(range(2, col, 3))
        total_start = col
        ws.cell(row=1, column=total_start, value="总计")
        ws.cell(row=4, column=total_start, value="消耗")
        ws.cell(row=4, column=total_start + 1, value="成交")
        ws.cell(row=4, column=total_start + 2, value="ROI")
        for r in (1, 2, 3):
            ws.merge_cells(start_row=r, start_column=total_start,
                           end_row=r, end_column=total_start + 2)

        for i in range(days):
            r = merge_excel.FIRST_DATA_ROW + i
            ws.cell(row=r, column=1, value=datetime(year, month, i + 1))
            for c in cost_cols:
                ws.cell(row=r, column=c + 2,
                        value=f"={get_column_letter(c + 1)}{r}/{get_column_letter(c)}{r}")
            ws.cell(row=r, column=total_start,
                    value="=" + "+".join(f"{get_column_letter(c)}{r}" for c in cost_cols))
            ws.cell(row=r, column=total_start + 1,
                    value="=" + "+".join(f"{get_column_letter(c + 1)}{r}" for c in cost_cols))
            ws.cell(row=r, column=total_start + 2,
                    value=f"={get_column_letter(total_start + 1)}{r}/"
                          f"{get_column_letter(total_start)}{r}")
        ws.cell(row=total_row, column=1, value="总计")
        for c in range(2, total_start + 3):
            ws.cell(row=total_row, column=c,
                    value=f"=SUM({get_column_letter(c)}{merge_excel.FIRST_DATA_ROW}:"
                          f"{get_column_letter(c)}{total_row - 1})")
        wb.save(self.book)

    def write_export(self, accounts, products, failures=None):
        with (self.export / "accounts.csv").open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["date", "account_id", "account_name", "cost", "gmv",
                        "order_cnt", "product_count"])
            for a in accounts:
                w.writerow([a["date"], a["account_id"], a.get("account_name", ""),
                            a.get("cost", 0), a.get("gmv", 0), a.get("order_cnt", 0),
                            a.get("product_count", 1)])
        with (self.export / "products.csv").open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["date", "account_id", "account_name", "product_id", "product_name",
                        "cost", "gmv", "order_cnt", "roi", "show_cnt", "click_cnt"])
            for p in products:
                w.writerow([p["date"], p["account_id"], p.get("account_name", ""),
                            p.get("product_id", ""), p.get("product_name", ""),
                            p.get("cost", 0), p.get("gmv", 0), p.get("order_cnt", 0),
                            p.get("roi", 0), p.get("show_cnt", 0), p.get("click_cnt", 0)])
        (self.export / "summary.json").write_text(
            json.dumps({"failures": failures or []}, ensure_ascii=False), encoding="utf-8")

    @staticmethod
    def prod(date, aid, pid, name, cost, gmv):
        return {"date": date, "account_id": aid, "account_name": "账户",
                "product_id": pid, "product_name": name,
                "cost": cost, "gmv": gmv, "order_cnt": 1, "roi": 0,
                "show_cnt": 0, "click_cnt": 0}

    def merge(self, start, end, ids=None, save=True, allow_partial=False):
        return merge_excel.run_merge(
            self.cfg, start, end, ids, excel_path=self.book, export_dir=self.export,
            backup=False, save=save, allow_partial=allow_partial)


class MergeTests(MergeTestBase):
    def test_normal_merge_writes_correct_row(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "account_name": "账户一",
              "cost": 123.45, "gmv": 456.78, "order_cnt": 7, "product_count": 2}],
            [self.prod("2026-11-03", "10000001", "P1001", "商品A", 100, 400)])
        res = self.merge("2026-11-03", "2026-11-03")
        self.assertTrue(res["saved"])
        self.assertEqual(res["written"], 1)
        self.assertTrue(all(c["ok"] for c in res["checks"]), res["checks"])
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["B7"].value, 100)   # 3 日 -> 第 7 行
        self.assertEqual(ws["C7"].value, 400)

    def test_day31_is_rejected_on_30_day_template(self):
        self.build_book([("商品A", "P1001", "10000001")], month=10, days=30)
        self.write_export(
            [{"date": "2026-10-31", "account_id": "10000001", "cost": 5, "gmv": 5}],
            [self.prod("2026-10-31", "10000001", "P1001", "商品A", 5, 5)])
        before = self.book.read_bytes()
        res = self.merge("2026-10-31", "2026-10-31")
        self.assertFalse(res["saved"])
        self.assertEqual(res["written"], 0)
        self.assertTrue(res["out_of_range"])
        check = next(c for c in res["checks"] if c["name"] == "日期行与月份校验")
        self.assertFalse(check["ok"])
        self.assertEqual(self.book.read_bytes(), before)

    def test_day31_is_supported_on_31_day_template(self):
        self.build_book([("商品A", "P1001", "10000001")], month=10, days=31)
        self.write_export(
            [{"date": "2026-10-31", "account_id": "10000001", "cost": 5, "gmv": 5}],
            [self.prod("2026-10-31", "10000001", "P1001", "商品A", 5, 5)])
        res = self.merge("2026-10-31", "2026-10-31")
        self.assertTrue(res["saved"])
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["B35"].value, 5)   # 31 日 -> 第 35 行
        self.assertEqual(ws["B36"].value, "=SUM(B5:B35)")

    def test_cross_month_range_is_rejected(self):
        self.build_book([("商品A", "P1001", "10000001")], month=10, days=30)
        self.write_export(
            [{"date": "2026-10-30", "account_id": "10000001", "cost": 1, "gmv": 1},
             {"date": "2026-11-01", "account_id": "10000001", "cost": 2, "gmv": 2}],
            [])
        with self.assertRaisesRegex(RuntimeError, "跨月"):
            self.merge("2026-10-30", "2026-11-01")

    def test_table_month_mismatch_is_rejected(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-10-15", "account_id": "10000001", "cost": 1, "gmv": 1}], [])
        with self.assertRaisesRegex(RuntimeError, "月份"):
            self.merge("2026-10-15", "2026-10-15")

    def test_export_failures_block_merge_unless_allowed(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "cost": 1, "gmv": 1}],
            [self.prod("2026-11-03", "10000001", "P1001", "商品A", 1, 1)],
            failures=[{"account_id": "10000002", "error": "boom"}])
        with self.assertRaisesRegex(RuntimeError, "失败账户"):
            self.merge("2026-11-03", "2026-11-03")
        res = self.merge("2026-11-03", "2026-11-03", allow_partial=True)
        self.assertTrue(res["saved"])

    def test_new_account_adds_group_before_total(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "cost": 100, "gmv": 400},
             {"date": "2026-11-03", "account_id": "10000009", "cost": 88.8, "gmv": 188.8}],
            [self.prod("2026-11-03", "10000001", "P1001", "商品A", 100, 400),
             self.prod("2026-11-03", "10000009", "P2001", "新商品", 88.8, 188.8)])
        res = self.merge("2026-11-03", "2026-11-03", ["10000001", "10000009"])
        self.assertTrue(res["saved"])
        self.assertEqual(len(res["added_columns"]), 1)
        self.assertEqual(res["added_columns"][0]["col"], "E")
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["E1"].value, "新商品")
        self.assertEqual(ws["E2"].value, "P2001")
        self.assertEqual(ws["E3"].value, "10000009")
        self.assertEqual(ws["H1"].value, "总计")

    def test_missing_product_does_not_fall_back_to_account_total(self):
        # 同一账户两个商品列：P1001 有数据，P1002 不在报表；
        # 账户总额 123.45 不能被回填到 P1002，否则总计列会重复相加。
        self.build_book([("商品A", "P1001", "10000001"),
                         ("商品B", "P1002", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "cost": 123.45, "gmv": 456.78}],
            [self.prod("2026-11-03", "10000001", "P1001", "商品A", 100, 400),
             self.prod("2026-11-03", "10000001", "P9999", "其他商品", 23.45, 56.78)])
        res = self.merge("2026-11-03", "2026-11-03")
        self.assertTrue(res["saved"])
        self.assertEqual(res["written"], 1)
        self.assertTrue(any("已留空" in x for x in res["skipped"]))
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["B7"].value, 100)
        self.assertIsNone(ws["E7"].value)

    def test_read_product_groups(self):
        self.build_book([("商品A", "P1001", "10000001"),
                         ("商品B", "P1002", "10000002")], month=11, days=30)
        groups = merge_excel.read_product_groups(self.book)
        self.assertEqual(len(groups), 2)
        self.assertEqual(groups[0]["product_id"], "P1001")
        self.assertEqual(groups[1]["account_id"], "10000002")

    def test_new_account_on_31_day_template(self):
        self.build_book([("商品A", "P1001", "10000001")], month=10, days=31)
        self.write_export(
            [{"date": "2026-10-31", "account_id": "10000001", "cost": 5, "gmv": 5},
             {"date": "2026-10-31", "account_id": "10000009", "cost": 8, "gmv": 16}],
            [self.prod("2026-10-31", "10000001", "P1001", "商品A", 5, 5),
             self.prod("2026-10-31", "10000009", "P2001", "新商品", 8, 16)])
        res = self.merge("2026-10-31", "2026-10-31", ["10000001", "10000009"])
        self.assertTrue(res["saved"])
        self.assertEqual(len(res["added_columns"]), 1)
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["B35"].value, 5)
        self.assertEqual(ws["E35"].value, 8)
        self.assertEqual(ws["H36"].value, "=SUM(H5:H35)")


class SerialDateMonthTests(MergeTestBase):
    """真实月度表 A 列是 Excel 日期序列号（46327…），月份必须同样能识别。"""

    def use_serial_dates(self, year=2026, month=11, days=30):
        wb = load_workbook(self.book)
        ws = wb[merge_excel.SHEET]
        base = datetime(1899, 12, 30)
        for i in range(days):
            ws.cell(row=merge_excel.FIRST_DATA_ROW + i, column=1,
                    value=(datetime(year, month, i + 1) - base).days)
        wb.save(self.book)

    def test_serial_dates_month_recognized(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.use_serial_dates()
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "cost": 10, "gmv": 20}],
            [self.prod("2026-11-03", "10000001", "P1001", "商品A", 10, 20)])
        res = self.merge("2026-11-03", "2026-11-03")
        self.assertTrue(res["saved"])
        detail = [c["detail"] for c in res["checks"] if c["name"] == "日期行与月份校验"][0]
        self.assertIn("模板月份=2026-11", detail)

    def test_serial_dates_reject_other_month(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.use_serial_dates()
        self.write_export(
            [{"date": "2026-10-03", "account_id": "10000001", "cost": 1, "gmv": 2}],
            [self.prod("2026-10-03", "10000001", "P1001", "商品A", 1, 2)])
        with self.assertRaisesRegex(RuntimeError, "不一致"):
            self.merge("2026-10-03", "2026-10-03")


class NewColumnNamingTests(MergeTestBase):
    """新增商品列的第 1 行用账户名称，商品名只作兜底。"""

    def test_new_column_header_uses_account_name(self):
        self.build_book([("商品A", "P1001", "10000001")], month=11, days=30)
        self.write_export(
            [{"date": "2026-11-03", "account_id": "10000001", "account_name": "老账户",
              "cost": 10, "gmv": 20},
             {"date": "2026-11-03", "account_id": "10000009", "account_name": "新账户名",
              "cost": 30, "gmv": 60}],
            [{"date": "2026-11-03", "account_id": "10000001", "account_name": "老账户",
              "product_id": "P1001", "product_name": "商品A", "cost": 10, "gmv": 20,
              "order_cnt": 1, "roi": 0, "show_cnt": 0, "click_cnt": 0},
             {"date": "2026-11-03", "account_id": "10000009", "account_name": "新账户名",
              "product_id": "P2001", "product_name": "新商品", "cost": 30, "gmv": 60,
              "order_cnt": 1, "roi": 0, "show_cnt": 0, "click_cnt": 0}])
        res = self.merge("2026-11-03", "2026-11-03", ["10000001", "10000009"])
        self.assertTrue(res["saved"])
        ws = load_workbook(self.book)[merge_excel.SHEET]
        self.assertEqual(ws["E1"].value, "新账户名")
        self.assertEqual(ws["E2"].value, "P2001")
        self.assertEqual(ws["E3"].value, "10000009")


if __name__ == "__main__":
    unittest.main()
