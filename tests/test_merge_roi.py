# -*- coding: utf-8 -*-
"""merge_roi 回归测试：周期追加、退款率表插列、旧总计行修复、标签自动生成。"""
from __future__ import annotations

import csv
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook, load_workbook

import kx_core as kx
import merge_roi

PIDS = [
    ("21384774414757", "365夜睡前故事"),
    ("24023046583757", "高情商故事绘本"),
    ("21026448882757", "小羊上山"),
    ("20679211961757", "宝宝学说话"),
    ("25004018621757", "福尔摩斯"),
    ("24622328577757", "民国课本"),
    ("21449382014757", "爆笑作文"),
    ("26897078104757", "小巴掌"),
    ("26874963835757", "大侦探"),
]


class RoiTestBase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self._old_data = kx.DATA_DIR
        kx.DATA_DIR = self.root / "data"
        kx.DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.book = self.root / "核算表.xlsx"
        self.week_csv = self.root / "week.csv"
        self.prev_csv = self.root / "prev.csv"
        self.syt_week = self.root / "syt_week.xlsx"
        self.syt_range = self.root / "syt_range.xlsx"
        self.cfg = {"roi_book_path": str(self.book), "auto_backup": False}

    def tearDown(self):
        kx.DATA_DIR = self._old_data
        self.tmp.cleanup()

    def _write_csv(self, path: Path, cost: float, gmv: float) -> None:
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["date", "account_id", "account_name", "product_id", "product_name",
                        "cost", "gmv", "order_cnt", "roi", "show_cnt", "click_cnt"])
            for pid, name in PIDS:
                w.writerow(["2026-11-02", "100", "acc", pid, name,
                            cost, gmv, 0, 0, 0, 0])

    def _write_syt(self, path: Path, rate: float) -> None:
        wb = Workbook()
        ws = wb.active
        ws.append(["商品ID", "商品名称", "成交件数", "成交退款率(订单)", "成交金额"])
        for pid, name in PIDS:
            ws.append([pid, name, 10, f"{rate * 100:.2f}%", 1000])
        wb.save(path)
        wb.close()

    def build_book(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "退款率"
        ws.append(["商品ID", "商品名", "自营保本roi", "5%盈利roi", "核算退款率",
                   "核算退款率(上周期)", "7日销售额（10.26-11.1）", "7日退款率（10.26-11.1）",
                   "30日销售额（10.3-11.1）", "30日退款率（10.3-11.1）"])
        for i, (pid, name) in enumerate(PIDS):
            r = 2 + i
            ws.cell(row=r, column=1, value=pid)
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=7, value=f"=盈利!E{2 + i}")
            ws.cell(row=r, column=8, value=0.16)
            ws.cell(row=r, column=9, value=f"=盈利!N{2 + i}")
            ws.cell(row=r, column=10, value=0.15)

        ws2 = wb.create_sheet("盈利")
        ws2.cell(row=1, column=1, value="周期")
        for i, (pid, name) in enumerate(PIDS):
            r = 2 + i
            ws2.cell(row=r, column=1, value="10.26-11.1")
            ws2.cell(row=r, column=2, value=pid)
            ws2.cell(row=r, column=3, value=name)
            ws2.cell(row=r, column=4, value=100)
            ws2.cell(row=r, column=5, value=200)
            ws2.cell(row=r, column=6, value=f"='1'!BA{4 + i}")
            ws2.cell(row=r, column=7, value=f"=E{r}/D{r}")
            ws2.cell(row=r, column=8, value=f"=E{r}/F{r}-D{r}")
            ws2.cell(row=r, column=10, value="10.3-11.1")
            ws2.cell(row=r, column=11, value=pid)
            ws2.cell(row=r, column=12, value=name)
            ws2.cell(row=r, column=13, value=300)
            ws2.cell(row=r, column=14, value=600)
            ws2.cell(row=r, column=15, value=f"='1'!BD{4 + i}")
            ws2.cell(row=r, column=16, value=f"=N{r}/M{r}")
            ws2.cell(row=r, column=17, value=f"=N{r}/O{r}-M{r}")
        # 旧总计行故意只加 2..9，漏最后一行，用来验证修复
        tr = 11
        ws2.cell(row=tr, column=1, value="10.26-11.1")
        ws2.cell(row=tr, column=3, value="总计")
        ws2.cell(row=tr, column=4, value="=SUM(D2:D9)")
        ws2.cell(row=tr, column=5, value="=SUM(E2:E9)")
        ws2.cell(row=tr, column=8, value="=SUM(H2:H9)")
        ws2.cell(row=tr, column=13, value="=SUM(M2:M9)")
        ws2.cell(row=tr, column=14, value="=SUM(N2:N9)")
        ws2.cell(row=tr, column=17, value="=SUM(Q2:Q9)")

        ws3 = wb.create_sheet("1")
        ws3["AY4"] = "=退款率!H2"
        ws3["BB4"] = "=退款率!J2"
        wb.save(self.book)
        wb.close()

    def prepare(self, week_cost=100, week_gmv=200, prev_cost=100, prev_gmv=200,
                week_rate=0.1356, range_rate=0.1604) -> None:
        self.build_book()
        self._write_csv(self.week_csv, week_cost, week_gmv)
        self._write_csv(self.prev_csv, prev_cost, prev_gmv)
        self._write_syt(self.syt_week, week_rate)
        self._write_syt(self.syt_range, range_rate)

    def run_it(self, **kw):
        return merge_roi.run_merge(
            self.cfg, "2026-11-02", "2026-11-08", "2026-10-10", "2026-11-08",
            book_path=self.book, week_csv=self.week_csv, prev_csv=self.prev_csv,
            syt_week=self.syt_week, syt_range=self.syt_range,
            backup=False, save=True, logger=lambda _m: None, **kw)


class MergeRoiTests(RoiTestBase):
    def test_append_period_and_insert_columns(self):
        self.prepare()
        res = self.run_it()
        self.assertTrue(all(c["ok"] for c in res["checks"]),
                        [c for c in res["checks"] if not c["ok"]])
        self.assertEqual(res["new_first"], 12)
        self.assertEqual(res["new_total"], 21)
        self.assertTrue(res["saved"])

        wb = load_workbook(self.book)
        ws_rate = wb["退款率"]
        ws_profit = wb["盈利"]
        self.assertEqual(ws_rate.max_column, 14)
        self.assertEqual(ws_rate["G1"].value, "7日销售额（11.2-11.8）")
        self.assertEqual(ws_rate["H1"].value, "7日退款率（11.2-11.8）")
        self.assertEqual(ws_rate["I1"].value, "30日销售额（10.10-11.8）")
        self.assertEqual(ws_rate["J1"].value, "30日退款率（10.10-11.8）")
        self.assertEqual(ws_rate["G2"].value, "=盈利!E12")
        self.assertAlmostEqual(ws_rate["H2"].value, 0.1356, places=4)
        self.assertAlmostEqual(ws_rate["J2"].value, 0.1604, places=4)
        # 旧块整体右移到 K:N
        self.assertEqual(ws_rate["K1"].value, "7日销售额（10.26-11.1）")
        self.assertEqual(ws_rate["K2"].value, "=盈利!E2")
        wb.close()

    def test_new_period_rows_and_formulas(self):
        self.prepare()
        self.run_it()
        wb = load_workbook(self.book)
        ws = wb["盈利"]
        self.assertEqual(ws["A12"].value, "11.2-11.8")
        self.assertEqual(ws["J12"].value, "10.10-11.8")
        self.assertEqual(ws["D12"].value, 100)
        self.assertEqual(ws["E12"].value, 200)
        self.assertEqual(ws["M12"].value, 200)
        self.assertEqual(ws["N12"].value, 400)
        self.assertEqual(ws["F12"].value, "='1'!BA4")
        self.assertEqual(ws["G12"].value, "=E12/D12")
        self.assertEqual(ws["H12"].value, "=E12/F12-D12")
        self.assertEqual(ws["O12"].value, "='1'!BD4")
        self.assertEqual(ws["P12"].value, "=N12/M12")
        self.assertEqual(ws["Q12"].value, "=N12/O12-M12")
        self.assertEqual(ws["C21"].value, "总计")
        self.assertEqual(ws["D21"].value, "=SUM(D12:D20)")
        self.assertEqual(ws["Q21"].value, "=SUM(Q12:Q20)")
        wb.close()

    def test_old_total_row_fixed(self):
        self.prepare()
        self.run_it()
        wb = load_workbook(self.book)
        ws = wb["盈利"]
        for col in ("D", "E", "H", "M", "N", "Q"):
            self.assertEqual(ws[f"{col}11"].value, f"=SUM({col}2:{col}10)")
        wb.close()

    def test_period_label_cross_month(self):
        self.assertEqual(merge_roi.date_label(date(2026, 9, 28), date(2026, 10, 4)),
                         "9.28-10.4")

    def test_duplicate_period_rejected(self):
        self.prepare()
        self.run_it()
        with self.assertRaises(RuntimeError) as ctx:
            self.run_it()
        self.assertIn("已存在周期", str(ctx.exception))

    def test_overwrite_existing_period(self):
        self.prepare()
        self.run_it()
        res = self.run_it(overwrite=True)
        self.assertTrue(res["saved"])
        self.assertTrue(res["overwritten"])
        self.assertEqual(res["new_first"], 12)
        wb = load_workbook(self.book)
        self.assertEqual(wb["盈利"].max_row, 21)   # 覆盖而不是追加
        self.assertEqual(wb["退款率"].max_column, 14)  # 不再插列
        wb.close()

    def test_missing_book_path_raises(self):
        self.cfg["roi_book_path"] = ""
        self.build_book()
        with self.assertRaises(RuntimeError):
            merge_roi.run_merge(
                {"roi_book_path": ""}, "2026-11-02", "2026-11-08",
                "2026-10-10", "2026-11-08",
                week_csv=self.week_csv, prev_csv=self.prev_csv,
                syt_week=self.syt_week, syt_range=self.syt_range,
                backup=False, save=False, logger=lambda _m: None)


if __name__ == "__main__":
    unittest.main()
