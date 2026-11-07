# -*- coding: utf-8 -*-
"""fetch_products 分页回归：totalCount 缺失时不能只取第一页。"""
from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import kx_core as kx


def _item(i: int) -> dict:
    return {
        "merchantProductId": f"P{i}",
        "itemName": f"商品{i}",
        "costTotal": 1000,
        "t0OrderPaymentAmt": 2000,
        "t0OrderCnt": 1,
        "t0Roi": 2.0,
        "showCnt": 10,
        "itemEntranceClkCnt": 1,
    }


class FakePage:
    def __init__(self, pages: dict[int, list[dict]], total_count: int = 0):
        self.pages = pages
        self.total_count = total_count
        self.calls: list[int] = []

    def evaluate(self, _script, arg):
        payload, _api = arg
        page_no = int(payload["pageInfo"]["currentPage"])
        self.calls.append(page_no)
        return {"ok": True, "data": {
            "detail": self.pages.get(page_no, []),
            "totalCount": self.total_count,
        }}

    def wait_for_timeout(self, _ms):
        pass


class PaginationTests(unittest.TestCase):
    def fetch(self, page):
        client = kx.Client({"page_size": 100})
        return client.fetch_products(page, "10000001", "账户", datetime(2026, 11, 3))

    def test_missing_total_count_pages_until_empty(self):
        page = FakePage({1: [_item(i) for i in range(100)],
                         2: [_item(i) for i in range(100, 200)],
                         3: []}, total_count=0)
        rows = self.fetch(page)
        self.assertEqual(len(rows), 200)
        self.assertEqual(page.calls, [1, 2, 3])

    def test_total_count_stops_early(self):
        page = FakePage({1: [_item(i) for i in range(100)],
                         2: [_item(i) for i in range(100, 150)],
                         3: []}, total_count=150)
        rows = self.fetch(page)
        self.assertEqual(len(rows), 150)
        self.assertEqual(page.calls, [1, 2])

    def test_missing_total_count_single_page(self):
        page = FakePage({1: [_item(i) for i in range(10)], 2: []}, total_count=0)
        rows = self.fetch(page)
        self.assertEqual(len(rows), 10)
        self.assertEqual(page.calls, [1, 2])


if __name__ == "__main__":
    unittest.main()
