# -*- coding: utf-8 -*-
"""realtime_history 单测：留痕追加、小时/当日取最后值、日(24h)/周(7天)/月(31天) 与本期口径。"""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import realtime_history as rh

NOW = datetime(2026, 11, 24, 18, 30, 0)


def snap(day: str, ts: str, rows) -> dict:
    return {
        "date": day,
        "fetched_at": f"{day} {ts}",
        "elapsed": 12.5,
        "accounts": [
            {"owner_id": "1", "owner_label": "主店", "account_id": aid,
             "account_name": f"账户{aid}", "cost": cost, "gmv": gmv, "order_cnt": orders}
            for aid, cost, gmv, orders in rows
        ],
        "failures": [],
    }


class RealtimeHistoryTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name) / "data"
        self.root.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        self.tmp.cleanup()

    # ---------------- 写入 ----------------
    def test_record_and_daily_last(self):
        self.assertTrue(rh.record(self.root, snap("2026-11-24", "09:05:00", [("1001", 10, 30, 1)])))
        self.assertTrue(rh.record(self.root, snap("2026-11-24", "15:30:00", [("1001", 26, 78, 4)])))
        daily = rh.daily_last(self.root)
        self.assertEqual(list(daily), ["2026-11-24"])
        self.assertEqual(daily["2026-11-24"]["1|1001"]["cost"], 26)
        self.assertEqual(daily["2026-11-24"]["1|1001"]["gmv"], 78)

    def test_record_skips_empty_rows(self):
        self.assertFalse(rh.record(self.root, {"date": "2026-11-24", "accounts": []}))
        self.assertEqual(rh.daily_last(self.root), {})

    def test_hourly_last(self):
        rh.record(self.root, snap("2026-11-24", "09:05:00", [("1001", 10, 30, 1)]))
        rh.record(self.root, snap("2026-11-24", "09:35:00", [("1001", 18, 50, 2)]))
        rh.record(self.root, snap("2026-11-24", "10:05:00", [("1001", 25, 70, 3)]))
        out = rh.hourly_last(self.root, "2026-11-24")
        self.assertEqual(sorted(out), [9, 10])
        self.assertEqual(out[9]["1|1001"]["cost"], 18)
        self.assertEqual(out[10]["1|1001"]["cost"], 25)

    # ---------------- 日视图 ----------------
    def test_day_recent_24h(self):
        for h in (9, 12, 18):
            rh.record(self.root, snap("2026-11-24", f"{h:02d}:10:00", [("1001", h * 10, 0, 0)]))
        out = rh.series(self.root, "day", "cost", "recent", now=NOW)
        self.assertEqual(len(out["buckets"]), 24)
        self.assertEqual(out["buckets"][-1]["label"], "18:00")
        self.assertEqual(out["buckets"][0]["label"], "19:00")  # 前一天 19 点
        vals = out["series"][0]["values"]
        self.assertEqual(vals[-1], 180.0)          # 18 点最后一次快照
        self.assertEqual(vals[-7], 120.0)          # 12 点
        self.assertEqual(vals[-6], 120.0)          # 13 点沿用 12 点累计值
        self.assertEqual(vals[-10], 90.0)          # 09 点
        self.assertIsNone(vals[-12])               # 07:00 当天尚无快照
        self.assertEqual(out["scope_label"], "最近24小时")

    def test_day_period(self):
        rh.record(self.root, snap("2026-11-24", "09:10:00", [("1001", 90, 0, 0)]))
        rh.record(self.root, snap("2026-11-23", "09:10:00", [("1001", 55, 0, 0)]))  # 前一天不计入
        out = rh.series(self.root, "day", "cost", "period", now=NOW)
        self.assertEqual(len(out["buckets"]), 24)
        self.assertEqual(out["buckets"][0]["label"], "00:00")
        vals = out["series"][0]["values"]
        self.assertEqual(vals[9], 90.0)
        self.assertEqual(vals[10], 90.0)   # 无快照的小时沿用当日累计值
        self.assertIsNone(vals[0])
        self.assertEqual(out["scope_label"], "本日")

    # ---------------- 周视图 ----------------
    def test_day_carry_resets_across_days(self):
        rh.record(self.root, snap("2026-11-23", "23:10:00", [("1001", 300, 0, 0)]))
        rh.record(self.root, snap("2026-11-24", "09:10:00", [("1001", 20, 0, 0)]))
        out = rh.series(self.root, "day", "cost", "recent", now=NOW)
        labels = [b["label"] for b in out["buckets"]]
        vals = out["series"][0]["values"]
        i2300 = labels.index("23:00")
        self.assertEqual(vals[i2300], 300.0)       # 前一天 23 点
        self.assertIsNone(vals[i2300 + 1])         # 00:00 新的一天，重新累计
        self.assertEqual(vals[labels.index("09:00")], 20.0)

    def test_week_recent_7days(self):
        days = ["2026-11-18", "2026-11-23", "2026-11-24"]
        for d in days:
            rh.record(self.root, snap(d, "18:00:00", [("1001", 10, 5, 0)]))
        out = rh.series(self.root, "week", "cost", "recent", now=NOW)
        self.assertEqual(len(out["buckets"]), 7)
        self.assertEqual([b["label"] for b in out["buckets"]],
                         ["11-18", "11-19", "11-20", "11-21", "11-22", "11-23", "11-24"])
        self.assertEqual(out["series"][0]["values"], [10.0, None, None, None, None, 10.0, 10.0])

    def test_week_period(self):
        # 2026-11-24 是周二，本周应为 11-23(周一) ~ 11-29(周日)
        rh.record(self.root, snap("2026-11-22", "18:00:00", [("1001", 99, 0, 0)]))  # 上周日
        rh.record(self.root, snap("2026-11-23", "18:00:00", [("1001", 11, 0, 0)]))
        out = rh.series(self.root, "week", "cost", "period", now=NOW)
        self.assertEqual([b["label"] for b in out["buckets"]],
                         ["11-23", "11-24", "11-25", "11-26", "11-27", "11-28", "11-29"])
        self.assertEqual(out["series"][0]["values"], [11.0, None, None, None, None, None, None])
        self.assertEqual(out["scope_label"], "本周")

    # ---------------- 月视图 ----------------
    def test_month_recent_31days(self):
        rh.record(self.root, snap("2026-11-24", "18:00:00", [("1001", 12, 0, 0)]))
        out = rh.series(self.root, "month", "cost", "recent", now=NOW)
        self.assertEqual(len(out["buckets"]), 31)
        self.assertEqual(out["buckets"][-1]["label"], "11-24")
        self.assertEqual(out["buckets"][0]["label"], "10-25")
        self.assertEqual(out["series"][0]["values"][-1], 12.0)

    def test_month_period(self):
        rh.record(self.root, snap("2026-11-01", "18:00:00", [("1001", 5, 0, 0)]))
        out = rh.series(self.root, "month", "cost", "period", now=NOW)
        self.assertEqual(len(out["buckets"]), 30)   # 11 月 30 天
        self.assertEqual(out["buckets"][0]["label"], "11-01")
        self.assertEqual(out["series"][0]["values"][0], 5.0)
        self.assertEqual(out["scope_label"], "本月")

    # ---------------- 其它 ----------------
    def test_series_metric_gmv(self):
        rh.record(self.root, snap("2026-11-24", "18:00:00", [("1001", 10, 35, 2)]))
        out = rh.series(self.root, "day", "gmv", "recent", now=NOW)
        self.assertEqual(out["series"][0]["values"][-1], 35.0)

    def test_day_anchor(self):
        rh.record(self.root, snap("2026-11-20", "14:10:00", [("1001", 42, 0, 0)]))
        out = rh.series(self.root, "day", "cost", "period", anchor="2026-11-20", now=NOW)
        self.assertEqual(len(out["buckets"]), 24)
        self.assertEqual(out["buckets"][0]["label"], "00:00")
        self.assertEqual(out["series"][0]["values"][14], 42.0)
        self.assertEqual(out["scope_label"], "2026-11-20 全天")

    def test_week_anchor_starts_monday(self):
        rh.record(self.root, snap("2026-11-16", "18:00:00", [("1001", 7, 0, 0)]))
        # 传周四，应回落到该周周一 11-16
        out = rh.series(self.root, "week", "cost", "period", anchor="2026-11-19", now=NOW)
        self.assertEqual([b["label"] for b in out["buckets"]],
                         ["11-16", "11-17", "11-18", "11-19", "11-20", "11-21", "11-22"])
        self.assertEqual(out["series"][0]["values"][0], 7.0)
        self.assertEqual(out["scope_label"], "2026-11-16 ~ 2026-11-22")

    def test_month_anchor(self):
        rh.record(self.root, snap("2026-10-31", "18:00:00", [("1001", 9, 0, 0)]))
        out = rh.series(self.root, "month", "cost", "period", anchor="2026-10", now=NOW)
        self.assertEqual(len(out["buckets"]), 31)
        self.assertEqual(out["buckets"][-1]["label"], "10-31")
        self.assertEqual(out["series"][0]["values"][-1], 9.0)
        self.assertEqual(out["scope_label"], "2026-10 全月")

    def test_periods_lists(self):
        rh.record(self.root, snap("2026-11-23", "18:00:00", [("1001", 1, 0, 0)]))
        rh.record(self.root, snap("2026-11-24", "18:00:00", [("1001", 2, 0, 0)]))
        p = rh.periods(self.root)
        self.assertEqual([x["key"] for x in p["days"]], ["2026-11-24", "2026-11-23"])
        self.assertEqual([x["key"] for x in p["weeks"]], ["2026-11-23"])
        self.assertEqual([x["key"] for x in p["months"]], ["2026-11"])

    def test_invalid_params(self):
        with self.assertRaises(ValueError):
            rh.series(self.root, "year", "cost")
        with self.assertRaises(ValueError):
            rh.series(self.root, "day", "orders")
        with self.assertRaises(ValueError):
            rh.series(self.root, "day", "cost", "all")

    def test_trim_keeps_last_lines(self):
        rh.MAX_LINES_PER_DAY = 3
        try:
            for i in range(5):
                rh.record(self.root, snap("2026-11-24", f"1{i}:00:00", [("1001", i, 0, 0)]))
            path = rh.history_dir(self.root) / "2026-11-24.jsonl"
            self.assertEqual(len(path.read_text(encoding="utf-8").splitlines()), 3)
            self.assertEqual(rh.daily_last(self.root)["2026-11-24"]["1|1001"]["cost"], 4.0)
        finally:
            rh.MAX_LINES_PER_DAY = 800

    def test_prune_old_files(self):
        old_day = (date.today() - timedelta(days=rh.KEEP_DAYS + 5)).isoformat()
        rh.record(self.root, snap(old_day, "18:00:00", [("1001", 1, 1, 0)]))
        rh.record(self.root, snap(date.today().isoformat(), "18:00:00", [("1001", 2, 2, 0)]))
        files = sorted(p.name for p in rh.history_dir(self.root).glob("*.jsonl"))
        self.assertEqual(files, [f"{date.today().isoformat()}.jsonl"])

    def test_accounts_index_updates(self):
        rh.record(self.root, snap("2026-11-24", "18:00:00", [("1001", 1, 1, 0)]))
        data = json.loads(rh.accounts_path(self.root).read_text(encoding="utf-8"))
        self.assertEqual(data["1|1001"], {"name": "账户1001", "owner": "主店"})
        out = rh.series(self.root, "day", "cost", "recent", now=NOW)
        self.assertEqual(out["series"][0]["name"], "账户1001")
        self.assertEqual(out["series"][0]["owner"], "主店")


if __name__ == "__main__":
    unittest.main()
