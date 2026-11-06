# -*- coding: utf-8 -*-
"""全量导出：全账号 × 全商品 × 指定日期（默认 T-1）

产出（默认 data/exports/<start>_<end>/）:
  products.csv   商品级明细
  accounts.csv   账户×日期汇总（由商品明细求和，与后台账户口径一致）
  明细_<start>_<end>.xlsx   两个 Sheet：账户汇总 / 商品明细

用法:
  py export_full.py                      # 前一自然日
  py export_full.py --date 2026-11-04
  py export_full.py --start 2026-11-01 --end 2026-11-05
  py export_full.py --date 2026-11-04 --accounts 21373112,20925477
"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import kx_core as kx

BASE = Path(__file__).resolve().parent

PRODUCT_HEADER = ["日期", "账户ID", "账户名", "商品ID", "商品名", "花费", "成交", "成交订单数", "ROI", "曝光", "点击"]
ACCOUNT_HEADER = ["日期", "账户ID", "账户名", "花费", "成交", "成交订单数", "商品数"]

# CSV 用英文表头（供脚本读取），Excel 用中文表头（供人看）
PRODUCT_CSV = ["date", "account_id", "account_name", "product_id", "product_name",
               "cost", "gmv", "order_cnt", "roi", "show_cnt", "click_cnt"]
ACCOUNT_CSV = ["date", "account_id", "account_name", "cost", "gmv", "order_cnt", "product_count"]


def _write_csv(path: Path, header: list[str], rows: list[list]):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def _write_xlsx(path: Path, acc_rows: list[list], prod_rows: list[list]):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "账户汇总"
    ws1.append(ACCOUNT_HEADER)
    for r in acc_rows:
        ws1.append(r)
    ws2 = wb.create_sheet("商品明细")
    ws2.append(PRODUCT_HEADER)
    for r in prod_rows:
        ws2.append(r)
    for ws, widths in ((ws1, [12, 12, 18, 12, 12, 12, 8]), (ws2, [12, 12, 18, 18, 46, 12, 12, 12, 8, 10, 10])):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="EDEDED")
            c.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        for row in ws.iter_rows(min_row=1):
            for c in row:
                if c.column > 4 or ws is ws1:
                    c.alignment = Alignment(horizontal="center")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def run_export(cfg: dict, start: str, end: str, account_ids: list[str] | None = None,
               outdir: Path | None = None, logger=print) -> dict:
    days = kx.day_range(start, end)
    outdir = outdir or (kx.DATA_DIR / "exports" / f"{start}_{end}")
    outdir.mkdir(parents=True, exist_ok=True)

    prod: list[kx.ProductRow] = []
    failures: list[dict] = []

    with kx.Client(cfg, logger) as client:
        all_accounts = client.list_accounts()
        accounts = kx.pick_accounts(all_accounts, cfg, account_ids)
        if not accounts:
            raise RuntimeError("没有可导出的账户（检查登录态或 include/exclude 配置）")
        logger(f"[export] 账户全量 {len(all_accounts)} 个，本次导出 {len(accounts)} 个 × {len(days)} 天 -> {outdir}")

        for acc in accounts:
            aid, aname = str(acc["id"]), acc.get("name") or ""
            logger(f"[export] 账户 {aid} {aname}")
            home = None
            try:
                home = client.open_account(aid)
                for day in days:
                    rows = client.fetch_products(home, aid, aname, day)
                    cost = round(sum(r.cost for r in rows), 2)
                    gmv = round(sum(r.gmv for r in rows), 2)
                    prod.extend(rows)
                    logger(f"    {day:%Y-%m-%d} 商品 {len(rows)} 个 花费 {cost} 成交 {gmv}")
            except Exception as e:  # noqa: BLE001
                logger(f"    失败: {type(e).__name__}: {e}")
                failures.append({"account_id": aid, "account_name": aname, "error": str(e)[:300]})
            finally:
                if home is not None:
                    try:
                        home.close()
                    except Exception:
                        pass

    acc_rows_data = kx.aggregate(prod)
    prod_rows = [[r.date, r.account_id, r.account_name, r.product_id, r.product_name,
                  r.cost, r.gmv, r.order_cnt, r.roi, r.show_cnt, r.click_cnt] for r in prod]
    acc_rows = [[a["date"], a["account_id"], a["account_name"], a["cost"], a["gmv"],
                 a["order_cnt"], a["product_count"]] for a in acc_rows_data]

    _write_csv(outdir / "products.csv", PRODUCT_CSV, prod_rows)
    _write_csv(outdir / "accounts.csv", ACCOUNT_CSV, acc_rows)
    xlsx = outdir / f"明细_{start}_{end}.xlsx"
    _write_xlsx(xlsx, acc_rows, prod_rows)

    summary = {
        "start": start, "end": end,
        "outdir": str(outdir),
        "accounts_ok": len({a["account_id"] for a in acc_rows_data}),
        "products": len(prod),
        "total_cost": round(sum(a["cost"] for a in acc_rows_data), 2),
        "total_gmv": round(sum(a["gmv"] for a in acc_rows_data), 2),
        "failures": failures,
        "files": {"products": str(outdir / "products.csv"), "accounts": str(outdir / "accounts.csv"),
                  "xlsx": str(xlsx)},
    }
    logger(f"[export] 完成：账户 {summary['accounts_ok']}，商品行 {summary['products']}，"
           f"消耗合计 {summary['total_cost']}，成交合计 {summary['total_gmv']}")
    if failures:
        logger(f"[export] 失败账户 {len(failures)} 个: {[f['account_id'] for f in failures]}")
    return summary


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date")
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--accounts", help="逗号分隔账户ID")
    ap.add_argument("--config", default=str(kx.CONFIG_PATH))
    args = ap.parse_args()

    cfg = kx.load_config(Path(args.config))
    cfg = kx.default_account_cfg(cfg)
    if args.start and args.end:
        start, end = args.start, args.end
    else:
        start = end = args.date or kx.default_date(cfg)
    ids = [x.strip() for x in args.accounts.split(",")] if args.accounts else None
    s = run_export(cfg, start, end, ids)
    return 0 if not s["failures"] else 1


if __name__ == "__main__":
    sys.exit(main())

