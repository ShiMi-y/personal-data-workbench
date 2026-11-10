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
  py export_full.py --date 2026-11-04 --accounts 10000001,10000002
"""
from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
import time
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
RAW_BASE = ["date", "account_id", "account_name"]


def _raw_header(rows: list[kx.ProductRow]) -> list[str]:
    """按接口实际返回字段顺序汇总表头（保留原始字段名与原始单位）"""
    keys: list[str] = []
    for r in rows:
        for k in (r.raw or {}).keys():
            if k not in keys:
                keys.append(k)
    return RAW_BASE + keys


def _raw_row(r: kx.ProductRow, keys: list[str]) -> list:
    raw = r.raw or {}
    out = [r.date, r.account_id, r.account_name]
    for k in keys:
        v = raw.get(k)
        out.append(json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v)
    return out


def _write_csv(path: Path, header: list[str], rows: list[list]):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def _write_xlsx(path: Path, acc_rows: list[list], prod_rows: list[list],
                raw_header: list[str] | None = None, raw_rows: list[list] | None = None):
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
    sheets = [(ws1, [12, 12, 18, 12, 12, 12, 8], False),
              (ws2, [12, 12, 18, 18, 46, 12, 12, 12, 8, 10, 10], False)]
    if raw_header and raw_rows:
        ws3 = wb.create_sheet("商品原始数据")
        ws3.append(raw_header)
        for r in raw_rows:
            ws3.append(r)
        sheets.append((ws3, [12, 14, 20] + [14] * (len(raw_header) - 3), True))
    for ws, widths, is_raw in sheets:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="EDEDED")
            c.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        for row in ws.iter_rows(min_row=1):
            for c in row:
                if is_raw or c.column > 4 or ws is ws1:
                    c.alignment = Alignment(horizontal="center")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def run_export(cfg: dict, start: str, end: str, account_ids: list[str] | None = None,
               outdir: Path | None = None, logger=print, range_mode: bool = False,
               cancel=None, retries: int = 2) -> dict:
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
        accounts_meta = [{"account_id": str(a["id"]), "account_name": a.get("name") or ""}
                         for a in accounts]

        for acc in accounts:
            if cancel and cancel():
                raise RuntimeError("任务已被用户停止")
            aid, aname = str(acc["id"]), acc.get("name") or ""
            logger(f"[export] 账户 {aid} {aname}")
            last_err = None
            for attempt in range(max(0, retries) + 1):
                if cancel and cancel():
                    raise RuntimeError("任务已被用户停止")
                home = None
                rows: list[kx.ProductRow] = []
                try:
                    home = client.open_account(aid)
                    if range_mode:
                        rows = client.fetch_products_range(home, aid, aname, days[0], days[-1])
                        logger(f"    {start}~{end} 商品行 {len(rows)} "
                               f"花费 {round(sum(r.cost for r in rows), 2)} "
                               f"成交 {round(sum(r.gmv for r in rows), 2)}")
                    else:
                        for day in days:
                            if cancel and cancel():
                                raise RuntimeError("任务已被用户停止")
                            day_rows = client.fetch_products(home, aid, aname, day)
                            rows.extend(day_rows)
                            logger(f"    {day:%Y-%m-%d} 商品 {len(day_rows)} 个 "
                                   f"花费 {round(sum(r.cost for r in day_rows), 2)} "
                                   f"成交 {round(sum(r.gmv for r in day_rows), 2)}")
                    last_err = None
                except Exception as e:  # noqa: BLE001
                    last_err = e
                    if attempt < retries:
                        wait = 3 * (attempt + 1)
                        logger(f"    失败: {type(e).__name__}: {e}（{wait}s 后自动重试，"
                               f"{attempt + 1}/{retries}）")
                        time.sleep(wait)
                finally:
                    if home is not None:
                        try:
                            home.close()
                        except Exception:
                            pass
                if last_err is None:
                    prod.extend(rows)
                    break
            if last_err is not None:
                logger(f"    失败: {type(last_err).__name__}: {last_err}（自动重试 {retries} 次仍失败）")
                failures.append({"account_id": aid, "account_name": aname,
                                 "error": str(last_err)[:300]})

    acc_rows_data = kx.aggregate(prod)
    prod_rows = [[r.date, r.account_id, r.account_name, r.product_id, r.product_name,
                  r.cost, r.gmv, r.order_cnt, r.roi, r.show_cnt, r.click_cnt] for r in prod]
    acc_rows = [[a["date"], a["account_id"], a["account_name"], a["cost"], a["gmv"],
                 a["order_cnt"], a["product_count"]] for a in acc_rows_data]

    _write_csv(outdir / "products.csv", PRODUCT_CSV, prod_rows)
    _write_csv(outdir / "accounts.csv", ACCOUNT_CSV, acc_rows)
    raw_header = _raw_header(prod)
    raw_keys = raw_header[len(RAW_BASE):]
    raw_rows = [_raw_row(r, raw_keys) for r in prod]
    raw_csv = outdir / "products_raw.csv"
    if raw_keys:
        _write_csv(raw_csv, raw_header, raw_rows)
    xlsx = outdir / f"明细_{start}_{end}.xlsx"
    _write_xlsx(xlsx, acc_rows, prod_rows,
                raw_header if raw_keys else None, raw_rows if raw_keys else None)

    summary = {
        "start": start, "end": end,
        "outdir": str(outdir),
        "accounts_ok": len({a["account_id"] for a in acc_rows_data}),
        "products": len(prod),
        "raw_columns": len(raw_keys),
        "total_cost": round(sum(a["cost"] for a in acc_rows_data), 2),
        "total_gmv": round(sum(a["gmv"] for a in acc_rows_data), 2),
        "failures": failures,
        "files": {"products": str(outdir / "products.csv"),
                  "accounts": str(outdir / "accounts.csv"),
                  "raw": str(raw_csv) if raw_keys else "",
                  "xlsx": str(xlsx)},
    }
    logger(f"[export] 完成：账户 {summary['accounts_ok']}，商品行 {summary['products']}，"
           f"原始字段 {summary['raw_columns']} 个，消耗合计 {summary['total_cost']}，成交合计 {summary['total_gmv']}")
    if failures:
        logger(f"[export] 失败账户 {len(failures)} 个: {[f['account_id'] for f in failures]}")
    summary["accounts_requested"] = accounts_meta
    try:
        # 供网页预览：勾选但无数据/失败的账户也能展示
        (outdir / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass
    return summary


def _read_dict_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _rows_to_objects(rows: list[dict]) -> list[kx.ProductRow]:
    out = []
    for r in rows:
        try:
            out.append(kx.ProductRow(
                date=str(r.get("date") or ""),
                account_id=str(r.get("account_id") or ""),
                account_name=str(r.get("account_name") or ""),
                product_id=str(r.get("product_id") or ""),
                product_name=str(r.get("product_name") or ""),
                cost=float(r.get("cost") or 0),
                gmv=float(r.get("gmv") or 0),
                order_cnt=int(float(r.get("order_cnt") or 0)),
                roi=float(r.get("roi") or 0),
                show_cnt=int(float(r.get("show_cnt") or 0)),
                click_cnt=int(float(r.get("click_cnt") or 0)),
            ))
        except (TypeError, ValueError):
            continue
    return out


def retry_accounts(cfg: dict, start: str, end: str, account_ids: list[str],
                   outdir: Path | None = None, logger=print,
                   retries: int = 2, cancel=None) -> dict:
    """只重试指定账户的区间导出，并合并回已有导出目录（替换这些账户的数据）"""
    ids = [str(x).strip() for x in account_ids if str(x).strip()]
    if not ids:
        raise ValueError("没有指定需要重试的账户")
    outdir = Path(outdir) if outdir else (kx.DATA_DIR / "exports" / f"{start}_{end}")
    main_prod = outdir / "products.csv"
    if not main_prod.exists():
        raise FileNotFoundError(f"导出目录不完整，请先重新拉取: {outdir}")
    tmp = outdir / "_retry"
    shutil.rmtree(tmp, ignore_errors=True)
    logger(f"[export] 重试 {len(ids)} 个账户（{start}~{end}）：{','.join(ids)}")
    res = run_export(cfg, start, end, ids, outdir=tmp, logger=logger,
                     range_mode=True, retries=retries, cancel=cancel)
    id_set = set(ids)
    failed_ids = {str(f.get("account_id")) for f in (res.get("failures") or [])}
    ok_ids = id_set - failed_ids

    # 只替换重试成功的账户；仍失败的账户保留原有数据，避免把已有结果删掉
    merged = [r for r in _read_dict_rows(main_prod) if str(r.get("account_id")) not in ok_ids]
    merged += _read_dict_rows(tmp / "products.csv")
    _write_csv(main_prod, PRODUCT_CSV, [[r.get(k, "") for k in PRODUCT_CSV] for r in merged])

    objs = _rows_to_objects(merged)
    acc_rows_data = kx.aggregate(objs)
    acc_rows = [[a["date"], a["account_id"], a["account_name"], a["cost"], a["gmv"],
                 a["order_cnt"], a["product_count"]] for a in acc_rows_data]
    _write_csv(outdir / "accounts.csv", ACCOUNT_CSV, acc_rows)
    prod_rows = [[r.date, r.account_id, r.account_name, r.product_id, r.product_name,
                  r.cost, r.gmv, r.order_cnt, r.roi, r.show_cnt, r.click_cnt] for r in objs]
    _write_xlsx(outdir / f"明细_{start}_{end}.xlsx", acc_rows, prod_rows)

    raw_main, raw_new = outdir / "products_raw.csv", tmp / "products_raw.csv"
    if raw_main.exists() and raw_new.exists():
        with raw_main.open(encoding="utf-8-sig") as f:
            header = next(csv.reader(f), [])
        with raw_new.open(encoding="utf-8-sig") as f:
            header_new = next(csv.reader(f), [])
        if header and header == header_new:
            raw_rows = [r for r in _read_dict_rows(raw_main)
                        if str(r.get("account_id")) not in ok_ids]
            raw_rows += _read_dict_rows(raw_new)
            _write_csv(raw_main, header, [[r.get(k, "") for k in header] for r in raw_rows])
        else:
            logger("[export] 原始数据表头不一致，未合并 products_raw.csv")
    shutil.rmtree(tmp, ignore_errors=True)

    sp = outdir / "summary.json"
    summary = {}
    if sp.exists():
        try:
            summary = json.loads(sp.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            summary = {}
    old_fail = [f for f in (summary.get("failures") or [])
                if str(f.get("account_id")) not in ok_ids]
    still = list(res.get("failures") or [])
    summary.update({
        "start": start, "end": end, "outdir": str(outdir),
        "accounts_ok": len({a["account_id"] for a in acc_rows_data}),
        "products": len(merged),
        "total_cost": round(sum(a["cost"] for a in acc_rows_data), 2),
        "total_gmv": round(sum(a["gmv"] for a in acc_rows_data), 2),
        "failures": old_fail + still,
    })
    requested = {str(x.get("account_id")): x for x in (summary.get("accounts_requested") or [])}
    for m in res.get("accounts_requested") or []:
        requested.setdefault(str(m.get("account_id")), m)
    summary["accounts_requested"] = list(requested.values())
    try:
        sp.write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError as e:
        logger(f"[export] summary.json 写入失败: {e}")
    logger(f"[export] 重试完成：成功 {len(ids) - len(still)} 个，仍失败 {len(still)} 个")
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