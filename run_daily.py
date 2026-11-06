# -*- coding: utf-8 -*-
"""一键：拉取全量数据 + 合并进月度表

用法:
  py run_daily.py                       # 默认前一自然日
  py run_daily.py 2026-11-04            # 指定单日
  py run_daily.py --start 2026-11-01 --end 2026-11-05
  py run_daily.py 2026-11-04 --accounts 10000001,10000002
  py run_daily.py 2026-11-04 --no-merge # 只导出不合并
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import export_full
import kx_core as kx
import merge_excel

BASE = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("date", nargs="?", help="单日 YYYY-MM-DD，默认 T-1")
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--accounts")
    ap.add_argument("--file", help="目标 Excel，默认取 config.excel_path")
    ap.add_argument("--no-merge", action="store_true")
    ap.add_argument("--config", default=str(kx.CONFIG_PATH))
    args = ap.parse_args()

    cfg = kx.load_config(Path(args.config))
    cfg = kx.default_account_cfg(cfg)
    if args.start and args.end:
        start, end = args.start, args.end
    else:
        start = end = args.date or kx.default_date(cfg)
    ids = [x.strip() for x in args.accounts.split(",")] if args.accounts else None

    s = export_full.run_export(cfg, start, end, ids)
    if s["failures"]:
        print("[run] 存在导出失败账户，已跳过合并（避免写入不完整数据）")
        return 1
    if args.no_merge:
        return 0

    res = merge_excel.run_merge(cfg, start, end, ids,
                                excel_path=Path(args.file) if args.file else None,
                                backup=bool(cfg.get("auto_backup", True)))
    print("\n===== 校验清单 =====")
    for i, c in enumerate(res["checks"], 1):
        print(f"{i}. [{'PASS' if c['ok'] else 'FAIL'}] {c['name']}" + (f"  {c['detail']}" if c["detail"] else ""))
    print(f"===== {'已保存' if res['saved'] else '未保存'} =====")
    return 0 if res["saved"] else 1


if __name__ == "__main__":
    sys.exit(main())

