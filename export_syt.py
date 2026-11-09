# -*- coding: utf-8 -*-
"""快手小店「生意通」商品列表取数

流程（已在真实环境验证）：
  1. 打开 syt.kwaixiaodian.com/zones/goodsManagement/goods_list
  2. 关掉新手引导浮层
  3. 选日期：自然周（上一周）或「近30日」平台预设
  4. 点「下载数据」→ 平台异步生成任务
  5. 进下载中心，按「提交时间晚于任务开始时间」识别本次新记录
  6. 等它完成后点「下载」落盘，openpyxl 校验表头与行数，写同名 .meta.json

用法：
  python export_syt.py --mode both
  python export_syt.py --mode week --account 30124ac4
  python export_syt.py --mode range --profile "D:\\path\\profile" --out "D:\\out"

说明：
  · 登录态来自账号的 Chrome 持久化 profile（与磁力金牛共用同一套小店登录态）
  · 日期输入框不能直接输入，必须点日历单元格
  · 平台会把日期截断到有数据的最后一天，所以不硬编码文件名
  · 平台记录 id 不保证递增，必须用提交时间判断新任务
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from playwright.sync_api import Page, sync_playwright

import kx_core as kx

GOODS_LIST = "https://syt.kwaixiaodian.com/zones/goodsManagement/goods_list"
DOWNLOAD_CENTER = "https://syt.kwaixiaodian.com/zones/syt-download-center/syt-download-record"
EXPECT_HEADER = ["商品ID", "商品名称", "成交件数", "成交退款率(订单)", "成交金额"]
RECORD_RE = re.compile(r"商品列表(\d{8}_\d{8})_(\d+)\.xlsx")
TIME_RE = re.compile(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")


def last_week(today: date | None = None) -> tuple[date, date]:
    """上一自然周（周一 ~ 周日）"""
    d = today or date.today()
    monday_this = d - timedelta(days=d.weekday())
    start = monday_this - timedelta(days=7)
    return start, start + timedelta(days=6)


def date_label(start: date, end: date) -> str:
    """周期标签：11.2-11.8"""
    return f"{start.month}.{start.day}-{end.month}.{end.day}"


def close_guide(page: Page, logger=print) -> None:
    for label in ("关闭", "我知道了"):
        try:
            loc = page.get_by_text(label, exact=True)
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=4000)
                page.wait_for_timeout(800)
                logger(f"[syt] 已关闭引导浮层（{label}）")
        except Exception:  # noqa: BLE001
            pass


def click_cell(page: Page, iso: str) -> bool:
    return bool(page.evaluate(
        """(t) => { const td = document.querySelector('.ant-picker-dropdown td[title="' + t + '"]');
                   if (!td) return false; td.click(); return true; }""", iso))


def pick_week(page: Page, monday: date, logger=print) -> None:
    iso = monday.isoformat()
    for attempt in range(1, 4):
        try:
            page.get_by_text("自然周", exact=True).first.click(timeout=8000)
        except Exception:  # noqa: BLE001
            page.evaluate("""() => { const el = Array.from(document.querySelectorAll('*'))
                .find(e => e.children.length === 0 && (e.innerText || '').trim() === '自然周');
                if (el) el.click(); }""")
        page.wait_for_timeout(2500)
        close_guide(page, logger)
        for _ in range(6):
            if click_cell(page, iso):
                logger(f"[syt] 已选自然周 {iso}")
                page.wait_for_timeout(2000)
                return
            page.evaluate("""() => { const b = document.querySelector(
                '.ant-picker-dropdown .ant-picker-header-prev-btn'); if (b) b.click(); }""")
            page.wait_for_timeout(1000)
        logger(f"[syt] 第 {attempt} 次没找到 {iso}，重开日期面板重试")
        try:
            page.keyboard.press("Escape")
        except Exception:  # noqa: BLE001
            pass
        page.wait_for_timeout(1200)
    raise RuntimeError(f"周选择器里找不到 {iso}")


def pick_preset(page: Page, label: str, logger=print) -> None:
    page.get_by_text(label, exact=True).first.click(timeout=8000)
    page.wait_for_timeout(2500)
    logger(f"[syt] 已选预设 {label}")


def click_export(page: Page, logger=print) -> None:
    page.get_by_text("下载数据", exact=True).first.click(timeout=10000)
    page.wait_for_timeout(3000)
    logger("[syt] 已点「下载数据」，平台开始生成任务")


def scan_records(dc: Page) -> list[dict]:
    """读取下载中心当前页记录：[{id, file, time}]"""
    try:
        rows = dc.evaluate(
            "() => Array.from(document.querySelectorAll('tbody tr')).map(tr => tr.innerText)")
    except Exception:  # noqa: BLE001
        rows = []
    out = []
    for t in rows or []:
        m = RECORD_RE.search(str(t))
        if not m:
            continue
        tm = TIME_RE.search(str(t))
        out.append({
            "id": m.group(2),
            "file": f"商品列表{m.group(1)}_{m.group(2)}.xlsx",
            "time": tm.group(1) if tm else "",
        })
    return out


def fetch_new_record(ctx, dc: Page, since: str, out_dir: Path,
                     timeout_s: int = 180, logger=print) -> Path:
    """等待提交时间 >= since 的新记录完成并下载"""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        dc.goto(DOWNLOAD_CENTER, wait_until="domcontentloaded", timeout=60000)
        dc.wait_for_timeout(4500)
        recs = scan_records(dc)
        new = [r for r in recs if r["time"] and r["time"] >= since]
        if new:
            new.sort(key=lambda r: (r["time"], int(r["id"])))
            rec = new[-1]
            fname = rec["file"]
            row = dc.locator("tr", has_text=fname).first
            try:
                row_txt = row.inner_text(timeout=8000)
            except Exception:  # noqa: BLE001
                row_txt = ""
            if "下载完成" in row_txt:
                logger(f"[syt] 新记录 {rec['id']} 已完成：{fname}")
                with dc.expect_download(timeout=60000) as dl:
                    row.get_by_text("下载", exact=True).first.click(timeout=15000)
                d = dl.value
                out_dir.mkdir(parents=True, exist_ok=True)
                path = out_dir / d.suggested_filename
                d.save_as(str(path))
                logger(f"[syt] 已下载 {path.name}（{path.stat().st_size} B）")
                return path
            if "失败" in row_txt:
                raise RuntimeError(f"平台任务失败：{fname}")
            logger(f"[syt] 新记录 {rec['id']} 生成中：{fname}")
        time.sleep(3)
    raise TimeoutError(f"{timeout_s}s 内未等到新任务完成")


def verify(path: Path) -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(max_row=3, values_only=True))
    header = [str(c) if c is not None else "" for c in rows[0]]
    body = [[str(c)[:20] if c is not None else "" for c in r][:6] for r in rows[1:3]]
    info = {
        "file": path.name,
        "sheet": ws.title,
        "rows": ws.max_row,
        "cols": ws.max_column,
        "header": header,
        "sample": body,
        "header_ok": header == EXPECT_HEADER,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "size": path.stat().st_size,
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    wb.close()
    return info


def run_one(ctx, dc: Page, mode: str, start: date, out_dir: Path,
            timeout_s: int, logger=print) -> dict:
    # 先停在下载中心取当前时间作为基线：平台记录 id 不递增，只能用提交时间判断
    dc.goto(DOWNLOAD_CENTER, wait_until="domcontentloaded", timeout=60000)
    dc.wait_for_timeout(4000)
    since = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logger(f"[syt] 任务基线时间={since}")

    page = ctx.new_page()
    page.goto(GOODS_LIST, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(9000)
    close_guide(page, logger)
    if mode == "week":
        pick_week(page, start, logger)
    else:
        pick_preset(page, "近30日", logger)
    click_export(page, logger)
    page.close()

    path = fetch_new_record(ctx, dc, since, out_dir, timeout_s, logger)
    info = verify(path)
    m = RECORD_RE.match(path.name)
    info.update({"mode": mode,
                 "start": m.group(1).split("_")[0] if m else "",
                 "end": m.group(1).split("_")[1] if m else "",
                 "record_id": m.group(2) if m else ""})
    (path.with_suffix(".meta.json")).write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    return info


def run_export(cfg: dict, out_dir: Path | None = None, modes=("week", "range"),
               timeout_s: int = 300, logger=print) -> dict:
    """供工作台调用：下载生意通商品列表"""
    profile = Path(cfg.get("_profile_dir") or (kx.PROFILES_DIR / "syt"))
    out_dir = Path(out_dir) if out_dir else (kx.DATA_DIR / "syt")
    profile.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    w_start, w_end = last_week()
    jobs = []
    if "week" in modes:
        jobs.append(("week", w_start, w_end))
    if "range" in modes:
        jobs.append(("range", w_start, w_end))

    results, errors = [], []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(profile), channel="chrome", headless=True,
            args=["--no-sandbox", "--disable-gpu"],
            viewport={"width": 1680, "height": 1050}, accept_downloads=True)
        dc = ctx.new_page()
        try:
            for mode, s, e in jobs:
                label = (f"上一周({date_label(s, e)})" if mode == "week"
                         else "近30日(平台口径)")
                logger(f"[syt] 开始：{label}")
                try:
                    results.append(run_one(ctx, dc, mode, s, out_dir, timeout_s, logger))
                except Exception as exc:  # noqa: BLE001
                    logger(f"[syt] 失败：{label} -> {type(exc).__name__}: {str(exc)[:200]}")
                    errors.append({"mode": mode, "error": str(exc)[:300]})
        finally:
            try:
                ctx.close()
            except Exception:  # noqa: BLE001
                pass

    logger("[syt] 汇总：")
    for r in results:
        logger(f"  ok {r['file']}  {r['rows']}x{r['cols']}  "
               f"{r['start']}~{r['end']}  表头{'匹配' if r['header_ok'] else '不匹配'}")
    for e in errors:
        logger(f"  x {e['mode']}  {e['error'][:120]}")
    return {"files": results, "errors": errors, "out_dir": str(out_dir),
            "week": date_label(w_start, w_end)}


def main() -> int:
    ap = argparse.ArgumentParser(description="快手生意通商品列表取数")
    ap.add_argument("--mode", choices=["week", "range", "both"], default="both")
    ap.add_argument("--account", help="登录账号 id，默认取配置里第一个")
    ap.add_argument("--profile", help="覆盖 profile 目录")
    ap.add_argument("--out", help="输出目录，默认 data/syt")
    ap.add_argument("--timeout", type=int, default=300, help="单任务等待秒数")
    args = ap.parse_args()

    cfg = kx.load_config()
    if args.account:
        acct = kx.find_login_account(cfg, args.account)
        if not acct:
            print(f"[syt] 未找到登录账号: {args.account}", flush=True)
            return 2
        cfg = kx.account_cfg(cfg, acct)
    else:
        cfg = kx.default_account_cfg(cfg)
    if args.profile:
        cfg["_profile_dir"] = args.profile

    modes = ("week", "range") if args.mode == "both" else (args.mode,)
    res = run_export(cfg, Path(args.out) if args.out else None, modes, args.timeout)
    return 0 if not res["errors"] and all(f.get("header_ok") for f in res["files"]) else 1


if __name__ == "__main__":
    sys.exit(main())
