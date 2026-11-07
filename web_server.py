# -*- coding: utf-8 -*-
"""本地数据工作台：选择日期/账户 -> 拉取全量 -> 预览 -> 合并进月度表

启动:
  py web_server.py            # 默认 http://127.0.0.1:8765
  py web_server.py --port 9000 --no-browser
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import urllib.error
import urllib.request
import uuid
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import app_runtime
import app_secret

# 必须在导入其他本地模块前接管 stdout/stderr：打包为无控制台窗口时 print 无处可去
app_runtime.install()

import export_full  # noqa: E402
import kx_core as kx  # noqa: E402
import merge_excel  # noqa: E402

WEB_DIR = kx.BUNDLE_DIR / "web"  # 打包后页面资源在解包目录

CFG = kx.load_config()
CFG_LOCK = threading.Lock()
REALTIME_LOCK = threading.Lock()
JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()
_SINGLE_INSTANCE_MUTEX = None


def _job_new(kind: str) -> str:
    jid = uuid.uuid4().hex[:12]
    with JOBS_LOCK:
        JOBS[jid] = {"id": jid, "kind": kind, "status": "running", "logs": [],
                     "progress": None, "result": None, "error": None}
    return jid


def _job_log(jid: str, msg: str):
    with JOBS_LOCK:
        job = JOBS.get(jid)
        kind = job["kind"] if job else "job"
        if job is not None:
            job["logs"].append(msg)
            if len(job["logs"]) > 800:
                del job["logs"][:200]
    # 同时落到 logs/ 日志文件；消息自带 [标签] 时不重复加
    print(msg if msg.startswith("[") else f"[{kind}] {msg}")


def _job_done(jid: str, result=None, error: str | None = None):
    with JOBS_LOCK:
        job = JOBS.get(jid)
        if job is not None:
            job["status"] = "error" if error else "done"
            job["result"] = result
            job["error"] = error


def _job_progress(jid: str, **fields):
    """任务进度（更新下载/解压时供网页轮询展示）"""
    with JOBS_LOCK:
        job = JOBS.get(jid)
        if job is None:
            return
        prog = dict(job.get("progress") or {})
        prog.update(fields)
        prog["updated_at"] = round(time.time(), 2)
        job["progress"] = prog


def _export_dir(start: str, end: str) -> Path:
    return kx.DATA_DIR / "exports" / f"{start}_{end}"


def _read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _excel_accounts(path: Path) -> list[str]:
    """读月度表第 3 行账户ID（按动态商品组列扫描，兼容后加列）"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            ws = wb["消耗成交"] if "消耗成交" in wb.sheetnames else wb[wb.sheetnames[0]]
            _total_start, cost_cols, _total_row = merge_excel._layout(ws)
            ids = []
            for col in cost_cols:
                v = ws.cell(row=3, column=col).value
                if v is not None and str(v).strip():
                    ids.append(str(v).strip())
            return ids
        finally:
            wb.close()
    except Exception as e:  # noqa: BLE001
        print(f"[web] 读取月度表账户列失败: {type(e).__name__}: {e}")
        return []


def _excel_products(path: Path) -> list[dict]:
    """读月度表第 1~3 行：商品名 / 商品ID / 所属账户ID"""
    try:
        return merge_excel.read_product_groups(path)
    except Exception as e:  # noqa: BLE001
        print(f"[web] 读取月度表商品列失败: {type(e).__name__}: {e}")
        return []


def _tag_legacy_realtime(data: dict | None, cfg: dict) -> dict | None:
    """旧版实时缓存没有 owner 字段，归到第一个（迁移）账号"""
    if not data:
        return data
    accounts = kx.login_accounts(cfg)
    if not accounts:
        return data
    owner = next((a for a in accounts if a.get("legacy_profile")), accounts[0])
    oid, olabel = str(owner.get("id")), str(owner.get("label") or owner.get("id"))
    for key in ("accounts", "products", "failures", "alerts"):
        for x in data.get(key) or []:
            if isinstance(x, dict) and not x.get("owner_id"):
                x["owner_id"] = oid
                x["owner_label"] = olabel
    return data


def api_state() -> dict:
    with CFG_LOCK:
        cfg = dict(CFG)
    login_accounts = _login_accounts_public(cfg)
    excel_accounts, excel_products, excel_ready = _excel_by_owner(cfg)
    private = app_secret.private_status()
    thresholds = {}
    for a in login_accounts:
        thresholds.update(a.get("alert_thresholds") or {})
    return {
        "login_accounts": login_accounts,
        "accounts": _cached_accounts(cfg),
        "excel_accounts": excel_accounts,
        "excel_products": excel_products,
        "excel_ready": excel_ready,
        "excel_total": len(login_accounts),
        "default_date": kx.default_date(cfg),
        "today": kx.today_str(),
        "version": str(cfg.get("version") or "1.0.0"),
        "update_source": str(cfg.get("update_source") or "github"),
        "webdav_configured": bool(app_secret.load_webdav()),
        "webdav_private_ready": private["ready"],
        "webdav_private_legacy": private["legacy"],
        "webdav_private_unlocked": private["unlocked"],
        "alert_thresholds": thresholds,
        "has_github_token": bool(os.environ.get("GITHUB_TOKEN", "").strip()),
        "realtime": _tag_legacy_realtime(kx.load_realtime_cache(), cfg),
        "headless": bool(cfg.get("headless", True)),
        "has_credentials": any(a.get("has_password") for a in login_accounts),
        "exports": sorted([p.name for p in (kx.DATA_DIR / "exports").glob("*") if p.is_dir()], reverse=True)[:20]
        if (kx.DATA_DIR / "exports").exists() else [],
    }


EXCEL_SUFFIXES = (".xlsx", ".xlsm")


def _config_snapshot() -> dict:
    with CFG_LOCK:
        return dict(CFG)


def _persist_config(cfg: dict):
    kx.CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    with CFG_LOCK:
        CFG.clear()
        CFG.update(cfg)


def _login_accounts_public(cfg: dict) -> list[dict]:
    """账号列表（密码不返回，只标记是否已设置）"""
    out = []
    for a in kx.login_accounts(cfg):
        p = str(a.get("excel_path") or "").strip()
        out.append({
            "id": str(a.get("id")),
            "label": str(a.get("label") or a.get("login_account") or a.get("id")),
            "login_account": str(a.get("login_account") or ""),
            "login_role": str(a.get("login_role") or "auto"),
            "excel_path": p,
            "excel_exists": bool(p) and kx.resolve_path(p).exists(),
            "enabled": a.get("enabled", True) is not False,
            "has_password": bool(a.get("password")),
            "include_accounts": list(a.get("include_accounts") or []),
            "exclude_accounts": list(a.get("exclude_accounts") or []),
            "alert_thresholds": dict(a.get("alert_thresholds") or {}),
            "legacy_profile": bool(a.get("legacy_profile")),
        })
    return out


def _enabled_login_accounts(cfg: dict, ids: list[str] | None = None) -> list[dict]:
    pool = [a for a in kx.login_accounts(cfg) if a.get("enabled", True) is not False]
    if ids:
        wanted = {str(x) for x in ids}
        pool = [a for a in pool if str(a.get("id")) in wanted]
    if not pool:
        raise RuntimeError("没有可用的登录账号，请先在「账号管理」添加并启用")
    return pool


def _owner_label(cfg: dict, owner_id: str) -> str:
    a = kx.find_login_account(cfg, owner_id)
    if not a:
        return str(owner_id)
    return str(a.get("label") or a.get("login_account") or a.get("id"))


def _tag_accounts(accounts: list[dict], owner: dict) -> list[dict]:
    out = []
    for x in accounts or []:
        row = dict(x)
        row["owner_id"] = str(owner.get("id"))
        row["owner_label"] = str(owner.get("label") or owner.get("login_account") or owner.get("id"))
        out.append(row)
    return out


def _cached_accounts(cfg: dict) -> list[dict]:
    """汇总各登录账号缓存的投放账户，并挂上账号标签"""
    out = []
    for a in kx.login_accounts(cfg):
        if a.get("enabled", True) is False:
            continue
        accs, _ = kx.load_accounts_cache(None if a.get("legacy_profile") else str(a.get("id")))
        out.extend(_tag_accounts(accs, a))
    return out


def _excel_by_owner(cfg: dict) -> tuple[list[str], list[dict], int]:
    """月度表里的账户ID / 商品映射（按登录账号分别读取）"""
    account_ids: list[str] = []
    products: list[dict] = []
    ready = 0
    for a in kx.login_accounts(cfg):
        p = str(a.get("excel_path") or "").strip()
        if not p:
            continue
        path = kx.resolve_path(p)
        if not path.exists():
            continue
        ready += 1
        for aid in _excel_accounts(path):
            if aid not in account_ids:
                account_ids.append(aid)
        products.extend(_tag_accounts(_excel_products(path), a))
    return account_ids, products, ready


def _export_owner_dir(start: str, end: str, owner_id: str) -> Path:
    return _export_dir(start, end) / str(owner_id)


def _owner_export_dir(start: str, end: str, account: dict) -> Path:
    """优先用按账号拆分的目录；兼容单账号时期的旧目录"""
    d = _export_owner_dir(start, end, str(account.get("id")))
    if d.exists():
        return d
    legacy = _export_dir(start, end)
    if account.get("legacy_profile") and (legacy / "accounts.csv").exists():
        return legacy
    return d


def _selected_owner_map(body: dict) -> tuple[list[dict], dict[str, list[str] | None]]:
    """请求里的登录账号 + 每个账号下要处理的投放账户"""
    cfg = _config_snapshot()
    owners = _enabled_login_accounts(cfg, body.get("owners") or None)
    raw = body.get("accounts_by_owner") or {}
    mapping: dict[str, list[str] | None] = {}
    for a in owners:
        ids = raw.get(str(a.get("id")))
        mapping[str(a.get("id"))] = [str(x) for x in ids] if ids else None
    return owners, mapping



def pick_excel_file(current: str = "") -> str:
    """弹出 Windows 原生文件选择框（只筛 Excel 工作簿），取消返回空串"""
    import ctypes
    from ctypes import wintypes

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD),
            ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE),
            ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR),
            ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD),
            ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD),
            ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD),
            ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR),
            ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD),
            ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR),
            ("lCustData", wintypes.LPARAM),
            ("lpfnHook", wintypes.LPVOID),
            ("lpTemplateName", wintypes.LPCWSTR),
            ("pvReserved", wintypes.LPVOID),
            ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = ctypes.create_unicode_buffer(4096)
    if current:
        buf.value = current

    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.lpstrFilter = ("Excel 工作簿 (*.xlsx;*.xlsm)\0*.xlsx;*.xlsm\0所有文件 (*.*)\0*.*\0\0")
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 4096
    ofn.lpstrTitle = "选择月度 Excel 文档"
    ofn.lpstrDefExt = "xlsx"
    ofn.nFilterIndex = 1
    # EXPLORER | FILEMUSTEXIST | PATHMUSTEXIST | NOCHANGEDIR
    ofn.Flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000008
    if current:
        parent = Path(current).parent
        if parent.exists():
            ofn.lpstrInitialDir = str(parent)

    if not ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return ""
    return buf.value


CONFIG_KEYS = ("headless", "date_offset_days", "page_size",
               "manual_login_timeout", "filter_by_excel",
               "auto_backup", "backup_limit", "version", "update_repo",
               "update_source", "update_webdav_path",
               "realtime_auto_refresh", "realtime_interval_minutes",
               "realtime_window_start", "realtime_window_end",
               "realtime_retry_enabled", "realtime_retry_max")


def api_config() -> dict:
    """读取配置（密码不回显，只返回是否已设置）"""
    with CFG_LOCK:
        cfg = dict(CFG)
    out = {k: cfg.get(k) for k in CONFIG_KEYS}
    web = cfg.get("web") or {}
    out["web"] = {"host": web.get("host", "127.0.0.1"), "port": int(web.get("port", 8765))}
    dav = app_secret.load_webdav()
    out["webdav_configured"] = bool(dav)
    private = app_secret.private_status()
    out["webdav_private_ready"] = private["ready"]
    out["webdav_private_legacy"] = private["legacy"]
    out["webdav_private_unlocked"] = private["unlocked"]
    out["webdav_url"] = (dav or {}).get("url") or ""
    out["webdav_username"] = (dav or {}).get("username") or ""
    return {"path": str(kx.CONFIG_PATH), "config": out}


def _as_int(v, name: str, lo: int, hi: int) -> int:
    try:
        n = int(str(v).strip())
    except Exception:  # noqa: BLE001
        raise ValueError(f"{name} 必须是整数")
    if not (lo <= n <= hi):
        raise ValueError(f"{name} 需在 {lo}~{hi} 之间")
    return n


def _as_time(v, name: str) -> str:
    """校验 HH:MM，返回补零后的字符串"""
    s = str(v or "").strip()
    try:
        hh, mm = s.split(":")
        hi, mi = int(hh), int(mm)
    except Exception:  # noqa: BLE001
        raise ValueError(f"{name} 格式应为 HH:MM")
    if not (0 <= hi <= 23 and 0 <= mi <= 59):
        raise ValueError(f"{name} 不合法: {s}")
    return f"{hi:02d}:{mi:02d}"


def _as_ids(v) -> list[str]:
    items = v if isinstance(v, list) else str(v or "").replace("，", ",").split(",")
    return [str(x).strip() for x in items if str(x).strip()]


ALERT_COOLDOWN = 1800  # 同一账户重复告警的最小间隔（秒）
_ALERT_STATE: dict[str, tuple[bool, float]] = {}


def _eval_alerts(cfg: dict, data: dict) -> list[dict]:
    """实时的账户消耗与配置预警线比较（严格大于才告警）"""
    th = cfg.get("alert_thresholds") or {}
    out = []
    for a in data.get("accounts") or []:
        aid = str(a.get("account_id"))
        try:
            limit = float(th.get(aid) or 0)
        except (TypeError, ValueError):
            continue
        if limit <= 0:
            continue
        cost = float(a.get("cost") or 0)
        if cost > limit + 1e-9:
            out.append({
                "account_id": aid,
                "account_name": a.get("account_name") or "",
                "product_name": a.get("product_name") or "",
                "cost": round(cost, 2),
                "limit": round(limit, 2),
                "over": round(cost - limit, 2),
            })
    return out


def _notify_alerts(alerts: list[dict], logger=None):
    """超线时发 Windows 气泡；同一账户 30 分钟内不重复弹"""
    now = time.time()
    over_ids = {a["account_id"] for a in alerts}
    for aid in list(_ALERT_STATE):
        if aid not in over_ids:
            _ALERT_STATE[aid] = (False, 0.0)

    fresh = []
    for a in alerts:
        was, last = _ALERT_STATE.get(a["account_id"], (False, 0.0))
        if (not was) or now - last >= ALERT_COOLDOWN:
            fresh.append(a)
            _ALERT_STATE[a["account_id"]] = (True, now)
        else:
            _ALERT_STATE[a["account_id"]] = (True, last)

    if not fresh:
        if logger and alerts:
            logger(f"[alert] {len(alerts)} 个账户超预警线（冷却中，不重复弹窗）")
        return
    lines = [f"{a['account_name'] or a['product_name']} 消耗 {a['cost']:.2f} / 预警线 {a['limit']:.2f}"
             for a in fresh[:4]]
    if len(fresh) > 4:
        lines.append(f"…另有 {len(fresh) - 4} 个账户")
    text = "\n".join(lines)
    if logger:
        logger(f"[alert] {len(fresh)} 个账户超预警线：" + "；".join(lines))
    app_runtime.notify(f"消耗预警（{len(fresh)} 个账户）", text)


def collect_realtime(cfg: dict, owners: list[dict] | None = None, logger=print,
                     blocking: bool = True) -> dict | None:
    """抓取一次实时数据并保存缓存/触发预警；供页面任务与后台调度共用"""
    if not REALTIME_LOCK.acquire(blocking=blocking):
        return None
    try:
        if owners is None:
            owners = _enabled_login_accounts(cfg)
        started = time.time()
        acc_rows, prod_rows, alerts, failures, owners_done = [], [], [], [], []
        total_cost = total_gmv = total_orders = 0.0
        for acct in owners:
            label = acct.get("label") or acct.get("id")
            logger(f"[实时] 账号「{label}」开始…")
            try:
                ecfg = kx.account_cfg(cfg, acct)
                res = kx.fetch_today(ecfg, None, logger=logger)
                acc_rows.extend(_tag_accounts(res.get("accounts"), acct))
                prod_rows.extend(_tag_accounts(res.get("products"), acct))
                for f in res.get("failures") or []:
                    failures.append(dict(f, owner_id=str(acct.get("id")), owner_label=label))
                owner_alerts = _eval_alerts(ecfg, res)
                for x in owner_alerts:
                    x["owner_id"] = str(acct.get("id"))
                    x["owner_label"] = label
                alerts.extend(owner_alerts)
                t = res.get("totals") or {}
                total_cost += float(t.get("cost") or 0)
                total_gmv += float(t.get("gmv") or 0)
                total_orders += float(t.get("order_cnt") or 0)
                owners_done.append({"id": str(acct.get("id")), "label": label, "ok": True})
            except Exception as e:  # noqa: BLE001
                logger(f"[实时] 账号「{label}」失败: {type(e).__name__}: {e}")
                failures.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                 "account_id": "", "account_name": "", "error": str(e)[:300]})
                owners_done.append({"id": str(acct.get("id")), "label": label, "ok": False,
                                    "error": str(e)[:300]})

        if not any(o["ok"] for o in owners_done):
            raise RuntimeError("全部登录账号实时抓取失败")
        combined = {
            "date": kx.today_str(),
            "fetched_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "elapsed": round(time.time() - started, 1),
            "accounts": acc_rows,
            "products": prod_rows,
            "totals": {
                "cost": round(total_cost, 2),
                "gmv": round(total_gmv, 2),
                "roi": round(total_gmv / total_cost, 2) if total_cost else 0,
                "order_cnt": int(total_orders),
                "account_rows": len(acc_rows),
                "product_rows": len(prod_rows),
            },
            "failures": failures,
            "owners": owners_done,
            "alerts": alerts,
        }
        kx.save_realtime_cache(combined)
        _notify_alerts(alerts, logger=logger)
        return combined
    finally:
        REALTIME_LOCK.release()


def _hm_minutes(value, default: str) -> int:
    try:
        hh, mm = str(value or default).split(":")
        return int(hh) * 60 + int(mm)
    except Exception:  # noqa: BLE001
        hh, mm = default.split(":")
        return int(hh) * 60 + int(mm)


def _in_realtime_window(cfg: dict) -> bool:
    tm = time.localtime()
    cur = tm.tm_hour * 60 + tm.tm_min
    a = _hm_minutes(cfg.get("realtime_window_start"), "08:50")
    b = _hm_minutes(cfg.get("realtime_window_end"), "18:05")
    return a <= cur <= b if a <= b else (cur >= a or cur <= b)


def realtime_scheduler_loop():
    """后台自动刷新：网页关闭、只留托盘时也能按间隔抓取并预警"""
    last_attempt = 0.0
    while True:
        time.sleep(20)
        cfg = _config_snapshot()
        if not cfg.get("realtime_auto_refresh"):
            continue
        try:
            interval = max(1, int(cfg.get("realtime_interval_minutes") or 5)) * 60
        except (TypeError, ValueError):
            interval = 300
        now = time.time()
        if now - last_attempt < interval:
            continue
        if not _in_realtime_window(cfg):
            continue
        try:
            cache_age = now - kx.REALTIME_CACHE.stat().st_mtime
        except OSError:
            cache_age = interval + 1
        if cache_age < interval:
            last_attempt = now
            continue
        try:
            owners = _enabled_login_accounts(cfg)
        except RuntimeError as e:
            last_attempt = now
            print(f"[auto] 跳过自动刷新: {e}")
            continue
        last_attempt = now
        print(f"[auto] 开始自动刷新（{len(owners)} 个账号，间隔 {interval // 60} 分钟）")
        try:
            res = collect_realtime(cfg, owners, logger=lambda m: print(f"[auto] {m}"),
                                   blocking=False)
            if res is None:
                print("[auto] 已有实时任务在运行，本轮跳过")
        except Exception as e:  # noqa: BLE001
            print(f"[auto] 自动刷新失败: {type(e).__name__}: {e}")



def api_save_config(body: dict) -> dict:
    """在网页上修改全局配置并写回 config.json（登录账号相关在「账号管理」里保存）"""
    with CFG_LOCK:
        cfg = dict(CFG)
    if "date_offset_days" in body:
        cfg["date_offset_days"] = _as_int(body["date_offset_days"], "日期偏移", 0, 30)
    if "page_size" in body:
        cfg["page_size"] = _as_int(body["page_size"], "页面条数", 10, 100)
    if "manual_login_timeout" in body:
        cfg["manual_login_timeout"] = _as_int(body["manual_login_timeout"], "扫码登录等待", 60, 3600)
    if "headless" in body:
        cfg["headless"] = bool(body["headless"])
    if "filter_by_excel" in body:
        cfg["filter_by_excel"] = bool(body["filter_by_excel"])
    if "auto_backup" in body:
        cfg["auto_backup"] = bool(body["auto_backup"])
    if "backup_limit" in body:
        n = _as_int(body["backup_limit"], "备份数量上限", -1, 999)
        if n == 0:
            raise ValueError("备份数量上限填 -1（不限制）或 1~999")
        cfg["backup_limit"] = n
    if "version" in body:
        v = str(body.get("version") or "").strip()
        if not v:
            raise ValueError("版本号不能为空")
        if len(v) > 32 or not re.fullmatch(r"[A-Za-z0-9._+\-]+", v):
            raise ValueError("版本号只能含字母、数字、点、下划线、加号、减号，长度 ≤ 32")
        cfg["version"] = v
    if "update_repo" in body:
        repo = str(body.get("update_repo") or "").strip().strip("/")
        if repo and not re.fullmatch(r"[\w.-]+/[\w.-]+", repo):
            raise ValueError("update_repo 格式应为 owner/repo")
        cfg["update_repo"] = repo
    if "update_source" in body:
        src = str(body.get("update_source") or "github").strip().lower()
        if src not in ("github", "webdav", "webdav_private"):
            raise ValueError("更新源只能是 github / webdav / webdav_private")
        cfg["update_source"] = src
    if "update_webdav_path" in body:
        cfg["update_webdav_path"] = str(body.get("update_webdav_path") or "/version.json").strip() or "/version.json"
    if "realtime_auto_refresh" in body:
        cfg["realtime_auto_refresh"] = bool(body["realtime_auto_refresh"])
    if "realtime_interval_minutes" in body:
        cfg["realtime_interval_minutes"] = _as_int(
            body["realtime_interval_minutes"], "自动刷新间隔（分钟）", 1, 720)
    if "realtime_window_start" in body:
        cfg["realtime_window_start"] = _as_time(body["realtime_window_start"], "自动刷新开始时间")
    if "realtime_window_end" in body:
        cfg["realtime_window_end"] = _as_time(body["realtime_window_end"], "自动刷新结束时间")
    if "realtime_retry_enabled" in body:
        cfg["realtime_retry_enabled"] = bool(body["realtime_retry_enabled"])
    if "realtime_retry_max" in body:
        cfg["realtime_retry_max"] = _as_int(body["realtime_retry_max"], "失败重试次数", 1, 10)
    if isinstance(body.get("web"), dict):
        web = dict(cfg.get("web") or {})
        if body["web"].get("host"):
            host = str(body["web"]["host"]).strip().lower()
            if host not in ("127.0.0.1", "localhost", "::1"):
                raise ValueError("网页监听地址仅支持本机回环地址（127.0.0.1 / localhost / ::1）")
            web["host"] = host
        if body["web"].get("port"):
            web["port"] = _as_int(body["web"]["port"], "端口", 1024, 65535)
        cfg["web"] = web

    _persist_config(cfg)
    print(f"[web] 配置已保存: version={cfg.get('version')} update_source={cfg.get('update_source')} "
          f"auto_backup={cfg.get('auto_backup')} backup_limit={cfg.get('backup_limit')} "
          f"accounts={len(kx.login_accounts(cfg))}")
    return {"ok": True, **api_config()}


def _alert_thresholds(v) -> dict:
    if not isinstance(v, dict):
        raise ValueError("alert_thresholds 必须是 {账户ID: 预警线} 的对象")
    th = {}
    for k, val in v.items():
        key = str(k).strip()
        if not key.isdigit():
            raise ValueError(f"账户ID 不合法: {key}")
        if val in (None, ""):
            continue
        try:
            num = round(float(val), 2)
        except (TypeError, ValueError):
            raise ValueError(f"账户 {key} 的预警线必须是数字")
        if num <= 0:
            continue
        if num > 10_000_000:
            raise ValueError(f"账户 {key} 的预警线过大")
        th[key] = num
    return th


def api_save_login_account(body: dict) -> dict:
    """新增 / 修改一个快手小店登录账号（账号与其月度表一对一）"""
    with CFG_LOCK:
        cfg = dict(CFG)
    accounts = [dict(a) for a in kx.login_accounts(cfg)]
    aid = str(body.get("id") or "").strip()
    cur = next((a for a in accounts if str(a.get("id")) == aid), None)

    if cur is None:
        cur = {
            "id": uuid.uuid4().hex[:8],
            "label": "",
            "login_account": "",
            "password": "",
            "login_role": "auto",
            "excel_path": "",
            "enabled": True,
            "include_accounts": [],
            "exclude_accounts": [],
            "alert_thresholds": {},
            "legacy_profile": False,
        }
        accounts.append(cur)

    if "label" in body or not cur.get("label"):
        label = str(body.get("label") or "").strip()
        cur["label"] = label or cur.get("login_account") or f"账号 {len(accounts)}"
    if "login_account" in body:
        login = str(body.get("login_account") or "").strip()
        for other in accounts:
            if other is cur:
                continue
            if str(other.get("login_account") or "") == login and login:
                raise ValueError(f"手机号 {login} 已被账号「{other.get('label')}」使用")
        cur["login_account"] = login
    if body.get("password"):
        cur["password"] = app_secret.protect_text(str(body["password"]))
    if "login_role" in body:
        role = str(body.get("login_role") or "auto").strip().lower()
        if role not in ("auto", "owner", "staff"):
            raise ValueError("登录身份只能是 auto / owner / staff")
        cur["login_role"] = role
    if "excel_path" in body:
        p = str(body.get("excel_path") or "").strip()
        if p and Path(p).suffix.lower() not in EXCEL_SUFFIXES:
            raise ValueError("月度表仅支持 .xlsx / .xlsm 格式")
        cur["excel_path"] = p
    if "enabled" in body:
        cur["enabled"] = bool(body["enabled"])
    if "include_accounts" in body:
        cur["include_accounts"] = _as_ids(body["include_accounts"])
    if "exclude_accounts" in body:
        cur["exclude_accounts"] = _as_ids(body["exclude_accounts"])
    if "alert_thresholds" in body:
        cur["alert_thresholds"] = _alert_thresholds(body.get("alert_thresholds") or {})

    if not str(cur.get("login_account") or "").strip() and not str(cur.get("excel_path") or "").strip():
        raise ValueError("请至少填写登录手机号或月度表路径")

    cfg["accounts"] = accounts
    _persist_config(cfg)
    login_log = (str(cur.get("login_account") or "")[:3] + "****") if cur.get("login_account") else "（用登录态）"
    print(f"[account] 已保存登录账号: id={cur['id']} label={cur.get('label')} "
          f"login={login_log} excel={cur.get('excel_path')} enabled={cur.get('enabled')}")
    return {"ok": True, "login_accounts": _login_accounts_public(cfg)}


def api_delete_login_account(body: dict) -> dict:
    aid = str(body.get("id") or "").strip()
    if not aid:
        raise ValueError("缺少账号 id")
    with CFG_LOCK:
        cfg = dict(CFG)
    before = kx.login_accounts(cfg)
    target = next((a for a in before if str(a.get("id")) == aid), None)
    if target is None:
        raise ValueError("账号不存在")
    cfg["accounts"] = [a for a in before if str(a.get("id")) != aid]
    _persist_config(cfg)
    print(f"[account] 已删除登录账号: id={aid} label={target.get('label')}，"
          f"剩余 {len(cfg['accounts'])} 个（浏览器登录态目录保留）")
    return {"ok": True, "login_accounts": _login_accounts_public(cfg)}


def api_preview(start: str, end: str, only_accounts: list[str] | None = None,
                only_owners: list[str] | None = None) -> dict:
    cfg = _config_snapshot()
    owners = [a for a in kx.login_accounts(cfg) if a.get("enabled", True) is not False]
    if only_owners:
        wanted_owners = {str(x) for x in only_owners}
        owners = [a for a in owners if str(a.get("id")) in wanted_owners]

    sources: list[tuple[dict | None, Path]] = []
    for a in owners:
        d = _owner_export_dir(start, end, a)
        if (d / "accounts.csv").exists():
            sources.append((a, d))
    if not sources:
        legacy = _export_dir(start, end)
        if (legacy / "accounts.csv").exists():
            sources.append((owners[0] if owners else None, legacy))

    accs, prods = [], []
    for owner, d in sources:
        a_rows = _read_csv(d / "accounts.csv")
        p_rows = _read_csv(d / "products.csv")
        if owner is not None:
            a_rows = _tag_accounts(a_rows, owner)
            p_rows = _tag_accounts(p_rows, owner)
        accs.extend(a_rows)
        prods.extend(p_rows)

    if only_accounts:
        accs = [a for a in accs if a["account_id"] in only_accounts]
        prods = [p for p in prods if p["account_id"] in only_accounts]

    # 勾选账户即使没进 CSV（未在月度表 / 拉取失败）也要在预览里出现
    failed: dict[tuple, str] = {}
    for owner, d in sources:
        sp = d / "summary.json"
        if not sp.exists():
            continue
        try:
            sj = json.loads(sp.read_text(encoding="utf-8"))
        except Exception:
            continue
        oid = str(owner.get("id")) if owner else ""
        for f in sj.get("failures") or []:
            failed[(oid, str(f.get("account_id") or ""))] = str(f.get("error") or "拉取失败")

    have = {(str(a.get("owner_id") or ""), str(a["account_id"])) for a in accs}
    if sources:
        cached = _cached_accounts(_config_snapshot())
        want_owners = {str(x) for x in (only_owners or [])}
        want_ids = {str(x) for x in (only_accounts or [])}
        for t in cached:
            oid, aid = str(t.get("owner_id") or ""), str(t.get("id"))
            if want_owners and oid not in want_owners:
                continue
            if want_ids and aid not in want_ids:
                continue
            if (oid, aid) in have:
                continue
            have.add((oid, aid))
            accs.append({
                "date": start if start == end else f"{start}~{end}",
                "owner_id": oid,
                "owner_label": t.get("owner_label") or "",
                "account_id": aid,
                "account_name": t.get("name") or "",
                "cost": 0, "gmv": 0, "order_cnt": 0, "product_count": 0, "roi": 0,
                "failed": True,
                "error": failed.get((oid, aid), "拉取失败"),
            })

    # 每个账户挂上主商品（消耗最大的那个）
    primary: dict[tuple, tuple] = {}
    for pr in prods:
        try:
            score = (float(pr.get("cost") or 0), float(pr.get("gmv") or 0))
        except (TypeError, ValueError):
            score = (0.0, 0.0)
        key = (pr.get("owner_id"), pr["account_id"])
        cur = primary.get(key)
        if cur is None or score > cur[0]:
            primary[key] = (score, pr.get("product_id") or "", pr.get("product_name") or "")
    for a in accs:
        item = primary.get((a.get("owner_id"), a["account_id"]))
        if item:
            a["product_id"], a["product_name"] = item[1], item[2]

    total_cost = round(sum(float(a["cost"]) for a in accs), 2)
    total_gmv = round(sum(float(a["gmv"]) for a in accs), 2)
    return {
        "start": start, "end": end,
        "exists": bool(sources),
        "owners": [{"id": str(a.get("id")), "label": a.get("label") or a.get("id")} for a, _ in sources if a],
        "accounts": accs,
        "products": prods,
        "totals": {
            "cost": total_cost,
            "gmv": total_gmv,
            "roi": round(total_gmv / total_cost, 2) if total_cost else 0,
            "account_rows": len(accs),
            "product_rows": len(prods),
        },
        "files": {
            "dir": str(sources[0][1]) if sources else str(_export_dir(start, end)),
            "accounts_csv": str(sources[0][1] / "accounts.csv") if sources else "",
            "products_csv": str(sources[0][1] / "products.csv") if sources else "",
            "xlsx": str(sources[0][1] / f"明细_{start}_{end}.xlsx") if sources else "",
        },
        "owner_files": [{
            "owner_id": str(a.get("id")),
            "owner_label": str(a.get("label") or a.get("id")),
            "dir": str(d),
            "accounts_csv": str(d / "accounts.csv"),
            "products_csv": str(d / "products.csv"),
            "products_raw_csv": str(d / "products_raw.csv"),
            "products_raw_exists": (d / "products_raw.csv").exists(),
            "xlsx": str(d / f"明细_{start}_{end}.xlsx"),
        } for a, d in sources if a is not None],
    }


APP_EXE = "快手数据工作台.exe"
UPDATE_DIR = kx.BASE / "updates"
UPDATE_PROGRESS = UPDATE_DIR / "update_progress.json"
UPDATER_PS1 = UPDATE_DIR / "apply_update.ps1"
UPDATER_ALIVE = UPDATE_DIR / "updater_alive.txt"


def _ver_tuple(v) -> tuple:
    nums = re.findall(r"\d+", str(v))
    return tuple(int(n) for n in nums[:6]) or (0,)


def _github_release(cfg: dict) -> dict:
    """取 GitHub 最新 Release（私有仓库可用环境变量 GITHUB_TOKEN）"""
    repo = str(cfg.get("update_repo") or "").strip().strip("/")
    if not repo:
        raise RuntimeError("未配置 update_repo（格式：owner/repo）")
    req = urllib.request.Request(
        f"https://api.github.com/repos/{repo}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "kuaishou-roi-workbench"},
    )
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


def _dav_conf(source: str = "webdav") -> dict:
    if source == "webdav_private":
        conf = app_secret.load_private()
        if not conf:
            raise RuntimeError("缺少私有 WebDAV 文件 webdav_private.dat（放到程序目录后重试）")
        return conf
    conf = app_secret.load_webdav()
    if not conf:
        raise RuntimeError("未配置 WebDAV（请在设置里填写地址、用户名、密码并保存）")
    return conf


def _dav_opener(conf: dict):
    from urllib.request import (HTTPDigestAuthHandler, HTTPPasswordMgrWithDefaultRealm,
                                build_opener)
    pm = HTTPPasswordMgrWithDefaultRealm()
    pm.add_password(None, conf["url"], conf.get("username", ""), conf.get("password", ""))
    return build_opener(HTTPDigestAuthHandler(pm))


def _dav_url(conf: dict, path: str) -> str:
    if path.startswith("http://") or path.startswith("https://"):
        return path
    base = conf["url"].rstrip("/")
    return base + (path if path.startswith("/") else "/" + path)


def _dav_release(cfg: dict) -> dict:
    """从 WebDAV 读取 version.json（Digest 认证）"""
    source = str(cfg.get("update_source") or "webdav").lower()
    conf = _dav_conf(source)
    path = str(cfg.get("update_webdav_path") or "/version.json").strip() or "/version.json"
    opener = _dav_opener(conf)
    req = urllib.request.Request(_dav_url(conf, path),
                                 headers={"User-Agent": "kuaishou-roi-workbench"})
    with opener.open(req, timeout=20) as r:
        info = json.loads(r.read().decode("utf-8"))
    dl = str(info.get("download") or info.get("url") or "").strip()
    if dl and not dl.startswith("http"):
        dl = f"{path.rsplit('/', 1)[0]}/{dl}"
    return {
        "version": str(info.get("version") or "").strip().lstrip("vV"),
        "notes": str(info.get("notes") or "")[:2000],
        "download": dl,
        "size": info.get("size"),
        "sha256": str(info.get("sha256") or "").strip(),
    }


def api_webdav_status() -> dict:
    conf = app_secret.load_webdav()
    private = app_secret.private_status()
    return {
        "configured": bool(conf),
        "url": (conf or {}).get("url", ""),
        "username": (conf or {}).get("username", ""),
        "has_password": bool((conf or {}).get("password")),
        "private_ready": private["ready"],
        "private_legacy": private["legacy"],
        "private_unlocked": private["unlocked"],
    }


def api_webdav_save(body: dict) -> dict:
    if body.get("clear"):
        app_secret.clear_webdav()
        print("[webdav] 已清除本机保存的连接信息")
        return {"ok": True, **api_webdav_status()}
    old = app_secret.load_webdav() or {}
    url = str(body.get("url") or "").strip().rstrip("/")
    user = str(body.get("username") or "").strip()
    pwd = str(body.get("password") or "")
    if not re.match(r"^https?://", url):
        raise ValueError("WebDAV 地址必须以 http:// 或 https:// 开头")
    if not user:
        raise ValueError("WebDAV 用户名不能为空")
    if not pwd:
        pwd = str(old.get("password") or "")
    if not pwd:
        raise ValueError("WebDAV 密码不能为空")
    app_secret.save_webdav(url, user, pwd)
    print("[webdav] 连接信息已加密保存到本机")
    return {"ok": True, **api_webdav_status()}


def api_webdav_test(source: str | None = None) -> dict:
    if not source:
        with CFG_LOCK:
            source = str(CFG.get("update_source") or "webdav")
    conf = _dav_conf(source)
    opener = _dav_opener(conf)
    req = urllib.request.Request(conf["url"].rstrip("/") + "/", method="PROPFIND",
                                 headers={"Depth": "0", "User-Agent": "kuaishou-roi-workbench"})
    with opener.open(req, timeout=15) as r:
        print(f"[webdav] 连接测试成功（{source or 'webdav'}）HTTP {r.status}")
        return {"ok": True, "status": r.status}


def _update_info(cfg: dict) -> dict:
    cur = str(cfg.get("version") or "1.0.0")
    src = str(cfg.get("update_source") or "github").strip().lower()
    if src not in ("github", "webdav", "webdav_private"):
        src = "github"
    if src in ("webdav", "webdav_private"):
        rel = _dav_release(cfg)
        latest = rel["version"]
        dl = rel["download"]
        asset = ({"name": dl.rsplit("/", 1)[-1] or "update.zip", "size": rel.get("size"),
                  "url": dl, "sha256": rel.get("sha256") or ""} if dl else None)
        notes = rel["notes"]
        page = ""
        repo = ""
    else:
        rel = _github_release(cfg)
        latest = str(rel.get("tag_name") or "").strip().lstrip("vV")
        assets = rel.get("assets") or []
        zips = [a for a in assets if str(a.get("name", "")).lower().endswith(".zip")]
        # 多个 zip 时优先 Windows 包（常见命名含 win / 快手）
        win_zips = [a for a in zips
                    if any(k in str(a.get("name", "")) for k in ("win", "Win", "快手"))]
        pick = (win_zips or zips)[0] if (win_zips or zips) else None
        asset = None if not pick else {
            "name": pick.get("name"),
            "size": pick.get("size"),
            "url": pick.get("browser_download_url"),
            # GitHub Release 资产自带 digest（形如 sha256:...）
            "sha256": pick.get("digest") or "",
        }
        notes = (rel.get("body") or "")[:2000]
        page = rel.get("html_url") or ""
        repo = str(cfg.get("update_repo") or "")

    return {
        "source": src,
        "repo": repo,
        "current": cur,
        "latest": latest,
        "has_update": _ver_tuple(latest) > _ver_tuple(cur),
        "notes": notes,
        "page_url": page,
        "asset": asset,
    }


def api_update_check() -> dict:
    with CFG_LOCK:
        cfg = dict(CFG)
    info = _update_info(cfg)
    print(f"[update] 检查更新: source={info['source']} current={info['current']} "
          f"latest={info['latest']} has_update={info['has_update']}")
    return info


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _download(cfg: dict, url: str, dst: Path, logger, on_progress=None, total_hint=None) -> dict:
    """下载升级包；on_progress 回调带 downloaded/total/speed/percent（用于网页进度条）"""
    logger("[update] 开始下载升级包…")
    src = str(cfg.get("update_source") or "github").strip().lower()
    if src in ("webdav", "webdav_private"):
        conf = _dav_conf(src)
        opener = _dav_opener(conf)
        req = urllib.request.Request(_dav_url(conf, url),
                                     headers={"User-Agent": "kuaishou-roi-workbench"})
        resp = opener.open(req, timeout=120)
    else:
        token = os.environ.get("GITHUB_TOKEN", "").strip()
        req = urllib.request.Request(url, headers={"User-Agent": "kuaishou-roi-workbench"})
        if token:
            req.add_header("Authorization", f"Bearer {token}")
            req.add_header("Accept", "application/octet-stream")
        resp = urllib.request.urlopen(req, timeout=120)

    def emit(**kw):
        if on_progress:
            on_progress(**kw)

    with resp, dst.open("wb") as f:
        try:
            total = int(resp.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            total = 0
        if total <= 0:
            try:
                total = int(total_hint or 0)
            except (TypeError, ValueError):
                total = 0
        done = 0
        speed = 0.0
        started = time.time()
        last_t, last_done = started, 0
        emit(downloaded=0, total=total, speed=0, percent=0 if total else None, elapsed=0)
        while True:
            chunk = resp.read(64 * 1024)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            now = time.time()
            if now - last_t >= 0.4:
                inst = (done - last_done) / (now - last_t)
                speed = inst if speed <= 0 else speed * 0.7 + inst * 0.3
                last_t, last_done = now, done
                emit(downloaded=done, total=total, speed=round(speed, 1),
                     percent=round(done * 100 / total, 1) if total else None,
                     elapsed=round(now - started, 1))
    elapsed = max(0.001, time.time() - started)
    avg = done / elapsed
    emit(downloaded=done, total=total or done, speed=round(speed or avg, 1),
         percent=100 if total else None, elapsed=round(elapsed, 1), done=True)
    logger(f"[update] 下载完成 {dst.stat().st_size / 1024 / 1024:.1f} MB"
           f"（{avg / 1024 / 1024:.1f} MB/s，用时 {elapsed:.1f}s）")
    return {"downloaded": done, "total": total, "speed": round(speed or avg, 1),
            "avg_speed": round(avg, 1), "elapsed": round(elapsed, 1)}


def _write_update_progress(**fields):
    """把升级进度写入 updates/update_progress.json，供独立升级窗口实时读取"""
    try:
        UPDATE_DIR.mkdir(parents=True, exist_ok=True)
        data = {k: v for k, v in fields.items() if v is not None}
        data["ts"] = round(time.time(), 2)
        tmp = UPDATE_PROGRESS.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, UPDATE_PROGRESS)
    except Exception as e:  # noqa: BLE001
        print(f"[update] 写进度文件失败: {type(e).__name__}: {e}")


def _clean_update_dir(keep: set[str] | None = None):
    """清空 updates/ 下的旧包与解压目录，保留升级窗口正在使用的文件"""
    if not UPDATE_DIR.exists():
        return
    keep = keep or set()
    for p in UPDATE_DIR.iterdir():
        if p.name in keep:
            continue
        try:
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            else:
                p.unlink(missing_ok=True)
        except Exception:
            pass


UPDATER_PS1_TEXT = r"""param(
  [string]$App,
  [string]$New,
  [int]$AppPid,
  [string]$ProgressFile,
  [string]$LogPath,
  [string]$Exe = '快手数据工作台.exe',
  [string]$AliveFile,
  [switch]$NoGui
)

$ErrorActionPreference = 'Continue'
$form = $null
$lblTitle = $null
$lblStage = $null
$lblDetail = $null
$lblStatus = $null
$bar = $null

function Write-Log([string]$msg) {
  try {
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
  } catch { }
}

function Read-Progress {
  if (-not (Test-Path -LiteralPath $ProgressFile)) { return $null }
  try {
    $raw = Get-Content -LiteralPath $ProgressFile -Raw -Encoding UTF8 -ErrorAction Stop
    if (-not $raw) { return $null }
    return ($raw | ConvertFrom-Json)
  } catch { return $null }
}

function Format-Size([double]$n) {
  if ($n -ge 1073741824) { return ('{0:N2} GB' -f ($n / 1073741824.0)) }
  if ($n -ge 1048576) { return ('{0:N1} MB' -f ($n / 1048576.0)) }
  if ($n -ge 1024) { return ('{0:N0} KB' -f ($n / 1024.0)) }
  return ('{0:N0} B' -f $n)
}

function Set-Stage([string]$stage, [string]$detail, [int]$permille) {
  if ($null -eq $form) { return }
  try {
    if ($stage) { $lblStage.Text = $stage }
    if ($null -ne $detail) { $lblDetail.Text = $detail }
    if ($permille -ge 0) {
      if ($permille -gt 1000) { $permille = 1000 }
      $bar.Value = $permille
    }
    $form.Refresh()
  } catch { }
}

function Update-Ui($p) {
  if ($null -eq $form -or $null -eq $p) { return }
  $phase = [string]$p.phase
  $text = [string]$p.text
  if ($phase -eq 'starting') {
    Set-Stage '1/4 准备升级' '正在准备升级…' 20
  } elseif ($phase -eq 'checking') {
    if ($text) { Set-Stage '1/4 检查更新' $text 60 } else { Set-Stage '1/4 检查更新' '正在检查最新版本…' 60 }
  } elseif ($phase -eq 'downloading') {
    $pct = -1.0
    if ($null -ne $p.percent) { $pct = [double]$p.percent }
    $dl = [double]$p.downloaded
    $tot = [double]$p.total
    $sp = [double]$p.speed
    $d = '已下载 ' + (Format-Size $dl)
    if ($tot -gt 0) { $d = $d + ' / ' + (Format-Size $tot) }
    if ($sp -gt 0) { $d = $d + ' · ' + (Format-Size $sp) + '/s' }
    $perm = 0
    if ($pct -ge 0) {
      $rp = [int][Math]::Round($pct)
      $d = $d + ' · ' + [string]$rp + '%'
      $perm = [int][Math]::Round($pct * 7.5)
    }
    Set-Stage '2/4 下载升级包' $d $perm
  } elseif ($phase -eq 'extracting') {
    $pct = 0.0
    if ($null -ne $p.percent) { $pct = [double]$p.percent }
    $ex = [int]$p.extracted
    $tf = [int]$p.total_files
    $rp = [int][Math]::Round($pct)
    $d = '已解压 ' + [string]$ex + ' / ' + [string]$tf + ' 个文件 · ' + [string]$rp + '%'
    Set-Stage '3/4 解压升级包' $d ([int](750 + $pct * 2.5))
  } elseif ($phase -eq 'ready') {
    if ($text) { Set-Stage '4/4 准备替换' $text 1000 } else { Set-Stage '4/4 准备替换' '升级包已就绪…' 1000 }
  } elseif ($phase -eq 'restarting') {
    if ($text) { Set-Stage '4/4 重启替换' $text 1000 } else { Set-Stage '4/4 重启替换' '程序即将退出并重启…' 1000 }
  } elseif ($phase -eq 'failed') {
    Set-Stage '升级失败' ([string]$p.error) -1
  }
}

function Show-Failure([string]$msg) {
  Write-Log ('失败: ' + $msg)
  if ($null -ne $form) {
    try {
      Set-Stage '升级失败' $msg -1
      $lblStatus.Text = '日志: ' + $LogPath
      $form.Refresh()
    } catch { }
    $until = (Get-Date).AddSeconds(25)
    while ((Get-Date) -lt $until) {
      [System.Windows.Forms.Application]::DoEvents()
      Start-Sleep -Milliseconds 200
    }
  }
  exit 1
}

function Remove-WithRetry([string]$path) {
  for ($i = 0; $i -lt 20; $i++) {
    try {
      if (-not (Test-Path -LiteralPath $path)) { return $true }
      Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
      return $true
    } catch {
      Start-Sleep -Milliseconds 600
    }
  }
  return (-not (Test-Path -LiteralPath $path))
}

function Move-WithRetry([string]$src, [string]$dst) {
  for ($i = 0; $i -lt 20; $i++) {
    try {
      Move-Item -LiteralPath $src -Destination $dst -ErrorAction Stop
      return $true
    } catch {
      Start-Sleep -Milliseconds 600
    }
  }
  return $false
}

function Copy-WithRetry([string]$src, [string]$dst) {
  for ($i = 0; $i -lt 20; $i++) {
    try {
      Copy-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
      return $true
    } catch {
      Start-Sleep -Milliseconds 600
    }
  }
  return $false
}

function Resolve-AppRoot([string]$root) {
  if (Test-Path -LiteralPath (Join-Path $root $Exe)) { return $root }
  $dirs = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)
  foreach ($d in $dirs) {
    if (Test-Path -LiteralPath (Join-Path $d.FullName $Exe)) { return $d.FullName }
  }
  return $root
}

Write-Log ('升级器启动 pid=' + [string]$AppPid + ' 安装目录=' + $App)
if ($AliveFile) {
  try { Set-Content -LiteralPath $AliveFile -Value ('pid=' + [string]$PID) -Encoding UTF8 } catch { }
}

if (-not $NoGui) {
  try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '快手数据工作台 · 升级程序'
    $form.ClientSize = New-Object System.Drawing.Size(440, 170)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = '正在升级快手数据工作台'
    $lblTitle.Location = New-Object System.Drawing.Point(18, 14)
    $lblTitle.Size = New-Object System.Drawing.Size(404, 26)
    $lblTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblTitle)

    $lblStage = New-Object System.Windows.Forms.Label
    $lblStage.Text = '准备中…'
    $lblStage.Location = New-Object System.Drawing.Point(18, 50)
    $lblStage.Size = New-Object System.Drawing.Size(404, 22)
    $form.Controls.Add($lblStage)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(18, 78)
    $bar.Size = New-Object System.Drawing.Size(404, 14)
    $bar.Minimum = 0
    $bar.Maximum = 1000
    $bar.Value = 0
    $form.Controls.Add($bar)

    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = '等待主程序下载升级包…'
    $lblDetail.Location = New-Object System.Drawing.Point(18, 100)
    $lblDetail.Size = New-Object System.Drawing.Size(404, 22)
    $form.Controls.Add($lblDetail)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = '升级期间请勿关闭此窗口'
    $lblStatus.Location = New-Object System.Drawing.Point(18, 126)
    $lblStatus.Size = New-Object System.Drawing.Size(404, 22)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(107, 114, 128)
    $form.Controls.Add($lblStatus)

    $form.Show()
    $form.Activate()
    [System.Windows.Forms.Application]::DoEvents()
    Write-Log '升级窗口已显示'
  } catch {
    Write-Log ('升级窗口创建失败: ' + $_.Exception.Message)
    $form = $null
  }
}

$deadline = (Get-Date).AddMinutes(60)
$p = $null
while ((Get-Date) -lt $deadline) {
  $p = Read-Progress
  if ($null -ne $p) { Update-Ui $p }
  if ($null -ne $p -and ([string]$p.phase -eq 'ready' -or [string]$p.phase -eq 'restarting')) { break }
  if ($null -ne $p -and [string]$p.phase -eq 'failed') { Show-Failure ([string]$p.error) }
  Start-Sleep -Milliseconds 250
  if ($null -ne $form) { [System.Windows.Forms.Application]::DoEvents() }
}

if ($null -eq $p -or -not ([string]$p.phase -eq 'ready' -or [string]$p.phase -eq 'restarting')) {
  Show-Failure '等待升级包超时（60 分钟）'
}

Write-Log ('升级包就绪，版本=' + [string]$p.latest)
Set-Stage '4/4 等待主程序退出' '等待旧进程结束后替换文件…' 1000

$alive = $true
$waitUntil = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $waitUntil) {
  $alive = $false
  try { $alive = [bool](Get-Process -Id $AppPid -ErrorAction SilentlyContinue) } catch { $alive = $false }
  if (-not $alive) { break }
  Start-Sleep -Milliseconds 500
  if ($null -ne $form) { [System.Windows.Forms.Application]::DoEvents() }
}
if ($alive) {
  Write-Log '等待超时：主程序仍在运行，放弃替换'
  Show-Failure '主程序没有退出，替换已中止（升级包已下载，可稍后重试）'
}
Write-Log '主程序已退出'

$New = Resolve-AppRoot $New
$internal = Join-Path $App '_internal'
$oldInternal = Join-Path $App '_internal.old'
$newInternal = Join-Path $New '_internal'
$appExe = Join-Path $App $Exe
$newExe = Join-Path $New $Exe

if (-not (Test-Path -LiteralPath $newInternal) -or -not (Test-Path -LiteralPath $newExe)) {
  Show-Failure '升级包不完整（缺少新的 _internal 或 exe）'
}

Set-Stage '4/4 替换程序文件' '正在替换程序文件…' 1000

if (Test-Path -LiteralPath $oldInternal) {
  Write-Log '清理残留 _internal.old'
  if (-not (Remove-WithRetry $oldInternal)) { Show-Failure '无法删除 _internal.old（可能仍有进程占用）' }
}

if (Test-Path -LiteralPath $internal) {
  Write-Log '重命名 _internal -> _internal.old'
  if (-not (Move-WithRetry $internal $oldInternal)) { Show-Failure '无法重命名 _internal（可能仍有进程占用）' }
} else {
  Write-Log '未发现旧 _internal（按全新安装处理）'
}

Write-Log '写入新的 _internal'
if (-not (Move-WithRetry $newInternal $internal)) {
  if ((Test-Path -LiteralPath $oldInternal) -and -not (Test-Path -LiteralPath $internal)) {
    Move-WithRetry $oldInternal $internal | Out-Null
  }
  Show-Failure '无法写入新 _internal，已回滚旧版本'
}

Write-Log '替换主程序 exe'
if (-not (Copy-WithRetry $newExe $appExe)) {
  Write-Log 'exe 替换失败，回滚 _internal'
  Remove-WithRetry $internal | Out-Null
  if (Test-Path -LiteralPath $oldInternal) { Move-WithRetry $oldInternal $internal | Out-Null }
  Show-Failure '无法替换主程序 exe（可能仍被占用），已回滚旧版本'
}

Write-Log '替换完成'
Set-Stage '升级完成' '正在启动新版本…' 1000
try {
  Start-Process -FilePath $appExe -WorkingDirectory $App
  Write-Log '已启动新版本'
} catch {
  Write-Log ('启动新版本失败: ' + $_.Exception.Message)
}

Start-Sleep -Seconds 2
if ($null -ne $form) { try { $form.Close() } catch { } }
exit 0
"""


def _write_updater() -> Path:
    """生成独立升级窗口脚本（PowerShell + WinForms）：显示阶段与进度，并完成替换"""
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        UPDATER_PS1.write_text(UPDATER_PS1_TEXT, encoding="utf-8-sig")
    except Exception:
        UPDATER_PS1.write_text(UPDATER_PS1_TEXT, encoding="utf-8")
    return UPDATER_PS1


def run_updater(ps1: Path, new_root: Path):
    """启动独立升级窗口（与主程序脱离，主程序退出后继续替换并重启）"""
    log_path = app_runtime.LOG_DIR / f"update-{time.strftime('%Y%m%d')}.log"
    try:
        UPDATER_ALIVE.unlink(missing_ok=True)
    except Exception:
        pass
    args = [
        "powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
        "-WindowStyle", "Hidden", "-File", str(ps1),
        "-App", str(kx.BASE), "-New", str(new_root),
        "-AppPid", str(os.getpid()), "-ProgressFile", str(UPDATE_PROGRESS),
        "-LogPath", str(log_path), "-Exe", APP_EXE, "-AliveFile", str(UPDATER_ALIVE),
    ]
    flags = 0x08000000  # CREATE_NO_WINDOW：DETACHED_PROCESS 会让 powershell 直接退出不执行脚本
    subprocess.Popen(args, creationflags=flags, close_fds=True)


def _wait_updater_alive(timeout: float = 10.0) -> bool:
    """等待升级窗口写下启动标记，确认辅助进程真的起来了（否则不退出主程序）"""
    end = time.time() + timeout
    while time.time() < end:
        try:
            if UPDATER_ALIVE.exists():
                return True
        except OSError:
            pass
        time.sleep(0.5)
    return False


def stage_update(logger=print, progress=None) -> dict:
    """检查 + 下载 + 解压升级包（替换由独立升级窗口负责）"""
    def emit(**kw):
        if progress:
            progress(kw)

    with CFG_LOCK:
        cfg = dict(CFG)
    if not (kx.BASE / "_internal").exists():
        raise RuntimeError("当前不是打包运行（缺少 _internal），无法自动升级")
    emit(phase="checking", text="正在检查最新版本…")
    info = _update_info(cfg)
    if not info["has_update"]:
        raise RuntimeError(f"已是最新版本（v{info['current']}）")
    asset = info.get("asset")
    if not asset:
        raise RuntimeError("该 Release 未上传 zip 升级包")

    _clean_update_dir({UPDATER_PS1.name, UPDATE_PROGRESS.name, UPDATER_ALIVE.name})
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)

    zip_path = UPDATE_DIR / Path(str(asset.get("name") or "update.zip")).name
    try:
        expected = int(asset.get("size") or 0)
    except (TypeError, ValueError):
        expected = 0
    emit(phase="downloading", downloaded=0, total=expected, speed=0,
         percent=0 if expected else None, elapsed=0)
    dl = _download(cfg, asset["url"], zip_path, logger,
                   on_progress=lambda **kw: emit(phase="downloading", **kw),
                   total_hint=expected)

    expected_sha = str(asset.get("sha256") or "").strip()
    if expected_sha:
        expected_sha = expected_sha.split(":")[-1].strip().lower()
        actual_sha = _sha256_file(zip_path)
        if actual_sha != expected_sha:
            raise RuntimeError(
                f"升级包 SHA256 校验失败（期望 {expected_sha[:12]}…，实际 {actual_sha[:12]}…）")
        logger("[update] SHA256 校验通过")
    else:
        logger("[update] 警告: 升级包未提供 sha256，跳过完整性校验")

    new_dir = UPDATE_DIR / "new"
    with zipfile.ZipFile(zip_path) as z:
        members = z.infolist()
        total_files = len(members) or 1
        emit(phase="extracting", extracted=0, total_files=total_files, percent=0,
             downloaded=dl["downloaded"], total=dl["total"])
        for i, m in enumerate(members, 1):
            z.extract(m, new_dir)
            if i % 40 == 0 or i == total_files:
                emit(phase="extracting", extracted=i, total_files=total_files,
                     percent=round(i * 100 / total_files, 1),
                     downloaded=dl["downloaded"], total=dl["total"])
    root = new_dir
    if not (root / APP_EXE).exists():
        subs = [p for p in new_dir.iterdir() if p.is_dir()]
        if len(subs) == 1 and (subs[0] / APP_EXE).exists():
            root = subs[0]
    if not (root / APP_EXE).exists() or not (root / "_internal").exists():
        raise RuntimeError("升级包里缺少 exe 或 _internal，已中止")

    emit(phase="ready", percent=100, text="升级包已就绪，准备重启…")
    logger(f"[update] 升级包就绪（v{info['latest']}），即将重启替换")
    return {
        "current": info["current"],
        "latest": info["latest"],
        "install_dir": str(kx.BASE),
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "KuaishouRoi/1.0"

    def log_message(self, fmt, *args):  # 安静一点
        pass

    # -- helpers -------------------------------------------------
    def _send(self, code: int, body: bytes, ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        if n > 2_000_000:
            raise ValueError("请求体过大")
        try:
            return json.loads(self.rfile.read(n).decode("utf-8") or "{}")
        except Exception:
            return {}

    @staticmethod
    def _local_host(value: str | None) -> str:
        if not value:
            return ""
        try:
            v = str(value).strip()
            if "://" in v:
                return (urlparse(v).hostname or "").lower()
            return (urlparse("//" + v).hostname or "").lower()
        except Exception:  # noqa: BLE001
            return ""

    def _request_allowed(self) -> bool:
        """只接受本机回环地址发起的请求，阻断 DNS rebinding / 跨站请求。"""
        allowed = {"127.0.0.1", "localhost", "::1"}
        host = self._local_host(self.headers.get("Host"))
        if host not in allowed:
            return False
        origin = self.headers.get("Origin")
        if origin and self._local_host(origin) not in allowed:
            return False
        referer = self.headers.get("Referer")
        if referer and self._local_host(referer) not in allowed:
            return False
        return True

    # -- routes --------------------------------------------------
    def do_GET(self):
        if not self._request_allowed():
            return self._json({"error": "forbidden"}, 403)
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path in ("/", "/index.html"):
                html = (WEB_DIR / "index.html").read_bytes()
                return self._send(200, html, "text/html; charset=utf-8")
            if u.path == "/api/state":
                return self._json(api_state())
            if u.path == "/api/config":
                return self._json(api_config())
            if u.path == "/api/preview":
                start, end = q.get("start"), q.get("end")
                if not start or not end:
                    return self._json({"error": "缺少 start/end"}, 400)
                ids = [x for x in (q.get("accounts") or "").split(",") if x]
                owners = [x for x in (q.get("owners") or "").split(",") if x]
                return self._json(api_preview(start, end, ids or None, owners or None))
            if u.path == "/api/open-file":
                p = Path(q.get("path", "")).resolve()
                data_root = kx.DATA_DIR.resolve()
                excel_paths = set()
                for a in kx.login_accounts(_config_snapshot()):
                    if a.get("excel_path"):
                        excel_paths.add(kx.resolve_path(str(a["excel_path"])).resolve())
                try:
                    under_data = p.is_relative_to(data_root)
                except AttributeError:  # Python < 3.9
                    under_data = str(p).startswith(str(data_root) + os.sep)
                allowed = under_data or p in excel_paths
                if not allowed or not p.is_file():
                    return self._json({"error": "路径不允许或文件不存在"}, 403)
                os.startfile(str(p))  # noqa: S606
                return self._json({"ok": True})
            if u.path == "/api/webdav":
                return self._json(api_webdav_status())
            if u.path == "/api/update/check":
                try:
                    return self._json(api_update_check())
                except app_secret.PrivatePassphraseRequired as e:
                    return self._json({"error": str(e), "need_passphrase": True}, 401)
                except Exception as e:  # noqa: BLE001
                    print(f"[web] 检查更新失败: {type(e).__name__}: {e}")
                    return self._json({"error": str(e)}, 502)
            if u.path == "/api/job":
                jid = q.get("id", "")
                with JOBS_LOCK:
                    job = JOBS.get(jid)
                    if not job:
                        return self._json({"error": "job 不存在"}, 404)
                    return self._json(dict(job))
            return self._json({"error": "not found"}, 404)
        except Exception as e:  # noqa: BLE001
            print(f"[web] {self.command} {self.path} 失败: {type(e).__name__}: {e}")
            traceback.print_exc()
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)

    def do_POST(self):
        if not self._request_allowed():
            return self._json({"error": "forbidden"}, 403)
        u = urlparse(self.path)
        try:
            if u.path == "/api/export":
                body = self._body()
                start, end = body.get("start"), body.get("end")
                if not start or not end:
                    return self._json({"error": "缺少 start/end"}, 400)
                try:
                    owners, mapping = _selected_owner_map(body)
                except RuntimeError as e:
                    return self._json({"error": str(e)}, 400)
                jid = _job_new("export")

                def work():
                    results, failures = [], []
                    for acct in owners:
                        label = acct.get("label") or acct.get("id")
                        _job_log(jid, f"[export] 账号「{label}」开始…")
                        try:
                            ecfg = kx.account_cfg(_config_snapshot(), acct)
                            outdir = _export_owner_dir(start, end, str(acct.get("id")))
                            s = export_full.run_export(ecfg, start, end, mapping.get(str(acct.get("id"))),
                                                       outdir=outdir,
                                                       logger=lambda m: _job_log(jid, m))
                            results.append({"owner_id": str(acct.get("id")), "owner_label": label, "summary": s})
                        except Exception as e:  # noqa: BLE001
                            _job_log(jid, f"[export] 账号「{label}」失败: {type(e).__name__}: {e}")
                            failures.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                             "error": str(e)[:300]})
                    if not results:
                        _job_done(jid, error=f"全部账号导出失败（{len(failures)} 个）")
                    else:
                        _job_done(jid, {"owners": results, "failures": failures})

                threading.Thread(target=work, daemon=True).start()
                return self._json({"job_id": jid})

            if u.path == "/api/login-accounts/save":
                try:
                    return self._json(api_save_login_account(self._body()))
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)

            if u.path == "/api/login-accounts/delete":
                try:
                    return self._json(api_delete_login_account(self._body()))
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)

            if u.path == "/api/accounts":
                body = self._body()
                jid = _job_new("accounts")

                def work_accounts():
                    cfg = _config_snapshot()
                    try:
                        owners = _enabled_login_accounts(cfg, body.get("owners") or None)
                    except RuntimeError as e:
                        _job_done(jid, error=str(e))
                        return
                    merged, failures = [], []
                    for acct in owners:
                        label = acct.get("label") or acct.get("id")
                        _job_log(jid, f"[accounts] 账号「{label}」获取账户列表…")
                        try:
                            ecfg = kx.account_cfg(cfg, acct)
                            accs = kx.fetch_accounts(ecfg, logger=lambda m: _job_log(jid, m))
                            merged.extend(_tag_accounts(accs, acct))
                        except Exception as e:  # noqa: BLE001
                            _job_log(jid, f"[accounts] 账号「{label}」失败: {type(e).__name__}: {e}")
                            failures.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                             "error": str(e)[:300]})
                    if merged:
                        _job_done(jid, {"accounts": merged, "failures": failures})
                    else:
                        _job_done(jid, error="全部账号获取账户列表失败")

                threading.Thread(target=work_accounts, daemon=True).start()
                return self._json({"job_id": jid})

            if u.path == "/api/config":
                try:
                    return self._json(api_save_config(self._body()))
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)

            if u.path == "/api/pick-excel":
                body = self._body() or {}
                try:
                    path = pick_excel_file(str(body.get("current") or ""))
                except Exception as e:  # noqa: BLE001
                    return self._json({"error": f"打开文件选择框失败: {type(e).__name__}: {e}"}, 500)
                return self._json({"path": path})

            if u.path == "/api/realtime":
                body = self._body()
                cfg = _config_snapshot()
                try:
                    owners = _enabled_login_accounts(cfg, body.get("owners") or None)
                except RuntimeError as e:
                    return self._json({"error": str(e)}, 400)
                jid = _job_new("realtime")

                def work_realtime():
                    try:
                        combined = collect_realtime(cfg, owners, logger=lambda m: _job_log(jid, m))
                        _job_done(jid, combined)
                    except Exception as e:  # noqa: BLE001
                        _job_log(jid, f"异常: {type(e).__name__}: {e}")
                        _job_done(jid, error=f"{type(e).__name__}: {e}")
                        traceback.print_exc()

                threading.Thread(target=work_realtime, daemon=True).start()
                return self._json({"job_id": jid})

            if u.path == "/api/webdav":
                try:
                    return self._json(api_webdav_save(self._body()))
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)
            if u.path == "/api/webdav/test":
                try:
                    try:
                        src = (self._body() or {}).get("source") or None
                    except Exception:  # noqa: BLE001
                        src = None
                    return self._json(api_webdav_test(src))
                except app_secret.PrivatePassphraseRequired as e:
                    return self._json({"error": str(e), "need_passphrase": True}, 401)
                except Exception as e:  # noqa: BLE001
                    print(f"[web] WebDAV 连接测试失败: {type(e).__name__}: {e}")
                    return self._json({"error": str(e)}, 502)
            if u.path == "/api/webdav/private":
                try:
                    body = self._body() or {}
                    action = str(body.get("action") or "generate")
                    passphrase = str(body.get("passphrase") or "")
                    if action == "clear":
                        app_secret.clear_private()
                        print("[webdav] 已删除私有文件 webdav_private.dat")
                    elif action == "unlock":
                        app_secret.unlock_private(passphrase)
                        print("[webdav] 私有 WebDAV 已用口令解锁（仅保存在内存）")
                    elif action == "lock":
                        app_secret.lock_private()
                        print("[webdav] 私有 WebDAV 已锁定，内存口令已清除")
                    else:
                        if not passphrase:
                            raise ValueError("生成私有文件需要设置口令")
                        app_secret.export_private(passphrase)
                        print("[webdav] 已生成/更新私有文件 webdav_private.dat（口令加密，可跨设备）")
                    return self._json(api_webdav_status())
                except ValueError as e:
                    return self._json({"error": str(e)}, 400)
                except Exception as e:  # noqa: BLE001
                    return self._json({"error": str(e)}, 400)
            if u.path == "/api/update/apply":
                jid = _job_new("update")
                print("[update] 收到升级请求")

                def work_update():
                    def on_prog(p):
                        _job_progress(jid, **p)
                        _write_update_progress(**p)

                    try:
                        _write_update_progress(phase="starting", percent=0, text="正在准备升级…")
                        run_updater(_write_updater(), UPDATE_DIR / "new")
                        if not _wait_updater_alive():
                            raise RuntimeError("升级窗口没有启动成功（可能被安全软件拦截），已取消本次升级")
                        _job_log(jid, "[update] 升级窗口已启动，开始下载升级包…")
                        res = stage_update(logger=lambda m: _job_log(jid, m), progress=on_prog)
                        _write_update_progress(phase="restarting", percent=100, latest=res["latest"],
                                               text="升级包已就绪，程序即将退出并重启…")
                        _job_log(jid, "[update] 替换程序已就绪，程序将在约 3 秒后退出并重启")
                        _job_done(jid, dict(res))
                        app_runtime.quit_app(3.0)
                    except Exception as e:  # noqa: BLE001
                        _write_update_progress(phase="failed", text="升级失败",
                                               error=f"{type(e).__name__}: {e}")
                        _job_log(jid, f"升级失败: {type(e).__name__}: {e}")
                        _job_done(jid, error=f"{type(e).__name__}: {e}")
                        traceback.print_exc()

                threading.Thread(target=work_update, daemon=True).start()
                return self._json({"job_id": jid})

            if u.path == "/api/merge":
                body = self._body()
                start, end = body.get("start"), body.get("end")
                if not start or not end:
                    return self._json({"error": "缺少 start/end"}, 400)
                try:
                    owners, mapping = _selected_owner_map(body)
                except RuntimeError as e:
                    return self._json({"error": str(e)}, 400)
                cfg = _config_snapshot()
                started_ts = time.time()
                started_at = time.strftime("%Y-%m-%d %H:%M:%S")
                results = []
                for acct in owners:
                    label = acct.get("label") or acct.get("id")
                    ecfg = kx.account_cfg(cfg, acct)
                    try:
                        res = merge_excel.run_merge(
                            ecfg, start, end, mapping.get(str(acct.get("id"))),
                            export_dir=_owner_export_dir(start, end, acct),
                            backup=bool(cfg.get("auto_backup", True)))
                        print(f"[merge] 账号「{label}」完成: written={res.get('written')} "
                              f"saved={res.get('saved')} locked={res.get('locked')}")
                        results.append(dict(res, owner_id=str(acct.get("id")), owner_label=label))
                    except Exception as e:  # noqa: BLE001
                        print(f"[merge] 账号「{label}」失败: {type(e).__name__}: {e}")
                        traceback.print_exc()
                        results.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                        "error": f"{type(e).__name__}: {e}", "saved": False,
                                        "locked": False, "written": 0, "checks": []})
                return self._json({
                    "results": results,
                    "saved": all(r.get("saved") for r in results) if results else False,
                    "locked": any(r.get("locked") for r in results),
                    "started_at": started_at,
                    "finished_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "elapsed": round(time.time() - started_ts, 1),
                })

            return self._json({"error": "not found"}, 404)
        except Exception as e:  # noqa: BLE001
            print(f"[web] {self.command} {self.path} 失败: {type(e).__name__}: {e}")
            traceback.print_exc()
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)


def _already_running(url: str) -> bool:
    """单实例检查：Windows 用命名互斥量，其他平台退回接口探测"""
    global _SINGLE_INSTANCE_MUTEX
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.CreateMutexW.restype = ctypes.c_void_p
            kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
            handle = kernel32.CreateMutexW(None, False, "Local\\KuaishouDataWorkbench")
            if not handle:
                return False
            _SINGLE_INSTANCE_MUTEX = handle  # 进程存活期间保持句柄，退出自动释放
            return ctypes.get_last_error() == 183  # ERROR_ALREADY_EXISTS
        except Exception:
            return False
    try:
        with urllib.request.urlopen(url + "api/state", timeout=1.5) as r:
            return 200 <= r.status < 300
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host")
    ap.add_argument("--port", type=int)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--no-tray", action="store_true", help="不显示托盘图标（前台调试用）")
    args = ap.parse_args()

    web = CFG.get("web") or {}
    host = args.host or web.get("host", "127.0.0.1")
    port = args.port or int(web.get("port", 8765))
    url = f"http://{host}:{port}/"

    if _already_running(url):
        print("[web] 已有实例在运行，改为打开浏览器")
        if not args.no_browser:
            webbrowser.open(url)
        return

    try:
        srv = ThreadingHTTPServer((host, port), Handler)
    except OSError as e:
        app_runtime.show_error("快手数据工作台", f"端口 {port} 无法监听（可能已经有一份在运行）\n\n{e}")
        return

    print(f"[web] 数据工作台已启动: {url}")
    print(f"[web] 日志目录: {app_runtime.LOG_DIR}")
    print(f"[web] 登录账号 {len(kx.login_accounts(CFG))} 个"
          + ("" if kx.login_accounts(CFG) else "（请先在「账号管理」添加）"))
    app_runtime.cleanup_old_version()
    threading.Thread(target=realtime_scheduler_loop, daemon=True, name="realtime-auto").start()
    print("[auto] 后台自动刷新调度已启动（按配置开关/间隔/时段）")
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    if args.no_tray:
        try:
            print("[web] Ctrl+C 停止")
            srv.serve_forever()
        except KeyboardInterrupt:
            print("\n[web] 已停止")
        finally:
            srv.server_close()
        return

    # 托盘常驻：HTTP 服务放后台线程，主线程跑托盘消息循环
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        if not app_runtime.run_tray(url, srv.shutdown):
            print("[web] 托盘不可用，改为前台常驻")
            srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()
        print("[web] 已退出")



if __name__ == "__main__":
    main()
