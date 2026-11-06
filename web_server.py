# -*- coding: utf-8 -*-
"""本地数据工作台：选择日期/账户 -> 拉取全量 -> 预览 -> 合并进月度表

启动:
  py web_server.py            # 默认 http://127.0.0.1:8765
  py web_server.py --port 9000 --no-browser
"""
from __future__ import annotations

import argparse
import csv
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
JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def _job_new(kind: str) -> str:
    jid = uuid.uuid4().hex[:12]
    with JOBS_LOCK:
        JOBS[jid] = {"id": jid, "kind": kind, "status": "running", "logs": [], "result": None, "error": None}
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


def _export_dir(start: str, end: str) -> Path:
    return kx.DATA_DIR / "exports" / f"{start}_{end}"


def _read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _excel_accounts(path: Path) -> list[str]:
    """读月度表第 3 行账户ID"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        ws = wb["消耗成交"] if "消耗成交" in wb.sheetnames else wb[wb.sheetnames[0]]
        ids = []
        for row in ws.iter_rows(min_row=3, max_row=3, min_col=2, max_col=28):
            for c in row:
                if c.value is not None and str(c.value).strip():
                    ids.append(str(c.value).strip())
        wb.close()
        return ids
    except Exception:
        return []


def _excel_products(path: Path) -> list[dict]:
    """读月度表第 1~3 行：商品名 / 商品ID / 所属账户ID"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        ws = wb["消耗成交"] if "消耗成交" in wb.sheetnames else wb[wb.sheetnames[0]]
        out = []
        for col in merge_excel.COST_COLS:
            aid = merge_excel._cell_id(ws.cell(row=3, column=col).value)
            if not aid:
                continue
            out.append({
                "account_id": aid,
                "product_id": merge_excel._cell_id(ws.cell(row=2, column=col).value),
                "product_name": str(ws.cell(row=1, column=col).value or "").strip(),
            })
        wb.close()
        return out
    except Exception:
        return []


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
        "realtime": kx.load_realtime_cache(),
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
               "realtime_window_start", "realtime_window_end")


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


def api_save_config(body: dict) -> dict:
    """在网页上修改全局配置并写回 config.json（登录账号相关在「账号管理」里保存）"""
    with CFG_LOCK:
        cfg = dict(CFG)
    if "date_offset_days" in body:
        cfg["date_offset_days"] = _as_int(body["date_offset_days"], "日期偏移", 0, 30)
    if "page_size" in body:
        cfg["page_size"] = _as_int(body["page_size"], "page_size", 10, 100)
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
    if isinstance(body.get("web"), dict):
        web = dict(cfg.get("web") or {})
        if body["web"].get("host"):
            web["host"] = str(body["web"]["host"]).strip()
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
            "xlsx": str(d / f"明细_{start}_{end}.xlsx"),
        } for a, d in sources if a is not None],
    }


APP_EXE = "快手数据工作台.exe"
UPDATE_DIR = kx.BASE / "updates"


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
                  "url": dl} if dl else None)
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


def _download(cfg: dict, url: str, dst: Path, logger):
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
    with resp, dst.open("wb") as f:
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
    logger(f"[update] 下载完成 {dst.stat().st_size / 1024 / 1024:.1f} MB")


def _write_updater(new_root: Path) -> Path:
    """生成延迟替换脚本：等本进程退出后换 exe 与 _internal，再启动新版本"""
    bat = UPDATE_DIR / "apply_update.bat"
    script = f"""@echo off
setlocal
set "APP={kx.BASE}"
set "NEW={new_root}"
set "PID={os.getpid()}"
:wait
tasklist /FI "PID eq %PID%" 2>nul | find "%PID%" >nul
if not errorlevel 1 (
  ping -n 2 127.0.0.1 >nul
  goto wait
)
if exist "%APP%\\_internal.old" rmdir /s /q "%APP%\\_internal.old"
if exist "%APP%\\_internal" move "%APP%\\_internal" "%APP%\\_internal.old"
if exist "%NEW%\\_internal" move "%NEW%\\_internal" "%APP%\\_internal"
if exist "%NEW%\\{APP_EXE}" move /y "%NEW%\\{APP_EXE}" "%APP%\\{APP_EXE}"
if not exist "%APP%\\_internal" if exist "%APP%\\_internal.old" move "%APP%\\_internal.old" "%APP%\\_internal"
start "" "%APP%\\{APP_EXE}"
del "%~f0"
"""
    try:
        bat.write_text(script, encoding="mbcs")
    except Exception:
        bat.write_text(script, encoding="utf-8")
    return bat


def stage_update(logger=print) -> dict:
    """下载 + 解压升级包并生成替换脚本（不负责退出进程）"""
    with CFG_LOCK:
        cfg = dict(CFG)
    if not (kx.BASE / "_internal").exists():
        raise RuntimeError("当前不是打包运行（缺少 _internal），无法自动升级")
    info = _update_info(cfg)
    if not info["has_update"]:
        raise RuntimeError(f"已是最新版本（v{info['current']}）")
    asset = info.get("asset")
    if not asset:
        raise RuntimeError("该 Release 未上传 zip 升级包")

    if UPDATE_DIR.exists():
        for p in UPDATE_DIR.iterdir():
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            else:
                p.unlink(missing_ok=True)
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)

    zip_path = UPDATE_DIR / asset["name"]
    _download(cfg, asset["url"], zip_path, logger)

    new_dir = UPDATE_DIR / "new"
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(new_dir)
    root = new_dir
    if not (root / APP_EXE).exists():
        subs = [p for p in new_dir.iterdir() if p.is_dir()]
        if len(subs) == 1 and (subs[0] / APP_EXE).exists():
            root = subs[0]
    if not (root / APP_EXE).exists() or not (root / "_internal").exists():
        raise RuntimeError("升级包里缺少 exe 或 _internal，已中止")

    bat = _write_updater(root)
    logger(f"[update] 升级包就绪（v{info['latest']}），即将重启替换")
    return {
        "current": info["current"],
        "latest": info["latest"],
        "install_dir": str(kx.BASE),
        "script": str(bat),
    }


def run_updater(bat: Path):
    flags = 0x00000008 | 0x00000200  # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(["cmd", "/c", str(bat)], creationflags=flags, close_fds=True)


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
        try:
            return json.loads(self.rfile.read(n).decode("utf-8") or "{}")
        except Exception:
            return {}

    # -- routes --------------------------------------------------
    def do_GET(self):
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
                allowed = str(p).startswith(str(data_root)) or p in excel_paths
                if not allowed or not p.exists():
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
                    started = time.time()
                    acc_rows, prod_rows, alerts, failures, owners_done = [], [], [], [], []
                    total_cost = total_gmv = total_orders = 0.0
                    for acct in owners:
                        label = acct.get("label") or acct.get("id")
                        _job_log(jid, f"[实时] 账号「{label}」开始…")
                        try:
                            ecfg = kx.account_cfg(cfg, acct)
                            res = kx.fetch_today(ecfg, None, logger=lambda m: _job_log(jid, m))
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
                            _job_log(jid, f"[实时] 账号「{label}」失败: {type(e).__name__}: {e}")
                            failures.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                             "account_id": "", "account_name": "", "error": str(e)[:300]})
                            owners_done.append({"id": str(acct.get("id")), "label": label, "ok": False,
                                                "error": str(e)[:300]})

                    if not any(o["ok"] for o in owners_done):
                        _job_done(jid, error="全部登录账号实时抓取失败")
                        return
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
                    _notify_alerts(alerts, logger=lambda m: _job_log(jid, m))
                    _job_done(jid, combined)

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
                    try:
                        res = stage_update(logger=lambda m: _job_log(jid, m))
                        run_updater(Path(res["script"]))
                        _job_log(jid, "[update] 替换脚本已启动，程序将在约 3 秒后退出并重启")
                        _job_done(jid, {k: v for k, v in res.items() if k != "script"})
                        app_runtime.quit_app(3.0)
                    except Exception as e:  # noqa: BLE001
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
                results = []
                for acct in owners:
                    label = acct.get("label") or acct.get("id")
                    ecfg = kx.account_cfg(cfg, acct)
                    try:
                        res = merge_excel.run_merge(
                            ecfg, start, end, mapping.get(str(acct.get("id"))),
                            export_dir=_owner_export_dir(start, end, str(acct.get("id"))),
                            backup=bool(cfg.get("auto_backup", True)))
                        print(f"[merge] 账号「{label}」完成: written={res.get('written')} "
                              f"saved={res.get('saved')} locked={res.get('locked')}")
                        results.append(dict(res, owner_id=str(acct.get("id")), owner_label=label))
                    except Exception as e:  # noqa: BLE001
                        print(f"[merge] 账号「{label}」失败: {type(e).__name__}: {e}")
                        results.append({"owner_id": str(acct.get("id")), "owner_label": label,
                                        "error": f"{type(e).__name__}: {e}", "saved": False,
                                        "locked": False, "written": 0, "checks": []})
                return self._json({
                    "results": results,
                    "saved": all(r.get("saved") for r in results) if results else False,
                    "locked": any(r.get("locked") for r in results),
                })

            return self._json({"error": "not found"}, 404)
        except Exception as e:  # noqa: BLE001
            print(f"[web] {self.command} {self.path} 失败: {type(e).__name__}: {e}")
            traceback.print_exc()
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)


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
