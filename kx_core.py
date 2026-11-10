# -*- coding: utf-8 -*-
"""快手磁力金牛取数核心：配置、登录、账户遍历、商品维度报表

数据链路（与后台页面一致）:
  投放账户列表 -> 点入账户首页（磁力金牛）
  -> 在账户页内用页面自带 axios（webpack 模块 50699）调用
     POST /rest/esp/promotion/report/getDetailReport
  签名头 esp-sature 由页面拦截器自动生成，无需自行签名

字段口径（已与后台页面逐项核对）:
  花费   = costTotal / 1000
  成交   = t0OrderPaymentAmt / 1000      # 总成交订单金额 = 直接 + 间接
  账户维度 = 该账户当日全部商品行求和
"""
from __future__ import annotations

import json
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Iterable

import app_secret

# Playwright 只在真正启动浏览器时导入：这样配置、导出、合并等纯逻辑
# 可以在未安装 playwright 的环境（如 WSL 单元测试）里直接运行。

# 打包成 exe 后：BASE 取 exe 所在目录（配置/数据/浏览器缓存在这里持久化），
# 只读资源（web 页面、配置模板）从 PyInstaller 解包目录取。
if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", BASE))
else:
    BASE = Path(__file__).resolve().parent
    BUNDLE_DIR = BASE

CONFIG_PATH = BASE / "config.json"
APP_VERSION = "1.3.4"
PROFILE_DIR = BASE / ".browser_profile"
PROFILES_DIR = BASE / ".browser_profiles"
DATA_DIR = BASE / "data"
LOGIN_URL = "https://login.kwaixiaodian.com/?biz=zone"
ACCOUNT_LIST_URL = "https://s.kwaixiaodian.com/zone/ad/account/bid"
NIU_HOME = "https://niu.e.kuaishou.com/home"
CST = timezone(timedelta(hours=8))

CHROME_ARGS = [
    "--no-sandbox",
    "--disable-gpu",
    "--disable-dev-shm-usage",
    "--disable-features=RendererCodeIntegrity",
    "--disable-blink-features=AutomationControlled",
]

# 页面原样使用的商品维度报表列（顺序与后台一致，缺列会导致退化为账户汇总）
PRODUCT_COLUMNS = [
    "merchantProductId", "costTotal", "t0OrderPaymentAmt", "t0Roi", "t0OrderCnt", "t0OrderCntCost",
    "showCnt", "itemEntranceClkCnt", "itemCardClkCnt", "qcpxWhiteboxDirectOrderCnt",
    "qcpxWhiteboxDirectOrderPaymentAmt", "netT0GMV", "netT0Roi", "netT0OrderCnt", "directGMV",
    "directOrderCntTotal", "directRoi", "directOrderCntCost", "netDirectGMV", "netDirectOrderCnt",
    "netDirectRoi", "netOrderPaymentAmtRatio", "directAttrSellerNewBuyerCnt", "t0ShopNewBuyerOrderCnt",
    "t0ShopNewBuyerOrderPaymentAmt", "qcpxDirectAttrSellerNewBuyerCnt", "qcpxT0ShopNewBuyerOrderCnt",
    "qcpxT0ShopNewBuyerOrderPaymentAmt", "qcpxMixApportionOrderPenetration",
    "qcpxMixApportionDirectGmvPenetration", "qcpxDirectOrderSellerBearApportionAmt",
    "qcpxDirectOrderPlatBearApportionAmt",
]

FETCH_PRODUCTS_JS = """
async ([payload, api]) => {
  if (!window.__wpr) {
    window.webpackChunkesp.push([[String(Math.random())], {}, function (r) { window.__wpr = r; }]);
  }
  if (!window.__wpr) return { ok: false, err: 'webpack require 注入失败' };
  let mod;
  try { mod = window.__wpr(50699); } catch (e) { return { ok: false, err: '模块 50699 不可用' }; }
  if (!mod || !mod.ZP) return { ok: false, err: 'axios 实例不可用' };
  try {
    const r = await mod.ZP.post(api, payload);
    return { ok: true, result: r && r.result, data: (r && r.data) || {} };
  } catch (e) {
    return { ok: false, err: (e && (e.msg || e.message)) || String(e) };
  }
}
"""

# 投放账户列表（快手小店后台）：登录后自动获取全量账户
ACCOUNT_LIST_JS = """
async (body) => {
  const r = await fetch('/rest/ad/kwaishop/account/myList', {
    method: 'POST', credentials: 'include',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body)
  });
  return {status: r.status, text: await r.text()};
}
"""

ACCOUNTS_CACHE = DATA_DIR / "accounts_cache.json"
REALTIME_CACHE = DATA_DIR / "realtime_cache.json"


def _account_page(data) -> tuple[list[dict], int | None]:
    """解析 myList 的分页返回：账户数组 + 总数（字段名各接口不统一）"""
    if isinstance(data, list):
        return data, None
    if not isinstance(data, dict):
        return [], None
    items: list[dict] = []
    for k in ("list", "records", "rows", "dataList", "pageData", "items"):
        v = data.get(k)
        if isinstance(v, list):
            items = v
            break
    total = None
    for k in ("total", "totalCount", "totalSize", "totalNum", "count"):
        v = data.get(k)
        if isinstance(v, bool):
            continue
        if isinstance(v, int):
            total = v
            break
        if isinstance(v, str) and v.isdigit():
            total = int(v)
            break
    return items, total


# ---------------------------------------------------------------- config
def load_config(path: Path | None = None) -> dict:
    p = path or CONFIG_PATH
    if not p.exists():
        tpl = BUNDLE_DIR / "config.example.json"
        if getattr(sys, "frozen", False) and tpl.exists():
            # 首次运行：用内置模板生成 config.json（模板里的说明文字不能当账号密码用）
            data = json.loads(tpl.read_text(encoding="utf-8"))
            data.pop("account", None)
            data.pop("password", None)
            p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            raise FileNotFoundError(f"配置文件不存在: {p}")
    cfg = json.loads(p.read_text(encoding="utf-8"))
    migrated = _normalize_accounts(cfg)
    cfg.setdefault("accounts", [])
    cfg.setdefault("headless", True)
    cfg.setdefault("date_offset_days", 1)
    cfg.setdefault("page_size", 100)
    cfg.setdefault("manual_login_timeout", 600)
    cfg.setdefault("auto_backup", True)
    cfg.setdefault("backup_limit", 10)
    cfg.setdefault("version", APP_VERSION)
    cfg.setdefault("update_repo", "")
    cfg.setdefault("update_source", "github")
    cfg.setdefault("update_webdav_path", "/version.json")
    cfg.setdefault("update_github_mirror_enabled", False)
    cfg.setdefault("update_github_mirror", "https://ghproxy.net/")
    cfg.setdefault("update_proxy", "")
    cfg.setdefault("realtime_auto_refresh", False)
    cfg.setdefault("realtime_interval_minutes", 5)
    cfg.setdefault("realtime_window_start", "08:50")
    cfg.setdefault("realtime_window_end", "18:05")
    cfg.setdefault("realtime_retry_enabled", True)
    cfg.setdefault("realtime_retry_max", 3)
    if migrated:
        try:
            p.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[config] 已自动升级配置结构：登录账号 {len(cfg['accounts'])} 个，密码已加密")
        except OSError as e:
            print(f"[config] 配置升级后写回失败: {e}")
    return cfg


def _ver_tuple(v) -> tuple:
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:6]) or (0,)


def _normalize_accounts(cfg: dict) -> bool:
    """兼容旧版单账号配置：统一为 accounts[]，并补齐字段、加密明文密码"""
    changed = False
    legacy = {
        "login_account": str(cfg.get("account") or "").strip(),
        "password": str(cfg.get("password") or ""),
        "login_role": str(cfg.get("login_role") or "auto"),
        "excel_path": str(cfg.get("excel_path") or "").strip(),
        "roi_book_path": str(cfg.get("roi_book_path") or "").strip(),
        "include_accounts": list(cfg.get("include_accounts") or []),
        "exclude_accounts": list(cfg.get("exclude_accounts") or []),
        "alert_thresholds": dict(cfg.get("alert_thresholds") or {}),
    }
    has_legacy = any((legacy["login_account"], legacy["password"], legacy["excel_path"]))

    raw = cfg.get("accounts")
    if not isinstance(raw, list):
        raw = []
        changed = True
    accts = [a for a in raw if isinstance(a, dict)]
    if len(accts) != len(raw):
        changed = True

    if not accts and has_legacy:
        accts = [{
            "id": "default",
            "label": "默认账号",
            "login_account": legacy["login_account"],
            "password": legacy["password"],
            "login_role": legacy["login_role"] or "auto",
            "excel_path": legacy["excel_path"],
            "roi_book_path": legacy["roi_book_path"],
            "enabled": True,
            "include_accounts": legacy["include_accounts"],
            "exclude_accounts": legacy["exclude_accounts"],
            "alert_thresholds": legacy["alert_thresholds"],
            "legacy_profile": True,
        }]
        changed = True

    seen_ids: set[str] = set()
    for idx, a in enumerate(accts):
        aid = str(a.get("id") or "")
        if not aid or aid in seen_ids:
            a["id"] = "default" if idx == 0 and "default" not in seen_ids else f"acc{idx + 1}"
            changed = True
        seen_ids.add(str(a["id"]))
        if not a.get("label"):
            a["label"] = str(a.get("login_account") or f"账号 {idx + 1}")
            changed = True
        if idx == 0:
            for k, v in legacy.items():
                if v and not a.get(k):
                    a[k] = v
                    changed = True
        for k, v in (("login_role", "auto"), ("enabled", True),
                     ("include_accounts", []), ("exclude_accounts", []),
                     ("alert_thresholds", {}), ("legacy_profile", False),
                     ("roi_book_path", ""), ("alert_cooldown_minutes", 30)):
            if k not in a:
                a[k] = v
                changed = True
        if "week_cost_source" in a:
            a.pop("week_cost_source", None)
            changed = True
        password = str(a.get("password") or "")
        if password and not app_secret.is_protected(password):
            try:
                a["password"] = app_secret.protect_text(password)
                changed = True
            except Exception as e:  # noqa: BLE001
                # 绝不把明文密码留在 config.json：加密失败时清空，要求重新录入
                a["password"] = ""
                changed = True
                print(f"[config] 账号密码加密失败，已清空该账号密码，请重新保存: {e}")

    if accts and PROFILE_DIR.exists() and not any(a.get("legacy_profile") for a in accts):
        accts[0]["legacy_profile"] = True
        changed = True

    if not accts:
        cfg["accounts"] = []
    else:
        cfg["accounts"] = accts

    # 旧版顶层字段迁移完成后移除，避免残留明文密码
    for k in ("account", "password", "excel_path", "roi_book_path", "login_role",
              "include_accounts", "exclude_accounts", "alert_thresholds"):
        if k in cfg:
            cfg.pop(k, None)
            changed = True

    # 程序升级后同步配置里的版本号，避免升级完成仍提示有新版本
    cur = str(cfg.get("version") or "").strip()
    if not cur or _ver_tuple(cur) < _ver_tuple(APP_VERSION):
        cfg["version"] = APP_VERSION
        changed = True
    return changed


def default_account_cfg(cfg: dict) -> dict:
    """CLI 用：取第一个账号的有效配置"""
    accounts = login_accounts(cfg)
    return account_cfg(cfg, accounts[0]) if accounts else cfg


def login_accounts(cfg: dict) -> list[dict]:
    return [a for a in (cfg.get("accounts") or [])
            if isinstance(a, dict) and a.get("id")]


def find_login_account(cfg: dict, account_id: str) -> dict | None:
    for a in login_accounts(cfg):
        if str(a.get("id")) == str(account_id):
            return a
    return None


def account_profile_dir(acct: dict) -> Path:
    """每个登录账号独立浏览器目录；迁移来的第一个账号沿用旧目录保留登录态"""
    if acct.get("legacy_profile"):
        return PROFILE_DIR
    return PROFILES_DIR / str(acct.get("id") or "default")


def account_cfg(cfg: dict, acct: dict) -> dict:
    """把全局配置与某个登录账号合并，供 Client / 导出 / 合并使用"""
    effective = dict(cfg)
    effective.update({
        "account": str(acct.get("login_account") or "").strip(),
        "password": app_secret.unprotect_text(acct.get("password")),
        "login_role": str(acct.get("login_role") or "auto"),
        "excel_path": str(acct.get("excel_path") or ""),
        "roi_book_path": str(acct.get("roi_book_path") or ""),
        "include_accounts": list(acct.get("include_accounts") or []),
        "exclude_accounts": list(acct.get("exclude_accounts") or []),
        "alert_thresholds": dict(acct.get("alert_thresholds") or {}),
        "_account_id": str(acct.get("id")),
        "_account_label": str(acct.get("label") or acct.get("login_account") or acct.get("id")),
        "_profile_dir": str(account_profile_dir(acct)),
        "_cache_owner": None if acct.get("legacy_profile") else str(acct.get("id")),
    })
    return effective


def resolve_path(v: str) -> Path:
    p = Path(v)
    return p if p.is_absolute() else (BASE / p).resolve()


def default_date(cfg: dict) -> str:
    off = int(cfg.get("date_offset_days", 1))
    return (datetime.now(CST) - timedelta(days=off)).strftime("%Y-%m-%d")


def day_range(start: str, end: str) -> list[datetime]:
    d0 = datetime.strptime(start, "%Y-%m-%d").replace(tzinfo=CST)
    d1 = datetime.strptime(end, "%Y-%m-%d").replace(tzinfo=CST)
    out, cur = [], d0
    while cur <= d1:
        out.append(cur)
        cur += timedelta(days=1)
    return out


# ---------------------------------------------------------------- client
@dataclass
class ProductRow:
    date: str
    account_id: str
    account_name: str
    product_id: str
    product_name: str
    cost: float
    gmv: float
    order_cnt: int
    roi: float
    show_cnt: int
    click_cnt: int
    raw: dict | None = None


class Client:
    """磁力金牛后台客户端（Playwright 驱动，复用登录态）"""

    def __init__(self, cfg: dict, logger=print):
        self.cfg = cfg
        self.log = logger
        self._pw = None
        self.ctx = None
        self.list_page = None

    # -- lifecycle --------------------------------------------------
    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *exc):
        self.close()

    def start(self, headless: bool | None = None):
        if headless is None:
            headless = bool(self.cfg.get("headless", True))
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.ctx = self._launch(headless)
        self.list_page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.list_page.set_default_timeout(45_000)
        self.ensure_login()

    def _launch(self, headless: bool):
        profile = Path(self.cfg.get("_profile_dir") or PROFILE_DIR)
        profile.mkdir(parents=True, exist_ok=True)
        return self._pw.chromium.launch_persistent_context(
            user_data_dir=str(profile),
            channel="chrome",
            headless=headless,
            viewport={"width": 1600, "height": 1000} if headless else None,
            no_viewport=not headless,
            args=(CHROME_ARGS + ([] if headless else ["--start-maximized"])),
        )

    def _relaunch_headed(self):
        self.log("[login] 切换为有头模式，等待人工完成登录")
        try:
            self.ctx.close()
        except Exception:
            pass
        self.ctx = self._launch(False)
        self.list_page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.list_page.set_default_timeout(45_000)

    def close(self):
        """关闭浏览器；ctx.close 偶发卡住，加超时避免整个流程挂死"""
        import threading
        if self.ctx is not None:
            def _close_ctx():
                try:
                    self.ctx.close()
                except Exception:
                    pass
            t = threading.Thread(target=_close_ctx, daemon=True)
            t.start()
            t.join(20)
        try:
            if self._pw is not None:
                self._pw.stop()
        except Exception:
            pass

    # -- login ------------------------------------------------------
    def ensure_login(self):
        page = self.list_page
        page.goto(ACCOUNT_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(3000)
        if "login.kwaixiaodian.com" not in page.url:
            self.log("[login] 登录态有效")
            return

        account = str(self.cfg.get("account") or "").strip()
        password = str(self.cfg.get("password") or "")
        if account and password:
            self.log(f"[login] 密码登录（账号 {account[:3]}****）")
            if self._password_login(page, account, password):
                # 登录成功后会先落在 www 首页，需再进一次后台确认登录态真正生效
                page.goto(ACCOUNT_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
                page.wait_for_timeout(4000)
                if "login.kwaixiaodian.com" not in page.url:
                    self.log("[login] 后台校验通过")
                    return
                self.log("[login] 密码登录成功但后台仍要求登录（账号可能未开通/未绑定店铺）")
            self.log("[login] 密码登录未成功（可能触发验证码/滑块）")

        self._manual_login(page)

    # 登录页身份 tab 与对应的提交按钮文案（两套界面文案不同，写死任一都会导致点不到按钮）
    LOGIN_ROLES = {
        "owner": ("我是店主", "快手APP账号授权登录"),
        "staff": ("我是员工", "快手商家账号登录"),
    }

    def _password_login(self, page, account: str, password: str) -> bool:
        role = str(self.cfg.get("login_role") or "auto").lower()
        if role not in ("auto", "owner", "staff"):
            role = "auto"
        # auto 先按子账号试：投放运营多用员工账号；失败且提示账号不存在时再按店主试
        order = ["staff", "owner"] if role == "auto" else [role]
        for i, r in enumerate(order):
            if self._submit_password(page, account, password, r):
                return True
            if i < len(order) - 1:
                self.log("[login] 该身份登录失败，切换另一种身份重试")
        return False

    def _submit_password(self, page, account: str, password: str, role: str) -> bool:
        tab, button = self.LOGIN_ROLES[role]
        label = "员工/子账号" if role == "staff" else "店主"
        try:
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(2500)
            page.get_by_text(tab, exact=True).first.click(timeout=8000)
            page.wait_for_timeout(600)
            page.get_by_text("密码登录", exact=True).first.click(timeout=8000)
            page.wait_for_timeout(800)
            page.locator('input[placeholder="请输入手机号"]').first.fill(account)
            page.locator('input[placeholder="请输入密码"]').first.fill(password)
            page.wait_for_timeout(200)
            page.get_by_role("button", name=button).first.click(timeout=8000)
            for _ in range(15):
                page.wait_for_timeout(1000)
                if "login.kwaixiaodian.com" not in page.url:
                    self.log(f"[login] 密码登录成功（{label}）")
                    return True
                err = self._login_error(page)
                if err:
                    self.log(f"[login] {label}登录失败：{err}")
                    return False
            self.log(f"[login] {label}登录超时未跳转")
            return False
        except Exception as e:  # noqa: BLE001
            self.log(f"[login] 密码登录异常: {type(e).__name__}: {e}")
            return False

    @staticmethod
    def _login_error(page) -> str:
        """读取登录页上的错误提示（没有则返回空串）"""
        try:
            return page.evaluate("""() => {
                const hit = Array.from(document.querySelectorAll('*'))
                    .filter(e => e.children.length === 0)
                    .map(e => (e.innerText || '').trim())
                    .find(t => /账号不存在|密码错误|密码不正确|未注册|操作频繁|请完成验证|验证码/.test(t));
                return hit || '';
            }""") or ""
        except Exception:  # noqa: BLE001
            return ""

    def _manual_login(self, page):
        self._relaunch_headed()
        page = self.list_page
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60_000)
        timeout = int(self.cfg.get("manual_login_timeout", 600))
        self.log(f"[login] 请在弹出的浏览器中扫码/密码登录（最多等待 {timeout} 秒）")
        deadline = time.time() + timeout
        while time.time() < deadline:
            page.wait_for_timeout(3000)
            url = page.url
            if "login.kwaixiaodian.com" not in url and "s.kwaixiaodian.com" in url:
                self.log("[login] 人工登录完成")
                return
            if "login.kwaixiaodian.com" not in url:
                page.goto(ACCOUNT_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
                page.wait_for_timeout(2000)
                if "login.kwaixiaodian.com" not in page.url:
                    self.log("[login] 人工登录完成")
                    return
        raise RuntimeError("登录超时，请检查账号密码或网络后重试")

    # -- accounts ---------------------------------------------------
    def list_accounts(self) -> list[dict]:
        """登录后自动获取全量投放账户（myList 接口）"""
        page = self.list_page
        if "s.kwaixiaodian.com" not in page.url:
            page.goto(ACCOUNT_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(3000)
        out, seen = [], set()
        page_size = 100
        total = None
        pages = 0
        for page_no in range(1, 201):
            payload = {"pageSize": page_size, "current": page_no,
                       "queryActiveFlag": False, "pageNo": page_no}
            res = page.evaluate(ACCOUNT_LIST_JS, payload)
            if not res or res.get("status") != 200:
                raise RuntimeError(f"账户列表接口异常: HTTP {(res or {}).get('status')}")
            j = json.loads(res.get("text") or "{}")
            items, page_total = _account_page(j.get("data"))
            if page_total is not None and total is None:
                total = page_total
            if not items:
                break
            pages = page_no
            new = 0
            for it in items:
                aid = it.get("accountId")
                if aid is None or str(aid) in seen:
                    continue
                seen.add(str(aid))
                new += 1
                out.append({
                    "id": str(aid),
                    "name": it.get("accountName") or "",
                    "type": it.get("accountTypeName") or "",
                    "status": it.get("reviewStatusName") or "",
                    "corp": it.get("corpName") or "",
                })
            if total is not None and len(out) >= total:
                break
            # 服务端会把 pageSize 压到默认值（如 15/20），不能再用 len(items) < page_size 判断结束；
            # 只有整页没出现新账户（分页参数未生效）才停，避免死循环
            if new == 0:
                break
        if total is not None and len(out) < total:
            print(f"[accounts] 账户列表分页未取全: {len(out)}/{total}，接口可能未按页返回")
        else:
            print(f"[accounts] 账户列表共 {len(out)} 个（{pages} 页）")
        save_accounts_cache(out, self.cfg.get("_cache_owner"))
        return out

    def _locate_account_row(self, page, account_id: str):
        """在账户列表定位某账户行：当前页 → 搜索框 → 逐页翻页"""
        row = page.locator("tr", has_text=str(account_id)).first
        if row.count() > 0:
            return row

        box = page.locator('input[placeholder*="账户"]').first
        if box.count() == 0:
            box = page.locator('input[placeholder*="搜索"]').first
        if box.count() > 0:
            try:
                box.fill(str(account_id))
                box.press("Enter")
                page.wait_for_timeout(2500)
                row = page.locator("tr", has_text=str(account_id)).first
                if row.count() > 0:
                    self.log(f"[accounts] 通过搜索定位账户 {account_id}")
                    return row
            except Exception:
                pass

        prev_first = ""
        for _ in range(10):
            try:
                nxt = page.locator("li[title='下一页'], button[aria-label='下一页'], "
                                   ".ant-pagination-next, .semi-page-next").first
                if nxt.count() == 0:
                    break
                first_txt = page.locator("tbody tr").first.inner_text()[:80]
            except Exception:
                break
            if first_txt == prev_first:
                break
            prev_first = first_txt
            try:
                nxt.click(timeout=5000)
                page.wait_for_timeout(1800)
            except Exception:
                break
            row = page.locator("tr", has_text=str(account_id)).first
            if row.count() > 0:
                self.log(f"[accounts] 翻页定位账户 {account_id}")
                return row
        return None

    def open_account(self, account_id: str):
        """从列表页点入账户首页，返回新页；调用方负责 close"""
        page = self.list_page
        if "s.kwaixiaodian.com" not in page.url:
            page.goto(ACCOUNT_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(3000)
        page.bring_to_front()
        row = self._locate_account_row(page, str(account_id))
        if row is None:
            raise RuntimeError(f"账户列表未找到 {account_id}")
        with self.ctx.expect_page(timeout=30_000) as info:
            row.locator("p.ad-shop-account-text-with-id-text-navigate").first.click()
        home = info.value
        home.wait_for_load_state("domcontentloaded", timeout=60_000)
        home.wait_for_timeout(4000)
        if not self._wait_ready(home, account_id):
            raise RuntimeError(f"账户页未就绪: {home.url}")
        return home

    @staticmethod
    def _wait_ready(page, account_id: str, timeout_s: int = 45) -> bool:
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            try:
                if str(page.evaluate("() => window.esp_AccountId || ''")) == str(account_id):
                    return True
            except Exception:
                pass
            page.wait_for_timeout(1000)
        return False

    # -- report -----------------------------------------------------
    def fetch_products_range(self, page, account_id: str, account_name: str,
                             start: datetime, end: datetime,
                             retries: int = 3) -> list[ProductRow]:
        """一次请求取整个区间（groupType=1 返回按天明细，每行带 reportDate）"""
        begin = int(start.timestamp() * 1000)
        end_ms = int(end.timestamp() * 1000) + 86_399_999
        page_size = int(self.cfg.get("page_size", 100))
        rows: list[ProductRow] = []
        raw_seen = 0
        total = 0
        page_no = 1
        while page_no <= 1000:
            payload = {
                "bizType": 2,
                "reportSceneType": 19321,
                "selectColumn": PRODUCT_COLUMNS,
                "whereCondition": [{"name": "sceneOrientedType", "values": ["24"], "compareOp": 0}],
                "sortCondition": [],
                "timeCondition": {"beginTime": begin, "endTime": end_ms, "groupType": 1},
                "havingCondition": [],
                "pageInfo": {"currentPage": page_no, "pageSize": page_size, "totalCount": 0},
                "bizSubType": 22,
                "viewType": 1,
            }
            res = None
            for attempt in range(retries):
                try:
                    res = page.evaluate(FETCH_PRODUCTS_JS, [payload, "/promotion/report/getDetailReport"])
                except Exception as e:  # noqa: BLE001
                    res = {"ok": False, "err": f"{type(e).__name__}: {e}"}
                if res and res.get("ok"):
                    break
                page.wait_for_timeout(800 * (attempt + 1))
            if not res or not res.get("ok"):
                raise RuntimeError(f"{account_id} 区间取数失败: {(res or {}).get('err')}")

            data = res.get("data") or {}
            chunk = data.get("detail") or []
            for r in chunk:
                pid = r.get("merchantProductId")
                if not pid:
                    continue
                d = str(r.get("reportDate") or r.get("reportDateStr") or "")[:10]
                if not d:
                    continue
                rows.append(ProductRow(
                    date=d,
                    account_id=str(account_id),
                    account_name=account_name,
                    product_id=str(pid),
                    product_name=(r.get("itemName") or "").strip(),
                    cost=round((r.get("costTotal") or 0) / 1000, 2),
                    gmv=round((r.get("t0OrderPaymentAmt") or 0) / 1000, 2),
                    order_cnt=int(r.get("t0OrderCnt") or 0),
                    roi=round(float(r.get("t0Roi") or 0), 4),
                    show_cnt=int(r.get("showCnt") or 0),
                    click_cnt=int(r.get("itemEntranceClkCnt") or 0),
                    raw=dict(r),
                ))
            raw_seen += len(chunk)
            total = int(data.get("totalCount") or 0)
            if not chunk:
                break
            if total > 0 and raw_seen >= total:
                break
            page_no += 1
        if total > 0 and raw_seen < total:
            self.log(f"[report] 警告: {account_id} 区间仅取到 {raw_seen}/{total} 行")
        return rows

    def fetch_products(self, page, account_id: str, account_name: str, day: datetime,
                       retries: int = 3) -> list[ProductRow]:
        begin = int(day.timestamp() * 1000)
        date_str = day.strftime("%Y-%m-%d")
        page_size = int(self.cfg.get("page_size", 100))
        rows: list[ProductRow] = []
        raw_seen = 0
        total = 0
        page_no = 1
        while page_no <= 1000:
            payload = {
                "bizType": 2,
                "reportSceneType": 19321,
                "selectColumn": PRODUCT_COLUMNS,
                "whereCondition": [{"name": "sceneOrientedType", "values": ["24"], "compareOp": 0}],
                "sortCondition": [],
                "timeCondition": {"beginTime": begin, "endTime": begin + 86_399_999, "groupType": 3},
                "havingCondition": [],
                "pageInfo": {"currentPage": page_no, "pageSize": page_size, "totalCount": 0},
                "bizSubType": 22,
                "viewType": 1,
            }
            res = None
            for attempt in range(retries):
                try:
                    res = page.evaluate(FETCH_PRODUCTS_JS, [payload, "/promotion/report/getDetailReport"])
                except Exception as e:  # noqa: BLE001
                    res = {"ok": False, "err": f"{type(e).__name__}: {e}"}
                if res and res.get("ok"):
                    break
                page.wait_for_timeout(800 * (attempt + 1))
            if not res or not res.get("ok"):
                raise RuntimeError(f"{account_id} {date_str} 取数失败: {(res or {}).get('err')}")

            data = res.get("data") or {}
            chunk = data.get("detail") or []
            for r in chunk:
                pid = r.get("merchantProductId")
                if not pid:
                    continue
                rows.append(ProductRow(
                    date=date_str,
                    account_id=str(account_id),
                    account_name=account_name,
                    product_id=str(pid),
                    product_name=(r.get("itemName") or "").strip(),
                    cost=round((r.get("costTotal") or 0) / 1000, 2),
                    gmv=round((r.get("t0OrderPaymentAmt") or 0) / 1000, 2),
                    order_cnt=int(r.get("t0OrderCnt") or 0),
                    roi=round(float(r.get("t0Roi") or 0), 4),
                    show_cnt=int(r.get("showCnt") or 0),
                    click_cnt=int(r.get("itemEntranceClkCnt") or 0),
                    raw=dict(r),
                ))
            raw_seen += len(chunk)
            total = int(data.get("totalCount") or 0)
            if not chunk:
                break
            # totalCount 缺失（0）时不能当成“已取完”，否则只取第一页；
            # 以空页作为结束条件，totalCount 仅用于提前结束。
            if total > 0 and raw_seen >= total:
                break
            page_no += 1
        else:
            raise RuntimeError(f"{account_id} {date_str} 分页超过 1000 页，疑似接口分页异常")
        if total > 0 and raw_seen < total:
            self.log(f"[report] 警告: {account_id} {date_str} 仅取到 {raw_seen}/{total} 行，数据可能不完整")
        return rows


def accounts_cache_path(owner: str | None = None) -> Path:
    return ACCOUNTS_CACHE if not owner else (DATA_DIR / f"accounts_cache_{owner}.json")


def save_accounts_cache(accounts: list[dict], owner: str | None = None):
    try:
        path = accounts_cache_path(owner)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({
            "updated_at": datetime.now(CST).strftime("%Y-%m-%d %H:%M:%S"),
            "accounts": accounts,
        }, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass


def load_accounts_cache(owner: str | None = None) -> tuple[list[dict], str]:
    path = accounts_cache_path(owner)
    if not path.exists():
        return [], ""
    try:
        j = json.loads(path.read_text(encoding="utf-8"))
        return j.get("accounts") or [], j.get("updated_at") or ""
    except Exception:
        return [], ""


def excel_account_ids(path: Path) -> list[str]:
    """读月度表第 3 行的账户ID（按动态商品组列扫描，兼容后加列）"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            ws = wb["消耗成交"] if "消耗成交" in wb.sheetnames else wb[wb.sheetnames[0]]
            max_col = int(ws.max_column or 0)
            total_start = None
            for c in range(2, max_col + 1):
                if str(ws.cell(row=1, column=c).value or "").strip() == "总计":
                    total_start = c
                    break
            if total_start is None:
                total_start = 29
            ids = []
            for col in range(2, total_start, 3):
                v = ws.cell(row=3, column=col).value
                if v is not None and str(v).strip():
                    ids.append(str(v).strip())
            return ids
        finally:
            wb.close()
    except Exception:
        return []


def pick_accounts(all_accounts: list[dict], cfg: dict, ids: list[str] | None = None) -> list[dict]:
    """账户筛选优先级：显式 ids > include/exclude 配置 > 与月度表取交集 > 全部"""
    exclude = {str(x) for x in (cfg.get("exclude_accounts") or [])}
    pool = [a for a in all_accounts if a["id"] not in exclude]

    if ids:
        wanted = {str(x) for x in ids}
        picked = [a for a in pool if a["id"] in wanted]
        got = {a["id"] for a in picked}
        # 接口列表里没返回的账户也保留：勾选过的账户都要尝试拉取并在预览里出现
        for x in ids:
            sid = str(x)
            if sid not in got:
                picked.append({"id": sid, "name": ""})
                got.add(sid)
        return picked

    include = cfg.get("include_accounts") or []
    if include:
        wanted = {str(x) for x in include}
        return [a for a in pool if a["id"] in wanted]

    if cfg.get("filter_by_excel", True) and cfg.get("excel_path"):
        try:
            excel = resolve_path(cfg["excel_path"])
        except Exception:
            excel = None
        if excel and excel.exists():
            table_ids = set(excel_account_ids(excel))
            if table_ids:
                inter = [a for a in pool if a["id"] in table_ids]
                if inter:
                    return inter
    return pool


def fetch_accounts(cfg: dict, logger=print) -> list[dict]:
    """独立获取账户列表（供网页刷新用）"""
    with Client(cfg, logger) as client:
        accounts = client.list_accounts()
        logger(f"[accounts] 获取到 {len(accounts)} 个账户")
        return accounts


def today_str() -> str:
    return datetime.now(CST).strftime("%Y-%m-%d")


def fetch_today(cfg: dict, account_ids: list[str] | None = None, logger=print) -> dict:
    """当日实时数据：全账号 × 全商品，聚合到账户维度（只写缓存，不落导出目录）"""
    started = time.time()
    day = datetime.strptime(today_str(), "%Y-%m-%d").replace(tzinfo=CST)
    prod: list[ProductRow] = []
    failures: list[dict] = []

    with Client(cfg, logger) as client:
        all_accounts = client.list_accounts()
        accounts = pick_accounts(all_accounts, cfg, account_ids)
        if not accounts:
            raise RuntimeError("没有可查询的账户（检查登录态或 include/exclude 配置）")
        try:
            retry_max = max(0, min(int(cfg.get("realtime_retry_max") or 3), 10))
        except (TypeError, ValueError):
            retry_max = 3
        if not cfg.get("realtime_retry_enabled"):
            retry_max = 0
        logger(f"[实时] 账户 {len(accounts)} 个，查询日期 {day:%Y-%m-%d}"
               + (f"，失败最多重试 {retry_max} 次" if retry_max else "，不重试"))
        for acc in accounts:
            aid, aname = str(acc["id"]), acc.get("name") or ""
            attempt = 0
            while True:
                home = None
                try:
                    home = client.open_account(aid)
                    rows = client.fetch_products(home, aid, aname, day)
                    prod.extend(rows)
                    logger(f"    {aname or aid} 商品 {len(rows)} 个 花费 {round(sum(r.cost for r in rows), 2)} "
                           f"成交 {round(sum(r.gmv for r in rows), 2)}")
                    break
                except Exception as e:  # noqa: BLE001
                    if attempt >= retry_max:
                        logger(f"    {aname or aid} 失败: {type(e).__name__}: {e}"
                               + (f"（已重试 {attempt} 次）" if attempt else ""))
                        failures.append({"account_id": aid, "account_name": aname,
                                         "error": str(e)[:300], "attempts": attempt + 1})
                        break
                    attempt += 1
                    logger(f"    {aname or aid} 第 {attempt}/{retry_max} 次重试: {type(e).__name__}: {e}")
                    time.sleep(min(2 * attempt, 6))
                finally:
                    if home is not None:
                        try:
                            home.close()
                        except Exception:
                            pass

    agg = aggregate(prod)
    total_cost = round(sum(a["cost"] for a in agg), 2)
    total_gmv = round(sum(a["gmv"] for a in agg), 2)

    # 账户行挂主商品（消耗最大的那个），便于页面展示「商品 + 账户」
    primary: dict[str, tuple] = {}
    for r in prod:
        cur = primary.get(r.account_id)
        if cur is None or (r.cost, r.gmv) > cur[0]:
            primary[r.account_id] = ((r.cost, r.gmv), r.product_id, r.product_name)
    accounts_out = []
    for a in agg:
        row = dict(a, roi=round(a["gmv"] / a["cost"], 2) if a["cost"] else 0.0)
        item = primary.get(a["account_id"])
        if item:
            row["product_id"], row["product_name"] = item[1], item[2]
        accounts_out.append(row)

    data = {
        "date": day.strftime("%Y-%m-%d"),
        "fetched_at": datetime.now(CST).strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed": round(time.time() - started, 1),
        "accounts": accounts_out,
        "products": [{"date": r.date, "account_id": r.account_id, "account_name": r.account_name,
                      "product_id": r.product_id, "product_name": r.product_name,
                      "cost": r.cost, "gmv": r.gmv, "order_cnt": r.order_cnt, "roi": r.roi} for r in prod],
        "totals": {
            "cost": total_cost,
            "gmv": total_gmv,
            "roi": round(total_gmv / total_cost, 2) if total_cost else 0,
            "order_cnt": sum(a["order_cnt"] for a in agg),
            "account_rows": len(agg),
            "product_rows": len(prod),
        },
        "failures": failures,
    }
    save_realtime_cache(data)
    logger(f"[实时] 完成：账户 {len(agg)} 消耗 {total_cost} 成交 {total_gmv} 用时 {data['elapsed']}s")
    return data


def save_realtime_cache(data: dict):
    try:
        REALTIME_CACHE.parent.mkdir(parents=True, exist_ok=True)
        REALTIME_CACHE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def load_realtime_cache() -> dict | None:
    if not REALTIME_CACHE.exists():
        return None
    try:
        return json.loads(REALTIME_CACHE.read_text(encoding="utf-8"))
    except Exception:
        return None


def aggregate(rows: Iterable[ProductRow]) -> list[dict]:
    """商品行 -> 账户×日期 汇总"""
    agg: dict[tuple, dict] = {}
    for r in rows:
        key = (r.date, r.account_id)
        a = agg.setdefault(key, {
            "date": r.date, "account_id": r.account_id, "account_name": r.account_name,
            "cost": 0.0, "gmv": 0.0, "order_cnt": 0, "product_count": 0,
        })
        a["cost"] = round(a["cost"] + r.cost, 2)
        a["gmv"] = round(a["gmv"] + r.gmv, 2)
        a["order_cnt"] += r.order_cnt
        a["product_count"] += 1
    return sorted(agg.values(), key=lambda x: (x["date"], x["account_id"]))