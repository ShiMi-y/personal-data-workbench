# -*- coding: utf-8 -*-
"""桌面运行时外壳：文件日志、托盘图标、启动失败提示

打包成无控制台窗口（--windowed）后 sys.stdout / sys.stderr 为 None，
此时任何一个 print() 都会抛 AttributeError，所以本模块必须在导入其他本地模块之前
调用 install()，把输出接到 logs/ 下的日志文件。
"""
from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path

KEEP_DAYS = 14


def app_dir() -> Path:
    """exe 所在目录；源码运行时为项目目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bundle_dir() -> Path:
    """只读资源目录；打包后为 PyInstaller 解包目录"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", app_dir()))
    return Path(__file__).resolve().parent


LOG_DIR = app_dir() / "logs"
_TRAY_THREAD_ID = 0
WM_QUIT = 0x0012
_NOTIFIER = None


class _LogStream:
    """按行加时间戳写文件，可选同时回显到原控制台"""

    def __init__(self, path: Path | None, echo=None):
        self._echo = echo
        self._fh = open(path, "a", encoding="utf-8", buffering=1) if path else None
        self._buf = ""
        self._lock = threading.Lock()

    @property
    def encoding(self) -> str:
        return "utf-8"

    def isatty(self) -> bool:
        return False

    def write(self, s):
        if not s:
            return 0
        if self._echo is not None:
            try:
                self._echo.write(s)
                self._echo.flush()
            except Exception:
                pass
        with self._lock:
            self._buf += s
            while "\n" in self._buf:
                line, self._buf = self._buf.split("\n", 1)
                if self._fh is not None:
                    self._fh.write(f"[{time.strftime('%H:%M:%S')}] {line}\n")
            if self._fh is not None:
                self._fh.flush()
        return len(s)

    def flush(self):
        with self._lock:
            if self._fh is not None:
                if self._buf:
                    self._fh.write(f"[{time.strftime('%H:%M:%S')}] {self._buf}")
                    self._buf = ""
                self._fh.flush()


def _cleanup_old_logs():
    try:
        cutoff = time.time() - KEEP_DAYS * 86400
        for p in LOG_DIR.glob("app-*.log"):
            if p.stat().st_mtime < cutoff:
                p.unlink()
    except Exception:
        pass


def _log_uncaught(exc_type, exc, tb):
    try:
        import traceback
        detail = "".join(traceback.format_exception(exc_type, exc, tb))
    except Exception:
        detail = f"{exc_type}: {exc}"
    try:
        print(f"[app] 未捕获异常:\n{detail}")
    except Exception:
        pass


def install_exception_hooks():
    """把主线程 / 子线程的未捕获异常也写进日志文件"""
    sys.excepthook = _log_uncaught

    def _thread_hook(args):
        _log_uncaught(args.exc_type, args.exc_value, args.exc_traceback)

    try:
        threading.excepthook = _thread_hook
    except Exception:
        pass


def install() -> Path | None:
    """接管 stdout/stderr（必须在其他本地模块导入前调用）"""
    global LOG_DIR
    stream, path = None, None
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        _cleanup_old_logs()
        path = LOG_DIR / f"app-{datetime.now():%Y%m%d}.log"
        stream = _LogStream(path, echo=sys.stdout)
    except Exception:
        stream = _LogStream(None, echo=sys.stdout)
    sys.stdout = sys.stderr = stream
    install_exception_hooks()
    if path is not None:
        print(f"[app] 日志文件: {path}")
    return path


def show_error(title: str, text: str):
    """无控制台时的兜底提示（阻塞到用户点确定）"""
    print(f"[app] {title}: {text}")
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, title, 0x10)
    except Exception:
        pass


def notify(title: str, text: str):
    """Windows 托盘气泡通知（无托盘时只写日志）"""
    print(f"[notify] {title}: {text}")
    fn = _NOTIFIER
    if fn is None:
        return
    try:
        fn(title, text)
    except Exception as e:  # noqa: BLE001
        print(f"[notify] 弹窗失败: {type(e).__name__}: {e}")


def _make_notifier(user32, shell32, nid_cls, hwnd):
    import ctypes

    def _fn(title: str, text: str):
        n = nid_cls()
        n.cbSize = ctypes.sizeof(nid_cls)
        n.hWnd = hwnd
        n.uID = 1
        n.uFlags = 0x00000010  # NIF_INFO
        n.szInfoTitle = str(title)[:63]
        n.szInfo = str(text)[:255]
        n.dwInfoFlags = 0x00000002  # NIIF_WARNING
        shell32.Shell_NotifyIconW(1, ctypes.byref(n))  # NIM_MODIFY

    return _fn


def quit_app(delay: float = 1.0):
    """延时退出：托盘模式发 WM_QUIT 结束消息循环，无托盘时直接退出进程"""
    def _later():
        time.sleep(delay)
        if _TRAY_THREAD_ID:
            try:
                import ctypes
                ctypes.windll.user32.PostThreadMessageW(_TRAY_THREAD_ID, WM_QUIT, 0, 0)
                return
            except Exception:
                pass
        os._exit(0)

    threading.Thread(target=_later, daemon=True).start()


def cleanup_old_version():
    """升级后遗留的 _internal.old（新版本已正常启动才删）"""
    def _work():
        time.sleep(5)
        p = app_dir() / "_internal.old"
        if p.exists():
            try:
                import shutil
                shutil.rmtree(p, ignore_errors=True)
                print("[app] 已清理旧版本目录 _internal.old")
            except Exception:
                pass

    threading.Thread(target=_work, daemon=True).start()


def _icon_path() -> Path:
    return bundle_dir() / "assets" / "app.ico"


def run_tray(url: str, on_quit) -> bool:
    """Win32 托盘图标（ctypes 实现，不依赖第三方库）

    双击图标打开工作台，右键菜单可打开工作台 / 日志目录 / 退出。
    返回 False 表示当前环境不可用，调用方应退回前台常驻。
    """
    try:
        import ctypes
        from ctypes import wintypes
    except Exception as e:  # noqa: BLE001
        print(f"[tray] ctypes 不可用: {e}")
        return False
    if sys.platform != "win32":
        print("[tray] 非 Windows 环境，跳过托盘")
        return False

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    shell32 = ctypes.WinDLL("shell32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    WM_DESTROY = 0x0002
    WM_TRAY = 0x0400 + 20
    WM_LBUTTONUP = 0x0202
    WM_LBUTTONDBLCLK = 0x0203
    WM_RBUTTONUP = 0x0205
    NIM_ADD, NIM_DELETE = 0, 2
    NIF_MESSAGE, NIF_ICON, NIF_TIP = 0x1, 0x2, 0x4
    IMAGE_ICON, LR_LOADFROMFILE, LR_DEFAULTSIZE = 1, 0x10, 0x40
    MF_STRING, MF_SEPARATOR = 0x0, 0x800
    TPM_RETURNCMD, TPM_RIGHTBUTTON = 0x0100, 0x0002
    IDI_APPLICATION = 32512
    CMD_OPEN, CMD_LOGS, CMD_QUIT = 1, 2, 3

    WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                 wintypes.WPARAM, wintypes.LPARAM)

    class WNDCLASSW(ctypes.Structure):
        _fields_ = [
            ("style", wintypes.UINT),
            ("lpfnWndProc", WNDPROC),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", wintypes.HINSTANCE),
            ("hIcon", wintypes.HICON),
            ("hCursor", wintypes.HANDLE),
            ("hbrBackground", wintypes.HBRUSH),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]

    class NOTIFYICONDATAW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("hWnd", wintypes.HWND),
            ("uID", wintypes.UINT),
            ("uFlags", wintypes.UINT),
            ("uCallbackMessage", wintypes.UINT),
            ("hIcon", wintypes.HICON),
            ("szTip", wintypes.WCHAR * 128),
            ("dwState", wintypes.DWORD),
            ("dwStateMask", wintypes.DWORD),
            ("szInfo", wintypes.WCHAR * 256),
            ("uVersion", wintypes.UINT),
            ("szInfoTitle", wintypes.WCHAR * 64),
            ("dwInfoFlags", wintypes.DWORD),
            ("guidItem", ctypes.c_byte * 16),
            ("hBalloonIcon", wintypes.HICON),
        ]

    # 64 位下指针型参数/返回值必须显式声明，否则会被截断成 int
    P = ctypes.POINTER
    user32.RegisterClassW.argtypes = (P(WNDCLASSW),)
    user32.CreateWindowExW.argtypes = (wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                                       wintypes.DWORD, ctypes.c_int, ctypes.c_int,
                                       ctypes.c_int, ctypes.c_int, wintypes.HWND,
                                       wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID)
    user32.CreateWindowExW.restype = wintypes.HWND
    user32.DestroyWindow.argtypes = (wintypes.HWND,)
    user32.DefWindowProcW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    user32.DefWindowProcW.restype = ctypes.c_ssize_t
    user32.GetMessageW.argtypes = (P(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT)
    user32.TranslateMessage.argtypes = (P(wintypes.MSG),)
    user32.DispatchMessageW.argtypes = (P(wintypes.MSG),)
    user32.PostQuitMessage.argtypes = (ctypes.c_int,)
    user32.CreatePopupMenu.restype = wintypes.HMENU
    user32.AppendMenuW.argtypes = (wintypes.HMENU, wintypes.UINT, ctypes.c_size_t, wintypes.LPCWSTR)
    user32.TrackPopupMenu.argtypes = (wintypes.HMENU, wintypes.UINT, ctypes.c_int, ctypes.c_int,
                                      ctypes.c_int, wintypes.HWND, ctypes.c_void_p)
    user32.TrackPopupMenu.restype = ctypes.c_int
    user32.DestroyMenu.argtypes = (wintypes.HMENU,)
    user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
    user32.GetCursorPos.argtypes = (P(wintypes.POINT),)
    user32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    user32.LoadImageW.restype = wintypes.HICON
    user32.LoadImageW.argtypes = (wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
                                  ctypes.c_int, ctypes.c_int, wintypes.UINT)
    user32.LoadIconW.restype = wintypes.HICON
    user32.LoadIconW.argtypes = (wintypes.HINSTANCE, ctypes.c_void_p)
    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    kernel32.GetModuleHandleW.argtypes = (wintypes.LPCWSTR,)
    shell32.Shell_NotifyIconW.argtypes = (wintypes.DWORD, P(NOTIFYICONDATAW))

    hinst = kernel32.GetModuleHandleW(None)
    class_name = f"KuaishouRoiTray_{os.getpid()}"

    def _do(cmd: int, hwnd) -> None:
        if cmd == CMD_OPEN:
            webbrowser.open(url)
        elif cmd == CMD_LOGS:
            try:
                os.startfile(str(LOG_DIR))
            except Exception as e:  # noqa: BLE001
                print(f"[tray] 打开日志目录失败: {e}")
        elif cmd == CMD_QUIT:
            print("[tray] 用户选择退出")
            user32.DestroyWindow(hwnd)

    def _popup_menu(hwnd) -> None:
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, MF_STRING, CMD_OPEN, "打开工作台")
        user32.AppendMenuW(menu, MF_STRING, CMD_LOGS, "打开日志目录")
        user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
        user32.AppendMenuW(menu, MF_STRING, CMD_QUIT, "退出")
        pt = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(hwnd)
        cmd = user32.TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                                    pt.x, pt.y, 0, hwnd, None)
        user32.PostMessageW(hwnd, 0, 0, 0)  # WM_NULL：让菜单正确关闭
        user32.DestroyMenu(menu)
        if cmd:
            _do(cmd, hwnd)

    def _wndproc(hwnd, msg, wparam, lparam):
        if msg == WM_TRAY:
            evt = lparam & 0xFFFF
            if evt in (WM_RBUTTONUP,):
                _popup_menu(hwnd)
            elif evt in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                _do(CMD_OPEN, hwnd)
            return 0
        if msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    wndproc = WNDPROC(_wndproc)
    wc = WNDCLASSW()
    wc.lpfnWndProc = wndproc
    wc.hInstance = hinst
    wc.lpszClassName = class_name
    if not user32.RegisterClassW(ctypes.byref(wc)):
        print(f"[tray] RegisterClassW 失败: {ctypes.get_last_error()}")
        return False

    hwnd = user32.CreateWindowExW(0, class_name, "快手数据工作台", 0, 0, 0, 0, 0,
                                  None, None, hinst, None)
    if not hwnd:
        print(f"[tray] CreateWindowExW 失败: {ctypes.get_last_error()}")
        return False

    hicon = None
    ico = _icon_path()
    if ico.exists():
        hicon = user32.LoadImageW(None, str(ico), IMAGE_ICON, 0, 0,
                                  LR_LOADFROMFILE | LR_DEFAULTSIZE)
    if not hicon:
        # MAKEINTRESOURCE(IDI_APPLICATION)
        hicon = user32.LoadIconW(None, ctypes.c_void_p(IDI_APPLICATION))

    nid = NOTIFYICONDATAW()
    nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
    nid.hWnd = hwnd
    nid.uID = 1
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
    nid.uCallbackMessage = WM_TRAY
    nid.hIcon = hicon
    nid.szTip = f"快手数据工作台 - {url}"[:127]
    if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
        print(f"[tray] Shell_NotifyIcon 失败: {ctypes.get_last_error()}")
        user32.DestroyWindow(hwnd)
        return False

    print("[tray] 托盘图标已启动（双击打开，右键退出）")
    global _TRAY_THREAD_ID, _NOTIFIER
    _TRAY_THREAD_ID = kernel32.GetCurrentThreadId()
    _NOTIFIER = _make_notifier(user32, shell32, NOTIFYICONDATAW, hwnd)
    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))
    _TRAY_THREAD_ID = 0
    _NOTIFIER = None

    shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
    try:
        on_quit()
    except Exception as e:  # noqa: BLE001
        print(f"[tray] 退出回调异常: {e}")
    return True


