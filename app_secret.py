# -*- coding: utf-8 -*-
"""本机密文存储：WebDAV 连接信息用 Windows DPAPI 加密后单独保存

- 文件：<exe 目录>/webdav_secret.dat（二进制，不进安装包、不进仓库）
- 加密：CryptProtectData（绑定当前 Windows 用户），换机器/换用户无法解密
- 使用时才解密，界面上不回显密码
"""
from __future__ import annotations

import ctypes
import json
import os
import sys
from ctypes import wintypes
from pathlib import Path

from app_runtime import app_dir

SECRET_PATH = app_dir() / "webdav_secret.dat"
PRIVATE_PATH = app_dir() / "webdav_private.dat"
PRIVATE_MAGIC = b"KXDAV2\n"
PRIVATE_KDF_ITER = 200_000
_PRIVATE_CACHE: dict | None = None


class PrivatePassphraseRequired(RuntimeError):
    """私有 WebDAV 文件需要口令解锁"""


def private_status() -> dict:
    if not PRIVATE_PATH.exists():
        return {"ready": False, "legacy": False, "unlocked": False}
    legacy = not PRIVATE_PATH.read_bytes().startswith(PRIVATE_MAGIC)
    return {"ready": True, "legacy": legacy, "unlocked": bool(_PRIVATE_CACHE)}


def _fernet(passphrase: str, salt: bytes):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt,
                     iterations=PRIVATE_KDF_ITER)
    key = kdf.derive(passphrase.encode("utf-8"))
    import base64
    return Fernet(base64.urlsafe_b64encode(key))


def _decode_private(passphrase: str) -> dict:
    raw = PRIVATE_PATH.read_bytes()
    if not raw.startswith(PRIVATE_MAGIC):
        # 旧版 DPAPI 文件（仅本机可用），保持兼容
        return json.loads(_dpapi(raw, protect=False).decode("utf-8"))
    salt, token = raw[len(PRIVATE_MAGIC):len(PRIVATE_MAGIC) + 16], raw[len(PRIVATE_MAGIC) + 16:]
    try:
        plain = _fernet(passphrase, salt).decrypt(token)
    except Exception as e:  # noqa: BLE001
        raise ValueError("口令错误，或私有文件已损坏") from e
    return json.loads(plain.decode("utf-8"))


def unlock_private(passphrase: str) -> dict:
    """校验口令并缓存到内存（不落盘）"""
    global _PRIVATE_CACHE
    j = _decode_private(passphrase)
    if not j.get("url"):
        raise ValueError("私有文件内容不完整（缺少 WebDAV 地址）")
    _PRIVATE_CACHE = j
    return j


def lock_private():
    global _PRIVATE_CACHE
    _PRIVATE_CACHE = None


class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(data: bytes):
    buf = ctypes.create_string_buffer(data, len(data))
    return _DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _dpapi(data: bytes, protect: bool) -> bytes:
    if sys.platform != "win32":
        raise RuntimeError("仅支持 Windows 加密存储")
    blob, _buf = _blob(data)
    out = _DATA_BLOB()
    crypt32 = ctypes.windll.crypt32
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    ok = fn(ctypes.byref(blob), None, None, None, None, 0x1, ctypes.byref(out))
    if not ok:
        raise RuntimeError("DPAPI 加解密失败（可能换用户或文件损坏）")
    try:
        raw = ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)
    return raw


PW_PREFIX = "dpapi:v1:"


def is_protected(value) -> bool:
    return str(value or "").startswith(PW_PREFIX)


def protect_text(text: str) -> str:
    """把明文（如账号密码）加密成 dpapi:v1:<base64>；已加密则原样返回"""
    s = str(text or "")
    if not s or is_protected(s):
        return s
    import base64
    blob = _dpapi(s.encode("utf-8"), protect=True)
    return PW_PREFIX + base64.b64encode(blob).decode("ascii")


def unprotect_text(value) -> str:
    """解密 dpapi:v1: 前缀的密文；无前缀视为旧版明文原样返回"""
    s = str(value or "")
    if not s or not is_protected(s):
        return s
    try:
        import base64
        raw = _dpapi(base64.b64decode(s[len(PW_PREFIX):]), protect=False)
        return raw.decode("utf-8")
    except Exception:
        return ""


def save_webdav(url: str, username: str, password: str) -> Path:
    data = json.dumps({"url": url, "username": username, "password": password},
                      ensure_ascii=False).encode("utf-8")
    SECRET_PATH.write_bytes(_dpapi(data, protect=True))
    return SECRET_PATH


def load_webdav() -> dict | None:
    if not SECRET_PATH.exists():
        return None
    try:
        raw = _dpapi(SECRET_PATH.read_bytes(), protect=False)
        j = json.loads(raw.decode("utf-8"))
    except Exception:
        return None
    if not j.get("url"):
        return None
    return j


def clear_webdav() -> None:
    try:
        SECRET_PATH.unlink(missing_ok=True)
    except Exception:
        pass


def load_private() -> dict | None:
    """私有模式：读取随程序分发的 webdav_private.dat（口令加密，需先解锁）"""
    global _PRIVATE_CACHE
    if not PRIVATE_PATH.exists():
        return None
    if _PRIVATE_CACHE:
        return _PRIVATE_CACHE
    if not PRIVATE_PATH.read_bytes().startswith(PRIVATE_MAGIC):
        # 旧版 DPAPI 文件，本机可直接解密
        try:
            j = json.loads(_dpapi(PRIVATE_PATH.read_bytes(), protect=False).decode("utf-8"))
        except Exception:
            return None
        return j if j.get("url") else None
    raise PrivatePassphraseRequired("私有 WebDAV 文件需要口令解锁")


def export_private(passphrase: str) -> Path:
    """把已保存的 WebDAV 连接用口令加密后写入私有文件（可跨设备）"""
    global _PRIVATE_CACHE
    if not str(passphrase or "").strip():
        raise ValueError("口令不能为空")
    if not SECRET_PATH.exists():
        raise RuntimeError("还没有可用的 WebDAV 连接，请先在 WebDAV 方式里保存")
    conf = load_webdav()
    if not conf:
        raise RuntimeError("无法读取本机 WebDAV 连接，请重新保存")
    salt = os.urandom(16)
    token = _fernet(passphrase, salt).encrypt(
        json.dumps(conf, ensure_ascii=False).encode("utf-8"))
    PRIVATE_PATH.write_bytes(PRIVATE_MAGIC + salt + token)
    _PRIVATE_CACHE = conf
    return PRIVATE_PATH


def clear_private():
    global _PRIVATE_CACHE
    _PRIVATE_CACHE = None
    try:
        PRIVATE_PATH.unlink(missing_ok=True)
    except Exception:
        pass

