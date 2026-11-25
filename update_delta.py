# -*- coding: utf-8 -*-
"""升级包的文件级增量：清单、差异、打增量包、叠加与校验。

设计目标（不依赖 Windows，可在 Linux 单测）：
- build_manifest(dir)   生成「每个文件的 sha256 + 大小」清单
- diff_manifests(a, b)  算出 b 相对 a 的变化：changed / added / deleted
- make_delta(...)       只把变化文件打进增量 zip，附上新版完整清单与删除清单
- apply_delta(...)      以旧安装目录为底，叠加增量 zip，产出完整的新目录
- verify_tree(dir, m)   逐文件校验哈希，任何不一致都返回问题清单

增量包只是一个普通 zip：
    manifest.json      新版完整清单（叠加后逐文件校验用）
    delta.json         {"from","to","deleted":[...]}
    <相对路径>          变化/新增的文件
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import time
import zipfile
from pathlib import Path

SCHEMA = 1
MANIFEST_NAME = "manifest.json"
DELTA_META_NAME = "delta.json"

# 安装目录里不属于「程序本体」的东西，不参与清单/校验
IGNORE_DIRS = {"updates", "data", "logs", "__pycache__", ".git"}
IGNORE_FILES = {
    "config.json", "config.example.json", "webdav_secret.dat", "webdav_private.dat",
    "update_cache.json", "release-version.txt",
}


class DeltaNotApplicable(RuntimeError):
    """本地环境与新清单对不上（缺文件/哈希不符），应回退全量包"""


def _rel(path: Path, root: Path) -> str:
    return str(path.relative_to(root)).replace("\\", "/")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def build_manifest(root, version: str = "", skip_dirs=None, skip_files=None) -> dict:
    """遍历目录生成清单（默认跳过配置/数据/日志等本机文件）"""
    root = Path(root)
    skip_dirs = set(skip_dirs or IGNORE_DIRS)
    skip_files = set(skip_files or IGNORE_FILES)
    files = []
    total = 0
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if d not in skip_dirs)
        for name in sorted(filenames):
            p = Path(dirpath) / name
            if not p.is_file():
                continue
            rel = _rel(p, root)
            if rel in skip_files or name in skip_files:
                continue
            size = p.stat().st_size
            files.append({"path": rel, "size": size, "sha256": sha256_file(p)})
            total += size
    return {
        "schema": SCHEMA,
        "version": version,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "total_size": total,
        "files": files,
    }


def delta_from_match(current: str, spec) -> bool:
    """增量包声明的来源是否匹配当前版本

    spec 支持：精确版本（1.5.6）、版本线通配（1.5.x / 1.5.*）、
    逗号分隔字符串或列表（["1.5.0", "1.5.1"]）、全量通配（*）。
    """
    cur = str(current or "").strip().lstrip("vV")
    if not cur:
        return False
    if isinstance(spec, (list, tuple, set)):
        return any(delta_from_match(cur, x) for x in spec)
    raw = str(spec or "").strip().lstrip("vV")
    if not raw:
        return False
    if "," in raw:
        return any(delta_from_match(cur, x) for x in raw.split(","))
    if raw == "*":
        return True
    if raw.endswith(".x") or raw.endswith(".*"):
        base = raw[:-2].rstrip(".")
        return bool(base) and (cur == base or cur.startswith(base + "."))
    return cur == raw


def load_manifest(source) -> dict:
    """接受 dict / 路径 / 文件对象"""
    if isinstance(source, dict):
        return source
    if hasattr(source, "read"):
        return json.loads(source.read().decode("utf-8") if isinstance(source.read(0), bytes) else source.read())
    data = json.loads(Path(source).read_text(encoding="utf-8"))
    if not isinstance(data, dict) or "files" not in data:
        raise ValueError("清单格式不正确（缺 files）")
    return data


def _index(manifest: dict) -> dict:
    return {f["path"]: f for f in manifest.get("files") or [] if f.get("path")}


def diff_manifests(prev: dict, cur: dict) -> dict:
    a, b = _index(prev), _index(cur)
    changed = [p for p in sorted(b) if p in a and a[p].get("sha256") != b[p].get("sha256")]
    added = [p for p in sorted(b) if p not in a]
    deleted = [p for p in sorted(a) if p not in b]
    return {
        "changed": changed,
        "added": added,
        "deleted": deleted,
        "download_bytes": sum(int(b[p].get("size") or 0) for p in changed + added),
        "total_bytes": sum(int(f.get("size") or 0) for f in cur.get("files") or []),
    }


def diff_manifests_many(olds, cur: dict) -> dict:
    """对多个旧版本清单取并集：任一旧版本不同的文件都要打进增量包

    用于「同一版本线内任意补丁版都能升级」的场景（如 1.5.x → 1.5.7）。
    """
    olds = [olds] if isinstance(olds, dict) else [m for m in olds if isinstance(m, dict)]
    b = _index(cur)
    changed: set[str] = set()
    added: set[str] = set()
    deleted: set[str] = set()
    for prev in olds:
        d = diff_manifests(prev, cur)
        changed.update(d["changed"])
        added.update(d["added"])
        deleted.update(d["deleted"])
    changed -= added
    return {
        "changed": sorted(changed),
        "added": sorted(added),
        "deleted": sorted(deleted),
        "download_bytes": sum(int(b[p].get("size") or 0) for p in changed | added if p in b),
        "total_bytes": sum(int(f.get("size") or 0) for f in cur.get("files") or []),
    }


def make_delta(prev_manifest, cur_manifest: dict, cur_root, out_zip,
               from_version="", from_label: str = "") -> dict:
    """生成增量 zip：新版清单 + 删除清单 + 变化文件

    prev_manifest 可以是单个清单，也可以是多个旧版本清单的列表（取并集，
    这样一个增量包就能覆盖同一版本线内的所有补丁版）。
    """
    cur_root = Path(cur_root)
    out_zip = Path(out_zip)
    olds = [prev_manifest] if isinstance(prev_manifest, dict) else list(prev_manifest)
    versions = [str(m.get("version") or "") for m in olds if isinstance(m, dict) and m.get("version")]
    diff = diff_manifests_many(olds, cur_manifest)
    cur_index = _index(cur_manifest)
    out_zip.parent.mkdir(parents=True, exist_ok=True)
    meta = {
        "schema": SCHEMA,
        "from": from_label or from_version or (versions[0] if len(versions) == 1 else versions),
        "from_versions": versions,
        "to": cur_manifest.get("version") or "",
        "deleted": diff["deleted"],
        "changed": len(diff["changed"]),
        "added": len(diff["added"]),
        "download_bytes": diff["download_bytes"],
        "total_bytes": diff["total_bytes"],
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.writestr(MANIFEST_NAME, json.dumps(cur_manifest, ensure_ascii=False, indent=1))
        z.writestr(DELTA_META_NAME, json.dumps(meta, ensure_ascii=False, indent=1))
        for rel in diff["changed"] + diff["added"]:
            src = cur_root / rel
            if not src.is_file():
                raise FileNotFoundError(f"清单里的文件不存在: {rel}")
            z.write(src, rel)
    return {**meta, "zip": str(out_zip), "zip_size": out_zip.stat().st_size}


def _safe_rel(rel: str) -> str:
    rel = str(rel).replace("\\", "/").lstrip("/")
    if not rel or rel.startswith("../") or "/../" in rel or rel in (".", ".."):
        raise DeltaNotApplicable(f"增量包里有可疑路径: {rel}")
    return rel


def read_delta(delta_zip) -> tuple[dict, dict]:
    with zipfile.ZipFile(delta_zip) as z:
        manifest = json.loads(z.read(MANIFEST_NAME).decode("utf-8"))
        meta = json.loads(z.read(DELTA_META_NAME).decode("utf-8"))
    return manifest, meta


def apply_delta(base_dir, delta_zip, out_dir, logger=print) -> dict:
    """以 base_dir（当前安装）为底，叠加增量包到 out_dir（完整的新目录）"""
    base_dir = Path(base_dir)
    out_dir = Path(out_dir)
    manifest, meta = read_delta(delta_zip)
    files = manifest.get("files") or []
    if out_dir.exists():
        shutil.rmtree(out_dir, ignore_errors=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 增量包里已经带上的文件（新增/变化的文件），不要求旧安装里存在
    with zipfile.ZipFile(delta_zip) as z:
        packed = {_safe_rel(i.filename) for i in z.infolist()
                  if i.filename not in (MANIFEST_NAME, DELTA_META_NAME) and not i.filename.endswith("/")}

    missing = []
    copied = 0
    for f in files:
        rel = _safe_rel(f["path"])
        src = base_dir / rel
        if not src.is_file():
            if rel not in packed:
                missing.append(rel)
            continue
        dst = out_dir / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        copied += 1
    if missing:
        raise DeltaNotApplicable("本地缺少文件: " + ", ".join(missing[:5])
                                 + (f" 等 {len(missing)} 个" if len(missing) > 5 else ""))

    overlay = 0
    with zipfile.ZipFile(delta_zip) as z:
        for info in z.infolist():
            name = info.filename
            if name in (MANIFEST_NAME, DELTA_META_NAME) or name.endswith("/"):
                continue
            rel = _safe_rel(name)
            dst = out_dir / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            with z.open(info) as src, dst.open("wb") as fh:
                shutil.copyfileobj(src, fh)
            overlay += 1

    removed = 0
    for rel in meta.get("deleted") or []:
        p = out_dir / _safe_rel(rel)
        if p.is_file():
            p.unlink()
            removed += 1

    logger(f"[update] 增量叠加完成: 复制 {copied} 个文件、覆盖 {overlay} 个、删除 {removed} 个")
    return {"manifest": manifest, "meta": meta, "copied": copied, "overlay": overlay, "removed": removed}


def verify_tree(root, manifest: dict, limit: int = 5) -> list[str]:
    """逐文件校验，返回问题描述（空列表 = 全部通过）"""
    root = Path(root)
    problems: list[str] = []
    for f in manifest.get("files") or []:
        rel = _safe_rel(f["path"])
        p = root / rel
        if not p.is_file():
            problems.append(f"缺少 {rel}")
        elif int(f.get("size") or 0) != p.stat().st_size:
            problems.append(f"大小不符 {rel}")
        elif sha256_file(p) != f.get("sha256"):
            problems.append(f"哈希不符 {rel}")
        if len(problems) >= limit:
            problems.append("…")
            break
    return problems
