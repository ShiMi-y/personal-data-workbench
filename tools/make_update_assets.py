# -*- coding: utf-8 -*-
"""CI/发版用：为 dist 目录生成升级清单，并在有上一版清单时打包「增量升级包」。

用法（在仓库根目录执行）：
  python tools/make_update_assets.py --dist "dist_ci/快手数据工作台" --version 1.6.0 --out release
  python tools/make_update_assets.py --dist ... --version 1.6.0 --out release \
      --prev-manifest prev/manifest.json --prev-version 1.5.2 \
      --delta-name personal-data-workbench-v1.6.0-from-v1.5.2-win64-delta.zip

产物：
  release/manifest.json               本版完整清单（下次发版用来算增量）
  release/<delta-name>                只含变化文件的增量包
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import update_delta as ud  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dist", required=True, help="PyInstaller 产物目录（含 exe 与 _internal）")
    ap.add_argument("--version", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--prev-manifest", default="")
    ap.add_argument("--prev-version", default="")
    ap.add_argument("--delta-name", default="")
    ap.add_argument("--manifest-only", action="store_true", help="只生成 manifest.json（给旧版本现算清单用）")
    args = ap.parse_args()

    dist = Path(args.dist)
    out = Path(args.out)
    if not dist.is_dir():
        print(f"[delta] 找不到产物目录: {dist}", file=sys.stderr)
        return 2
    out.mkdir(parents=True, exist_ok=True)

    cur = ud.build_manifest(dist, args.version)
    (out / "manifest.json").write_text(json.dumps(cur, ensure_ascii=False, indent=1), encoding="utf-8")
    mb = cur["total_size"] / 1024 / 1024
    print(f"[delta] 清单: {len(cur['files'])} 个文件 / {mb:.1f} MB → {out / 'manifest.json'}")

    if args.manifest_only:
        return 0

    prev_path = Path(args.prev_manifest) if args.prev_manifest else None
    if prev_path and prev_path.is_file() and args.delta_name:
        prev = ud.load_manifest(prev_path)
        delta = ud.make_delta(prev, cur, dist, out / args.delta_name,
                              from_version=args.prev_version or prev.get("version") or "")
        pct = delta["download_bytes"] * 100 / max(1, delta["total_bytes"])
        print(f"[delta] 增量包: {args.delta_name} "
              f"({delta['zip_size'] / 1024 / 1024:.1f} MB，变化 {delta['download_bytes'] / 1024 / 1024:.1f} MB = {pct:.1f}%)")
        print(f"[delta] 变化 {delta['changed']} / 新增 {delta['added']} / 删除 {len(delta['deleted'])}")
    else:
        print("[delta] 没有上一版清单，本次只生成全量包 + 清单（客户端会走全量升级）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
