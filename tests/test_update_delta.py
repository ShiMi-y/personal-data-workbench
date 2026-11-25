# -*- coding: utf-8 -*-
"""update_delta 单测：清单、差异、增量包、叠加与校验。"""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import update_delta as ud


def write(root: Path, rel: str, data) -> Path:
    p = root / rel
    p.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(data, bytes):
        p.write_bytes(data)
    else:
        p.write_text(str(data), encoding="utf-8")
    return p


def make_tree(root: Path, version: str) -> None:
    write(root, "快手数据工作台.exe", f"exe-{version}")
    write(root, "_internal/base_library.zip", "lib-" + version)
    write(root, "_internal/web/index.html", f"<html>{version}</html>")
    write(root, "_internal/web/legacy.html", "<html>old</html>")
    write(root, "_internal/keep.dat", "same-bytes")


class UpdateDeltaTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.v1 = self.root / "v1"
        self.v2 = self.root / "v2"
        make_tree(self.v1, "1.0.0")
        make_tree(self.v2, "1.1.0")

    def tearDown(self):
        self.tmp.cleanup()

    def test_build_manifest_skips_local_files(self):
        write(self.v1, "config.json", "{}")
        write(self.v1, "logs/app.log", "x")
        write(self.v1, "data/cache.json", "{}")
        m = ud.build_manifest(self.v1, "1.0.0")
        paths = [f["path"] for f in m["files"]]
        self.assertIn("快手数据工作台.exe", paths)
        self.assertIn("_internal/web/index.html", paths)
        for bad in ("config.json", "logs/app.log", "data/cache.json"):
            self.assertNotIn(bad, paths)
        self.assertGreater(m["total_size"], 0)

    def test_diff(self):
        (self.v2 / "_internal/web/legacy.html").unlink()
        write(self.v2, "_internal/web/new.html", "<html>new</html>")
        a = ud.build_manifest(self.v1, "1.0.0")
        b = ud.build_manifest(self.v2, "1.1.0")
        d = ud.diff_manifests(a, b)
        self.assertEqual(d["deleted"], ["_internal/web/legacy.html"])
        self.assertEqual(d["added"], ["_internal/web/new.html"])
        self.assertIn("快手数据工作台.exe", d["changed"])
        self.assertNotIn("_internal/keep.dat", d["changed"])

    def test_delta_roundtrip(self):
        a = ud.build_manifest(self.v1, "1.0.0")
        (self.v2 / "_internal/web/legacy.html").unlink()
        b = ud.build_manifest(self.v2, "1.1.0")
        delta_zip = self.root / "delta.zip"
        meta = ud.make_delta(a, b, self.v2, delta_zip, from_version="1.0.0")
        self.assertTrue(delta_zip.is_file())
        self.assertIn("_internal/web/legacy.html", meta["deleted"])

        with zipfile.ZipFile(delta_zip) as z:
            names = z.namelist()
        self.assertIn("manifest.json", names)
        self.assertIn("delta.json", names)
        self.assertIn("快手数据工作台.exe", names)
        self.assertNotIn("_internal/keep.dat", names)      # 未变化的文件不进增量包

        out = self.root / "staged"
        ud.apply_delta(self.v1, delta_zip, out, logger=lambda *_: None)
        self.assertEqual(ud.verify_tree(out, b), [])
        self.assertFalse((out / "_internal/web/legacy.html").exists())
        self.assertEqual((out / "_internal/web/index.html").read_text(encoding="utf-8"), "<html>1.1.0</html>")

    def test_verify_detects_change(self):
        m = ud.build_manifest(self.v1, "1.0.0")
        write(self.v1, "_internal/web/index.html", "<html>tampered</html>")
        problems = ud.verify_tree(self.v1, m)
        self.assertTrue(problems)
        self.assertTrue(any("index.html" in p for p in problems))

    def test_apply_delta_missing_base_file(self):
        a = ud.build_manifest(self.v1, "1.0.0")
        b = ud.build_manifest(self.v2, "1.1.0")
        delta_zip = self.root / "delta2.zip"
        ud.make_delta(a, b, self.v2, delta_zip, from_version="1.0.0")
        broken = self.root / "broken"
        make_tree(broken, "1.0.0")
        (broken / "_internal/keep.dat").unlink()
        with self.assertRaises(ud.DeltaNotApplicable):
            ud.apply_delta(broken, delta_zip, self.root / "out2", logger=lambda *_: None)

    def test_apply_rejects_unsafe_path(self):
        bad = self.root / "bad.zip"
        with zipfile.ZipFile(bad, "w") as z:
            z.writestr("manifest.json", json.dumps({"schema": 1, "files": []}))
            z.writestr("delta.json", json.dumps({"schema": 1, "from": "1", "to": "2", "deleted": []}))
            z.writestr("../evil.txt", "boom")
        with self.assertRaises(ud.DeltaNotApplicable):
            ud.apply_delta(self.v1, bad, self.root / "out3", logger=lambda *_: None)


if __name__ == "__main__":
    unittest.main()


class MultiFromDeltaTest(unittest.TestCase):
    """同一版本线内任意补丁版都能用同一个增量包（并集）"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.dirs = {}
        for name in ("v150", "v153", "v157"):
            d = self.root / name
            d.mkdir()
            write(d, "快手数据工作台.exe", f"exe-{name}")
            write(d, "_internal/web/index.html", "<html>v150</html>" if name == "v150" else f"<html>{name}</html>")
            write(d, "_internal/keep.dat", "same")
            self.dirs[name] = d
        # v153 有、v157 已删除的文件；v150 也有（用于验证删除清单）
        write(self.dirs["v150"], "_internal/old.txt", "old")
        write(self.dirs["v153"], "_internal/old.txt", "old")
        # v150 的 index.html 与 v157 相同，v153 不同：并集必须包含它
        write(self.dirs["v157"], "_internal/web/index.html", "<html>v150</html>")
        # v153 新增且 v157 保留的文件
        write(self.dirs["v153"], "_internal/new.txt", "new")
        write(self.dirs["v157"], "_internal/new.txt", "new")

    def tearDown(self):
        self.tmp.cleanup()

    def _manifest(self, name, version):
        return ud.build_manifest(self.dirs[name], version)

    def test_union_covers_all_old_versions(self):
        m150 = self._manifest("v150", "1.5.0")
        m153 = self._manifest("v153", "1.5.3")
        m157 = self._manifest("v157", "1.5.7")
        diff = ud.diff_manifests_many([m150, m153], m157)
        self.assertIn("快手数据工作台.exe", diff["changed"])
        self.assertIn("_internal/web/index.html", diff["changed"])  # 仅 v153 与新版不同
        self.assertIn("_internal/old.txt", diff["deleted"])

    def test_delta_applies_to_every_old_version(self):
        m150 = self._manifest("v150", "1.5.0")
        m153 = self._manifest("v153", "1.5.3")
        m157 = self._manifest("v157", "1.5.7")
        out = self.root / "delta.zip"
        ud.make_delta([m150, m153], m157, self.dirs["v157"], out, from_label="1.5.x")

        import zipfile
        with zipfile.ZipFile(out) as z:
            meta = json.loads(z.read(ud.DELTA_META_NAME).decode("utf-8"))
        self.assertEqual(meta["from"], "1.5.x")
        self.assertEqual(sorted(meta["from_versions"]), ["1.5.0", "1.5.3"])

        for name in ("v150", "v153"):
            target = self.root / f"out_{name}"
            ud.apply_delta(self.dirs[name], out, target)
            self.assertEqual(ud.verify_tree(target, m157), [], name)


if __name__ == "__main__":
    unittest.main()


class DeltaFromMatchTest(unittest.TestCase):
    def test_exact_and_wildcard(self):
        self.assertTrue(ud.delta_from_match("1.5.6", "1.5.6"))
        self.assertFalse(ud.delta_from_match("1.5.7", "1.5.6"))
        self.assertTrue(ud.delta_from_match("1.5.0", "1.5.x"))
        self.assertTrue(ud.delta_from_match("1.5.9", "1.5.*"))
        self.assertFalse(ud.delta_from_match("1.6.0", "1.5.x"))
        self.assertTrue(ud.delta_from_match("1.5.4", ["1.5.3", "1.5.4"]))
        self.assertTrue(ud.delta_from_match("1.5.4", "1.5.3,1.5.4"))
        self.assertTrue(ud.delta_from_match("1.5.4", "*"))
        self.assertFalse(ud.delta_from_match("1.5.4", ""))
