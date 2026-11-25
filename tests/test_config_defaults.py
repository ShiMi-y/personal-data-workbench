# -*- coding: utf-8 -*-
"""配置默认值：版本文件路径（update_webdav_path）的默认与旧值迁移"""
import json
import tempfile
import unittest
from pathlib import Path

import kx_core as kx


def _load(tmp: Path, data: dict) -> tuple[dict, dict]:
    cfg_path = tmp / "config.json"
    cfg_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    cfg = kx.load_config(cfg_path)
    return cfg, json.loads(cfg_path.read_text(encoding="utf-8"))


class UpdateWebdavPathTest(unittest.TestCase):
    def test_default_when_missing(self):
        with tempfile.TemporaryDirectory() as d:
            cfg, saved = _load(Path(d), {"accounts": []})
            self.assertEqual(cfg["update_webdav_path"], "/kuaishou/version.json")
            self.assertEqual(saved["update_webdav_path"], "/kuaishou/version.json")

    def test_legacy_default_is_migrated(self):
        with tempfile.TemporaryDirectory() as d:
            cfg, saved = _load(Path(d), {"accounts": [], "update_webdav_path": "/version.json"})
            self.assertEqual(cfg["update_webdav_path"], "/kuaishou/version.json")
            self.assertEqual(saved["update_webdav_path"], "/kuaishou/version.json")

    def test_custom_path_is_kept(self):
        with tempfile.TemporaryDirectory() as d:
            cfg, saved = _load(Path(d), {"accounts": [], "update_webdav_path": "/other/version.json"})
            self.assertEqual(cfg["update_webdav_path"], "/other/version.json")
            self.assertEqual(saved["update_webdav_path"], "/other/version.json")


if __name__ == "__main__":
    unittest.main()
