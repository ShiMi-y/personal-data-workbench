# 个人定制数据工作台

本地运行的数据工作台，用于按日抓取投放平台数据、核对后回填月度 Excel。

## 开发与测试

纯逻辑（合并、分页、配置）可以在 WSL/Linux 上直接测试，不依赖 Windows 和 Playwright：

```bash
python3 -m unittest discover -s tests -v
```

Windows 端到端（真实登录、托盘、DPAPI、自动升级）需要 Windows + Chrome；
`run_tests.bat` 跑回归，`build_windows.bat` 打包 exe：

```bat
py -m pip install openpyxl playwright cryptography
copy config.example.json config.json
py web_server.py
```

详细安装、配置、口径与发布流程见 `使用说明.md` 与 `发布规则.md`。
