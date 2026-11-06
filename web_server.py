(some characters truncated)...
questHandler):
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

