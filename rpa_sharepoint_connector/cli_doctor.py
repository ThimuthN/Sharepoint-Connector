"""Diagnostics command for local SharePoint connector execution."""
import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.request import urlopen

from .token_store import TokenStore


def cmd_doctor(args, unlink_if_exists, print_actionable_error) -> None:
    """Run preflight diagnostics for local/bot execution."""
    profile_name = args.profile or "default"
    store_dir = args.store_dir
    as_json = bool(getattr(args, "json", False))
    checks = []

    try:
        store = TokenStore(store_dir=store_dir)
        store_path = Path(store.store_dir)

        try:
            store_path.mkdir(parents=True, exist_ok=True)
            probe = store_path / ".doctor_write_probe"
            probe.write_text("ok", encoding="utf-8")
            unlink_if_exists(probe)
            checks.append(
                {
                    "name": "token_store_writable",
                    "ok": True,
                    "detail": str(store_path),
                }
            )
        except Exception as exc:
            checks.append(
                {
                    "name": "token_store_writable",
                    "ok": False,
                    "detail": str(exc),
                }
            )

        class _Handler(BaseHTTPRequestHandler):
            def log_message(self, format, *args):
                return

            def do_GET(self):
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"ok")

        try:
            server = HTTPServer(("127.0.0.1", 0), _Handler)
            port = server.server_port
            thread = threading.Thread(target=server.handle_request, daemon=True)
            thread.start()
            body = urlopen(f"http://127.0.0.1:{port}", timeout=5).read().decode("utf-8")
            server.server_close()
            checks.append(
                {
                    "name": "localhost_callback",
                    "ok": body == "ok",
                    "detail": f"http://127.0.0.1:{port}",
                }
            )
        except Exception as exc:
            checks.append(
                {
                    "name": "localhost_callback",
                    "ok": False,
                    "detail": str(exc),
                }
            )

        profile_data = store.load_profile(profile_name)
        if profile_data:
            missing = []
            for field in ("client_id", "tenant_id", "refresh_token"):
                if not profile_data.get(field):
                    missing.append(field)
            checks.append(
                {
                    "name": "profile_exists",
                    "ok": True,
                    "detail": profile_name,
                }
            )
            checks.append(
                {
                    "name": "profile_required_fields",
                    "ok": len(missing) == 0,
                    "detail": "missing: " + ", ".join(missing) if missing else "ok",
                }
            )
            checks.append(
                {
                    "name": "profile_target_bound",
                    "ok": bool(profile_data.get("drive_id")),
                    "detail": profile_data.get("drive_name") or "(unbound)",
                }
            )
        else:
            checks.append(
                {
                    "name": "profile_exists",
                    "ok": False,
                    "detail": f"{profile_name} not found",
                }
            )

        ok_all = all(check["ok"] for check in checks)
        output = {
            "profile": profile_name,
            "ok": ok_all,
            "checks": checks,
        }

        if as_json:
            print(json.dumps(output, indent=2))
        else:
            print(f"\nDoctor report for profile: {profile_name}")
            print("=" * 60)
            for check in checks:
                status = "OK" if check["ok"] else "FAIL"
                print(f"{status:5} {check['name']}: {check['detail']}")
            print()

        if not ok_all:
            sys.exit(1)
    except Exception as error:
        print_actionable_error(error, command="doctor", profile=profile_name)
