import argparse
import os
import signal
import subprocess
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
VIEWER_DIR = ROOT_DIR / "apps" / "viewer"
BRIDGE_DIR = ROOT_DIR / "apps" / "bridge"

if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from config.settings import DATA_DIR, UPLOAD_FOLDER  # noqa: E402


def _start_app(app_dir: Path, port: int) -> subprocess.Popen:
    env = os.environ.copy()
    env["PORT"] = str(port)
    return subprocess.Popen([sys.executable, "app.py"], cwd=str(app_dir), env=env)


def _wait_for_all(processes: list[subprocess.Popen]) -> int:
    try:
        exit_code = 0
        for proc in processes:
            code = proc.wait()
            if code != 0 and exit_code == 0:
                exit_code = code
        return exit_code
    except KeyboardInterrupt:
        for proc in processes:
            if proc.poll() is None:
                proc.terminate()
        for proc in processes:
            if proc.poll() is None:
                try:
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    proc.kill()
        return 130


def _run_viewer(host: str = "0.0.0.0", port: int = 5000, debug: bool = True) -> None:
    """同一プロセスで viewer を起動（DB・upload ディレクトリを用意）。"""
    from apps.viewer.app import app  # noqa: E402
    from apps.viewer.app import ensure_db_schema, init_db  # noqa: E402

    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    init_db()
    ensure_db_schema()
    app.run(host=host, port=port, debug=debug)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run asprova apps from unified platform root."
    )
    parser.add_argument(
        "target",
        nargs="?",
        choices=["viewer", "bridge", "all"],
        default="all",
        help="all: viewer + bridge (default) | viewer: in-process viewer only | bridge: bridge only",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host for in-process viewer (default: 0.0.0.0)",
    )
    parser.add_argument(
        "--viewer-port",
        type=int,
        default=5000,
        help="Port for viewer (default: 5000)",
    )
    parser.add_argument(
        "--bridge-port",
        type=int,
        default=5001,
        help="Port for bridge subprocess (default: 5001)",
    )
    parser.add_argument(
        "--no-debug",
        action="store_true",
        help="Disable Flask debug for in-process viewer",
    )

    args = parser.parse_args()

    if args.target == "viewer":
        _run_viewer(
            host=args.host,
            port=args.viewer_port,
            debug=not args.no_debug,
        )
        return 0

    targets = []
    if args.target in ("bridge", "all"):
        targets.append(("bridge", BRIDGE_DIR, args.bridge_port))
    if args.target == "all":
        targets.insert(0, ("viewer", VIEWER_DIR, args.viewer_port))

    missing = [name for name, app_dir, _ in targets if not (app_dir / "app.py").exists()]
    if missing:
        print(f"Missing app.py for: {', '.join(missing)}", file=sys.stderr)
        return 1

    processes = []
    for name, app_dir, port in targets:
        print(f"Starting {name} on port {port} (subprocess) ...")
        processes.append(_start_app(app_dir, port))

    return _wait_for_all(processes)


if __name__ == "__main__":
    if os.name == "nt":
        signal.signal(signal.SIGINT, signal.default_int_handler)
    raise SystemExit(main())
