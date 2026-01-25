import os
import sys
import socket
import time
import threading
import webbrowser
import traceback
from pathlib import Path


def _find_free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _wait_port(host: str, port: int, timeout_s: int = 40) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.2)
    return False


def _log(base_dir: Path, msg: str) -> None:
    try:
        (base_dir / "launcher.log").write_text(msg, encoding="utf-8")
    except Exception:
        pass


def _resolve_app_py(base_dir: Path) -> Path:
    """
    Cherche app.py dans les emplacements possibles (dev + PyInstaller).
    """
    candidates = [
        base_dir / "app.py",
        Path(__file__).parent / "app.py",
    ]

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "app.py")

    # certains builds mettent les datas dans _internal
    candidates += [
        base_dir / "_internal" / "app.py",
        (Path(meipass) / "_internal" / "app.py") if meipass else None,
    ]
    candidates = [c for c in candidates if c is not None]

    for p in candidates:
        if p.exists():
            return p

    raise FileNotFoundError("app.py introuvable. Cherché dans:\n" + "\n".join(str(p) for p in candidates))


def main() -> None:
    base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

    # Anti-chauffe / anti-reload
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_RUN_ON_SAVE"] = "false"
    os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"

    try:
        app_path = _resolve_app_py(base_dir)
    except Exception:
        _log(base_dir, "CRASH (app.py):\n" + traceback.format_exc())
        return

    host = "127.0.0.1"
    port = _find_free_port()
    url = f"http://{host}:{port}"

    crash = {"trace": None}

    def run_streamlit():
        try:
            from streamlit.web import bootstrap

            args = [
                f"--server.address={host}",
                f"--server.port={port}",
                "--server.headless=true",
                "--server.runOnSave=false",
                "--server.fileWatcherType=none",
                "--browser.gatherUsageStats=false",
                "--server.enableCORS=false",
                "--server.enableXsrfProtection=false",
            ]

            # compat signatures
            try:
                bootstrap.run(str(app_path), "", args, {})
            except TypeError:
                bootstrap.run(str(app_path), False, args, {})

        except Exception:
            crash["trace"] = traceback.format_exc()

    t = threading.Thread(target=run_streamlit, daemon=True)
    t.start()

    if _wait_port(host, port, timeout_s=45):
        webbrowser.open(url)
        _log(base_dir, f"READY\nURL={url}\nAPP={app_path}\nBASE={base_dir}\n")
    else:
        _log(base_dir, (crash["trace"] or "TIMEOUT sans trace") + f"\nURL={url}\nAPP={app_path}\nBASE={base_dir}\n")

    t.join()


if __name__ == "__main__":
    main()
