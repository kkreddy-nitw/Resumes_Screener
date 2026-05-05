#!/usr/bin/env python3
"""Windows service wrapper for Resumes Screener."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

try:
    import servicemanager
    import win32event
    import win32service
    import win32serviceutil
except ImportError as exc:  # pragma: no cover - Windows only
    raise SystemExit("pywin32 is required on Windows. Run: python -m pip install -r requirements.txt") from exc


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "windows_service_config.json"
LOG_DIR = BASE_DIR / "logs"


def _load_config() -> dict[str, str]:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {
        "ProjectRoot": str(BASE_DIR),
        "PythonExe": sys.executable,
        "Host": "0.0.0.0",
        "Port": "80",
    }


def _terminate(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=20)
    except subprocess.TimeoutExpired:
        process.kill()


class ResumesScreenerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ResumesScreener"
    _svc_display_name_ = "Resumes Screener"
    _svc_description_ = "Runs the Resumes Screener Flask and Waitress web application."

    def __init__(self, args: list[str]) -> None:
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.process: subprocess.Popen[bytes] | None = None
        self.log_handle = None

    def SvcStop(self) -> None:
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        _terminate(self.process)
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self) -> None:
        servicemanager.LogInfoMsg("Starting Resumes Screener service")
        config = _load_config()
        project_root = Path(config.get("ProjectRoot", str(BASE_DIR))).resolve()
        python_exe = config.get("PythonExe") or sys.executable

        LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.log_handle = open(LOG_DIR / "resumes_screener_service.log", "ab", buffering=0)

        env = os.environ.copy()
        env["RESUMES_SCREENER_HOST"] = str(config.get("Host", "0.0.0.0"))
        env["RESUMES_SCREENER_PORT"] = str(config.get("Port", "80"))

        self.process = subprocess.Popen(
            [python_exe, str(project_root / "resumes_screener.py")],
            cwd=str(project_root),
            env=env,
            stdout=self.log_handle,
            stderr=subprocess.STDOUT,
        )

        while True:
            rc = win32event.WaitForSingleObject(self.stop_event, 5000)
            if rc == win32event.WAIT_OBJECT_0:
                break
            if self.process.poll() is not None:
                servicemanager.LogErrorMsg(
                    f"Resumes Screener exited with code {self.process.returncode}"
                )
                break

        _terminate(self.process)
        if self.log_handle:
            self.log_handle.close()
        servicemanager.LogInfoMsg("Stopped Resumes Screener service")


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(ResumesScreenerService)
