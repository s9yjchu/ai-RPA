"""B2C RPA GCP 에이전트 — Windows 서비스 래퍼 (pywin32).

설치:   python -m src.agent_service install
시작:   python -m src.agent_service start
중지:   python -m src.agent_service stop
제거:   python -m src.agent_service remove
상태:   python -m src.agent_service status
"""

from __future__ import annotations

import sys
import threading
from pathlib import Path

import servicemanager
import win32event
import win32service
import win32serviceutil

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SERVICE_NAME = "B2CRPAAgent"
SERVICE_DISPLAY = "B2C RPA GCP Agent"
SERVICE_DESC = "GCP Pub/Sub 메시지를 수신하여 B2C 고객지표 RPA를 실행합니다."


class AgentService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY
    _svc_description_ = SERVICE_DESC

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self._stop_event = win32event.CreateEvent(None, 0, 0, None)
        self._agent_thread: threading.Thread | None = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self._stop_event)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )
        self._run()

    def _run(self):
        # Add project root to path so src package is importable
        sys.path.insert(0, str(PROJECT_ROOT))

        from src.config import load_config
        from src.logger import setup_logging
        from src.gcp_agent import run_agent

        try:
            config = load_config()
        except RuntimeError as exc:
            servicemanager.LogErrorMsg(f"설정 오류: {exc}")
            return

        setup_logging(config.runtime.logs_dir)

        # Run agent in a daemon thread so SvcStop can interrupt it
        self._agent_thread = threading.Thread(
            target=run_agent, args=(config,), daemon=True
        )
        self._agent_thread.start()

        # Wait until stop signal
        win32event.WaitForSingleObject(self._stop_event, win32event.INFINITE)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STOPPED,
            (self._svc_name_, ""),
        )


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Called by Windows SCM — hand off to pywin32
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(AgentService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(AgentService)
