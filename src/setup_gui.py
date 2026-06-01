"""고객지표 RPA 설정 마법사 — 비개발자용 초기 설정 GUI.

실행: python -m src.setup_gui
     (또는 setup.bat 에서 자동으로 실행됨)
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import (
    BooleanVar, Entry, Frame, Label, Scrollbar, StringVar, Text,
    Tk, Toplevel, filedialog, messagebox, ttk,
)

# 프로젝트 루트 (src/ 의 상위)
ROOT = Path(__file__).parent.parent.resolve()
ENV_PATH = ROOT / ".env"
CREDS_PATH = ROOT / "credentials.json"
TOKEN_PATH = ROOT / "token.json"
BAT_PATH = ROOT / "run_rpa.bat"

TASK_DAILY = "고객지표RPA_daily"
TASK_MONTHLY = "고객지표RPA_monthly"


# ── .env 읽기/쓰기 ────────────────────────────────────────────────────

def _read_env() -> dict[str, str]:
    result: dict[str, str] = {}
    if not ENV_PATH.exists():
        return result
    for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, _, v = line.partition("=")
            result[k.strip()] = v.strip()
    return result


def _write_env(updates: dict[str, str]) -> None:
    """기존 .env 파일에서 지정 키만 업데이트합니다 (주석·순서 보존)."""
    if not ENV_PATH.exists():
        shutil.copy(ROOT / ".env.example", ENV_PATH)

    lines = ENV_PATH.read_text(encoding="utf-8").splitlines()
    written = set()
    new_lines = []

    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            k = stripped.split("=", 1)[0].strip()
            if k in updates:
                new_lines.append(f"{k}={updates[k]}")
                written.add(k)
                continue
        new_lines.append(line)

    # 아직 파일에 없는 키는 끝에 추가
    for k, v in updates.items():
        if k not in written:
            new_lines.append(f"{k}={v}")

    ENV_PATH.write_text("\n".join(new_lines) + "\n", encoding="utf-8")


# ── 상태 체크 함수들 ──────────────────────────────────────────────────

def check_packages() -> bool:
    try:
        import playwright  # noqa: F401
        import gspread     # noqa: F401
        return True
    except ImportError:
        return False


def check_spc_login() -> bool:
    env = _read_env()
    return bool(env.get("OLAP_ID")) and bool(env.get("OLAP_PW"))


def check_google_auth() -> bool:
    if not TOKEN_PATH.exists() or not CREDS_PATH.exists():
        return False
    try:
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file(
            str(TOKEN_PATH),
            ["https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/gmail.send"],
        )
        return creds is not None and (creds.valid or creds.refresh_token is not None)
    except Exception:
        return False


def check_email() -> bool:
    env = _read_env()
    return bool(env.get("REPORT_SENDER")) and bool(env.get("REPORT_RECIPIENTS"))


def check_scheduler() -> bool:
    try:
        result = subprocess.run(
            ["schtasks", "/query", "/tn", TASK_DAILY],
            capture_output=True, timeout=10,
        )
        return result.returncode == 0
    except Exception:
        return False


# ── 메인 윈도우 ───────────────────────────────────────────────────────

class SetupApp:
    def __init__(self, root: Tk):
        self.root = root
        root.title("고객지표 RPA 설정")
        root.resizable(False, False)
        root.geometry("500x400")

        # 항목별 상태 변수
        self._status: dict[str, BooleanVar] = {
            "packages": BooleanVar(),
            "spc":      BooleanVar(),
            "google":   BooleanVar(),
            "email":    BooleanVar(),
            "scheduler": BooleanVar(),
        }

        self._build_ui()
        self.refresh_status()

    # ── UI 구성 ───────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 20, "pady": 6}

        header = Label(
            self.root,
            text="고객지표 RPA 설정 마법사",
            font=("맑은 고딕", 14, "bold"),
        )
        header.pack(pady=(20, 4))

        Label(
            self.root,
            text="아래 4가지 항목을 순서대로 설정하세요.",
            font=("맑은 고딕", 10),
            fg="#555555",
        ).pack()

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=20, pady=10)

        items_frame = Frame(self.root)
        items_frame.pack(fill="x", **pad)

        # (key, label, button_text, handler)
        rows = [
            ("packages", "패키지 및 브라우저 설치됨",  None,        None),
            ("spc",      "SPC 로그인 정보",            "설정하기 ▶", self._open_spc_dialog),
            ("google",   "Google 인증",                "인증하기 ▶", self._open_google_dialog),
            ("email",    "이메일 알림 설정",            "설정하기 ▶", self._open_email_dialog),
            ("scheduler","자동 실행 등록",              "등록하기 ▶", self._register_scheduler),
        ]

        self._status_labels: dict[str, Label] = {}

        for key, label_text, btn_text, handler in rows:
            row = Frame(items_frame)
            row.pack(fill="x", pady=3)

            status_lbl = Label(row, text="⬜", font=("맑은 고딕", 12), width=3)
            status_lbl.pack(side="left")
            self._status_labels[key] = status_lbl

            Label(row, text=label_text, font=("맑은 고딕", 11), anchor="w").pack(
                side="left", fill="x", expand=True
            )

            if btn_text and handler:
                ttk.Button(row, text=btn_text, command=handler, width=12).pack(side="right")

        ttk.Separator(self.root, orient="horizontal").pack(fill="x", padx=20, pady=10)

        btn_frame = Frame(self.root)
        btn_frame.pack(pady=(0, 16))

        ttk.Button(
            btn_frame, text="테스트 실행 (쓰기 없음)",
            command=self._run_test, width=22,
        ).pack(side="left", padx=8)

        self._run_btn = ttk.Button(
            btn_frame, text="지금 실행",
            command=self._run_now, width=12, state="disabled",
        )
        self._run_btn.pack(side="left", padx=8)

    # ── 상태 갱신 ─────────────────────────────────────────────────────

    def refresh_status(self):
        checks = {
            "packages": check_packages,
            "spc":      check_spc_login,
            "google":   check_google_auth,
            "email":    check_email,
            "scheduler": check_scheduler,
        }
        all_ok = True
        for key, fn in checks.items():
            ok = fn()
            self._status[key].set(ok)
            lbl = self._status_labels[key]
            if ok:
                lbl.config(text="✅", fg="#27ae60")
            else:
                lbl.config(text="❌", fg="#e74c3c")
                all_ok = False

        self._run_btn.config(state="normal" if all_ok else "disabled")

    # ── SPC 로그인 다이얼로그 ─────────────────────────────────────────

    def _open_spc_dialog(self):
        dlg = Toplevel(self.root)
        dlg.title("SPC 로그인 정보")
        dlg.geometry("320x180")
        dlg.resizable(False, False)
        dlg.grab_set()

        env = _read_env()

        Label(dlg, text="SPC 아이디:", font=("맑은 고딕", 10)).place(x=20, y=30)
        id_var = StringVar(value=env.get("OLAP_ID", ""))
        Entry(dlg, textvariable=id_var, width=24, font=("맑은 고딕", 10)).place(x=110, y=28)

        Label(dlg, text="비밀번호:", font=("맑은 고딕", 10)).place(x=20, y=70)
        pw_var = StringVar(value=env.get("OLAP_PW", ""))
        Entry(dlg, textvariable=pw_var, show="●", width=24, font=("맑은 고딕", 10)).place(x=110, y=68)

        def save():
            if not id_var.get().strip() or not pw_var.get().strip():
                messagebox.showwarning("입력 오류", "아이디와 비밀번호를 모두 입력하세요.", parent=dlg)
                return
            _write_env({"OLAP_ID": id_var.get().strip(), "OLAP_PW": pw_var.get().strip()})
            dlg.destroy()
            self.refresh_status()

        ttk.Button(dlg, text="저장", command=save, width=10).place(x=100, y=120)
        ttk.Button(dlg, text="취소", command=dlg.destroy, width=10).place(x=200, y=120)

    # ── Google 인증 다이얼로그 ────────────────────────────────────────

    def _open_google_dialog(self):
        dlg = Toplevel(self.root)
        dlg.title("Google 인증")
        dlg.geometry("420x240")
        dlg.resizable(False, False)
        dlg.grab_set()

        file_var = StringVar(value="credentials.json 없음" if not CREDS_PATH.exists() else "✅  credentials.json 선택됨")

        Label(dlg, text="Google 인증", font=("맑은 고딕", 12, "bold")).pack(pady=(16, 4))
        Label(
            dlg,
            text="credentials.json 파일을 선택하세요.\n(mkt_edm_rpa 폴더에 있는 파일을 사용할 수 있습니다.)",
            font=("맑은 고딕", 9), fg="#555555", justify="center",
        ).pack()

        file_lbl = Label(dlg, textvariable=file_var, font=("맑은 고딕", 9), fg="#1a6bb5")
        file_lbl.pack(pady=4)

        def browse():
            path = filedialog.askopenfilename(
                parent=dlg,
                title="credentials.json 선택",
                filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            )
            if path:
                shutil.copy(path, CREDS_PATH)
                file_var.set("✅  credentials.json 선택됨")

        btn_frame = Frame(dlg)
        btn_frame.pack(pady=6)
        ttk.Button(btn_frame, text="파일 선택...", command=browse, width=14).pack(side="left", padx=6)

        status_lbl = Label(dlg, text="", font=("맑은 고딕", 10))
        status_lbl.pack()

        def do_auth():
            if not CREDS_PATH.exists():
                messagebox.showwarning("파일 없음", "먼저 credentials.json 파일을 선택하세요.", parent=dlg)
                return
            status_lbl.config(text="⏳  브라우저에서 Google 로그인 중...", fg="#e67e22")
            dlg.update()

            def auth_thread():
                try:
                    sys.path.insert(0, str(ROOT))
                    from src.sheets_writer import get_client
                    get_client(CREDS_PATH, TOKEN_PATH)
                    dlg.after(0, lambda: status_lbl.config(text="✅  인증 완료!", fg="#27ae60"))
                    dlg.after(500, lambda: [dlg.destroy(), self.refresh_status()])
                except Exception as exc:
                    dlg.after(0, lambda: status_lbl.config(text=f"❌  오류: {exc}", fg="#e74c3c"))

            threading.Thread(target=auth_thread, daemon=True).start()

        ttk.Button(dlg, text="Google 로그인 시작", command=do_auth, width=20).pack(pady=4)
        ttk.Button(dlg, text="취소", command=dlg.destroy, width=10).pack()

    # ── 이메일 다이얼로그 ─────────────────────────────────────────────

    def _open_email_dialog(self):
        dlg = Toplevel(self.root)
        dlg.title("이메일 알림 설정")
        dlg.geometry("400x180")
        dlg.resizable(False, False)
        dlg.grab_set()

        env = _read_env()

        Label(dlg, text="발송자 Gmail:", font=("맑은 고딕", 10)).place(x=20, y=30)
        sender_var = StringVar(value=env.get("REPORT_SENDER", ""))
        Entry(dlg, textvariable=sender_var, width=28, font=("맑은 고딕", 10)).place(x=130, y=28)

        Label(dlg, text="수신자 주소:", font=("맑은 고딕", 10)).place(x=20, y=70)
        Label(dlg, text="(쉼표로 구분)", font=("맑은 고딕", 8), fg="#888").place(x=20, y=88)
        recv_var = StringVar(value=env.get("REPORT_RECIPIENTS", ""))
        Entry(dlg, textvariable=recv_var, width=28, font=("맑은 고딕", 10)).place(x=130, y=68)

        def save():
            if not sender_var.get().strip():
                messagebox.showwarning("입력 오류", "발송자 Gmail 주소를 입력하세요.", parent=dlg)
                return
            _write_env({
                "REPORT_SENDER": sender_var.get().strip(),
                "REPORT_RECIPIENTS": recv_var.get().strip(),
                "GMAIL_CREDENTIALS_PATH": "./credentials.json",
                "GMAIL_TOKEN_PATH": "./token.json",
            })
            dlg.destroy()
            self.refresh_status()

        ttk.Button(dlg, text="저장", command=save, width=10).place(x=120, y=130)
        ttk.Button(dlg, text="취소", command=dlg.destroy, width=10).place(x=230, y=130)

    # ── Task Scheduler 등록 ───────────────────────────────────────────

    def _register_scheduler(self):
        bat = str(BAT_PATH)
        tasks = [
            (TASK_DAILY,   f'"{bat}" daily',   "/sc", "DAILY",   "/st", "08:30", "/du", "0003:00"),
            (TASK_MONTHLY, f'"{bat}" monthly', "/sc", "MONTHLY", "/d",  "1",     "/st", "08:30",
             "/du", "0096:00"),
        ]

        try:
            for t in tasks:
                name = t[0]
                tr   = t[1]
                opts = list(t[2:])
                cmd = (
                    ["schtasks", "/create", "/tn", name, "/tr", tr]
                    + opts
                    + ["/ri", "30", "/f"]
                )
                subprocess.run(cmd, check=True, capture_output=True, timeout=15)

            messagebox.showinfo(
                "등록 완료",
                "자동 실행이 등록되었습니다.\n\n"
                "• 일별: 매일 08:30 (30분 간격, 3시간 반복)\n"
                "• 월별: 매월 1일 08:30 (30분 간격, 4일 반복)",
            )
            self.refresh_status()

        except subprocess.CalledProcessError as exc:
            stderr = exc.stderr.decode("cp949", errors="replace") if exc.stderr else ""
            messagebox.showerror(
                "등록 실패",
                f"Task Scheduler 등록에 실패했습니다.\n\n"
                f"관리자 권한으로 실행하거나 IT 담당자에게 문의하세요.\n\n"
                f"오류: {stderr[:200]}",
            )
        except Exception as exc:
            messagebox.showerror("오류", str(exc))

    # ── 테스트 실행 ───────────────────────────────────────────────────

    def _run_test(self):
        self._open_run_window(dry_run=True)

    def _run_now(self):
        if not messagebox.askyesno(
            "실행 확인",
            "지금 바로 어제 날짜 데이터를 업데이트하시겠습니까?\n\n"
            "Google Sheets 에 실제로 기록됩니다.",
        ):
            return
        self._open_run_window(dry_run=False)

    def _open_run_window(self, dry_run: bool):
        win = Toplevel(self.root)
        title = "테스트 실행 (쓰기 없음)" if dry_run else "지금 실행"
        win.title(title)
        win.geometry("600x400")
        win.grab_set()

        Label(win, text=title, font=("맑은 고딕", 12, "bold")).pack(pady=(12, 4))

        frame = Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=6)

        sb = Scrollbar(frame)
        sb.pack(side="right", fill="y")

        log_box = Text(
            frame, yscrollcommand=sb.set, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", wrap="word",
        )
        log_box.pack(fill="both", expand=True)
        sb.config(command=log_box.yview)

        close_btn = ttk.Button(win, text="닫기", command=win.destroy, state="disabled")
        close_btn.pack(pady=8)

        def append(text: str):
            log_box.config(state="normal")
            log_box.insert("end", text)
            log_box.see("end")
            log_box.config(state="disabled")

        def run():
            python = str(ROOT / ".venv" / "Scripts" / "python.exe")
            if not Path(python).exists():
                python = sys.executable

            env = os.environ.copy()
            if dry_run:
                env["DRY_RUN"] = "1"

            try:
                proc = subprocess.Popen(
                    [python, "-m", "src.main", "daily"],
                    cwd=str(ROOT),
                    env=env,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                for line in proc.stdout:
                    win.after(0, append, line)
                proc.wait()
                msg = "\n✅ 완료\n" if proc.returncode == 0 else f"\n❌ 오류 (exit {proc.returncode})\n"
                win.after(0, append, msg)
            except Exception as exc:
                win.after(0, append, f"\n오류: {exc}\n")
            finally:
                win.after(0, lambda: close_btn.config(state="normal"))

        threading.Thread(target=run, daemon=True).start()


# ── 진입점 ────────────────────────────────────────────────────────────

def main():
    root = Tk()
    app = SetupApp(root)  # noqa: F841
    root.mainloop()


if __name__ == "__main__":
    main()
