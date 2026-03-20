#!/usr/bin/env python3.14
# -*- coding: utf-8 -*-
"""
AutoExec.pyw - PC 관리 / WOL 부팅 / 자동실행 스케줄러
Python 3.14, tkinter + SQLite3 + JSON(로컬 UI 설정)
"""

import os
import sys
import json
import socket
import struct
import sqlite3
import subprocess
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date

from dotenv import load_dotenv

# ─── 경로 설정 ────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(SCRIPT_DIR, ".env")
JSON_PATH = os.path.join(SCRIPT_DIR, "AutoExec.json")
AUTOEXEC_DB = os.path.join(SCRIPT_DIR, "AutoExec.db")

load_dotenv(ENV_PATH)


def _to_hm(val):
    """timedelta / str → 'HH:MM' 변환 (pymysql TIME 컬럼 대응)"""
    from datetime import timedelta
    if isinstance(val, timedelta):
        total_sec = int(val.total_seconds())
        h, m = divmod(total_sec // 60, 60)
        return f"{h:02d}:{m:02d}"
    s = str(val).strip()
    # "H:MM:SS" or "HH:MM:SS" → "HH:MM"
    parts = s.split(":")
    if len(parts) >= 2:
        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    return s.zfill(5)


# ═══════════════════════════════════════════════════════════
#  SQLite3 헬퍼
# ═══════════════════════════════════════════════════════════
def _dict_factory(cursor, row):
    """sqlite3 Row → dict 변환 (pymysql DictCursor 호환)"""
    return {col[0]: row[idx] for idx, col in enumerate(cursor.description)}


def get_db_connection():
    """SQLite3 연결 생성"""
    conn = sqlite3.connect(AUTOEXEC_DB)
    conn.row_factory = _dict_factory
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def db_init():
    """테이블 생성 (최초 실행 시)"""
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pcs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL DEFAULT '',
            ip TEXT NOT NULL DEFAULT '',
            mac TEXT NOT NULL DEFAULT '',
            auto_boot INTEGER NOT NULL DEFAULT 0,
            boot_start TEXT NOT NULL DEFAULT '00:00',
            boot_end TEXT NOT NULL DEFAULT '00:00',
            skip_holiday INTEGER NOT NULL DEFAULT 1,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL DEFAULT '',
            enabled INTEGER NOT NULL DEFAULT 1,
            run_time TEXT NOT NULL DEFAULT '00:00',
            executable TEXT NOT NULL DEFAULT '',
            arguments TEXT NOT NULL DEFAULT '',
            python_venv TEXT NOT NULL DEFAULT '',
            skip_holiday INTEGER NOT NULL DEFAULT 1,
            repeat_mode TEXT NOT NULL DEFAULT 'once',
            repeat_interval INTEGER NOT NULL DEFAULT 0,
            repeat_end_time TEXT DEFAULT '23:59',
            last_run TEXT NOT NULL DEFAULT '',
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS closed_days (
            date_int INTEGER PRIMARY KEY,
            reason TEXT NOT NULL DEFAULT ''
        )
    """)
    conn.commit()
    conn.close()


def db_fetch_pcs():
    """PC 목록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM pcs ORDER BY sort_order, id")
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_pc(pc_id, name, ip, mac, auto_boot, boot_start, boot_end, skip_holiday):
    """PC 추가 또는 수정"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if pc_id:
            cur.execute(
                "UPDATE pcs SET name=?, ip=?, mac=?, auto_boot=?, boot_start=?, boot_end=?, skip_holiday=? WHERE id=?",
                (name, ip, mac, int(auto_boot), boot_start, boot_end, int(skip_holiday), pc_id),
            )
        else:
            cur.execute(
                "INSERT INTO pcs (name, ip, mac, auto_boot, boot_start, boot_end, skip_holiday, sort_order) "
                "VALUES (?,?,?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM pcs))",
                (name, ip, mac, int(auto_boot), boot_start, boot_end, int(skip_holiday)),
            )
        conn.commit()
    finally:
        conn.close()


def db_delete_pc(pc_id):
    """PC 삭제"""
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM pcs WHERE id=?", (pc_id,))
        conn.commit()
    finally:
        conn.close()


def db_swap_sort_order(table, id_a, id_b):
    """두 레코드의 sort_order를 교환"""
    if table not in ("pcs", "tasks"):
        return
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(f"SELECT id, sort_order FROM [{table}] WHERE id IN (?, ?)", (id_a, id_b))
        rows = {r["id"]: r["sort_order"] for r in cur.fetchall()}
        if len(rows) == 2:
            cur.execute(f"UPDATE [{table}] SET sort_order=? WHERE id=?", (rows[id_b], id_a))
            cur.execute(f"UPDATE [{table}] SET sort_order=? WHERE id=?", (rows[id_a], id_b))
        conn.commit()
    finally:
        conn.close()


def db_fetch_tasks():
    """자동실행 목록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM tasks ORDER BY sort_order, id")
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_task(task_id, name, enabled, run_time, executable, arguments, python_venv, skip_holiday,
                   repeat_mode="once", repeat_interval=0, repeat_end_time="23:59"):
    """자동실행 추가 또는 수정"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if task_id:
            cur.execute(
                "UPDATE tasks SET name=?, enabled=?, run_time=?, executable=?, "
                "arguments=?, python_venv=?, skip_holiday=?, "
                "repeat_mode=?, repeat_interval=?, repeat_end_time=? WHERE id=?",
                (name, int(enabled), run_time, executable, arguments, python_venv, int(skip_holiday),
                 repeat_mode, repeat_interval, repeat_end_time, task_id),
            )
        else:
            cur.execute(
                "INSERT INTO tasks (name, enabled, run_time, executable, arguments, python_venv, skip_holiday, "
                "repeat_mode, repeat_interval, repeat_end_time, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM tasks))",
                (name, int(enabled), run_time, executable, arguments, python_venv, int(skip_holiday),
                 repeat_mode, repeat_interval, repeat_end_time),
            )
        conn.commit()
    finally:
        conn.close()


def db_delete_task(task_id):
    """자동실행 삭제"""
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))
        conn.commit()
    finally:
        conn.close()


def db_update_task_last_run(task_id, date_str):
    """자동실행 last_run 갱신"""
    conn = get_db_connection()
    try:
        conn.execute("UPDATE tasks SET last_run=? WHERE id=?", (date_str, task_id))
        conn.commit()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════
#  JSON 로컬 설정 (윈도우 좌표, 최상위)
# ═══════════════════════════════════════════════════════════
def load_local_settings():
    """로컬 UI 설정 로드"""
    defaults = {"window": {"x": 200, "y": 200, "width": 620, "height": 600, "topmost": False}}
    if os.path.exists(JSON_PATH):
        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in defaults.items():
                if k not in data:
                    data[k] = v
            return data
        except Exception:
            pass
    return defaults


def save_local_settings(settings):
    """로컬 UI 설정 저장"""
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════
#  휴장일 (AutoExec.db closed_days 테이블)
# ═══════════════════════════════════════════════════════════
def load_closed_days():
    """휴장일 목록 로드 (YYYYMMDD 정수 set)"""
    closed = set()
    if not os.path.exists(AUTOEXEC_DB):
        return closed
    try:
        conn = sqlite3.connect(AUTOEXEC_DB)
        cur = conn.cursor()
        cur.execute("SELECT date_int FROM closed_days")
        for row in cur.fetchall():
            closed.add(row[0])
        conn.close()
    except Exception:
        pass
    return closed


# ═══════════════════════════════════════════════════════════
#  텔레그램
# ═══════════════════════════════════════════════════════════
def send_telegram(message):
    """텔레그램 메시지 전송"""
    token = os.getenv("TELEGRAM_TOKEN", "")
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "")
    if not token or not chat_id:
        return False
    try:
        import urllib.request
        import urllib.parse
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        data = urllib.parse.urlencode({"chat_id": chat_id, "text": message}).encode()
        req = urllib.request.Request(url, data=data, method="POST")
        urllib.request.urlopen(req, timeout=10)
        return True
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════
#  WOL (Wake-on-LAN)
# ═══════════════════════════════════════════════════════════
def send_wol(mac_str):
    """매직 패킷 전송"""
    mac_str = mac_str.replace("-", "").replace(":", "")
    if len(mac_str) != 12:
        return False
    mac_bytes = bytes.fromhex(mac_str)
    magic = b"\xff" * 6 + mac_bytes * 16
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        s.sendto(magic, ("255.255.255.255", 9))
    return True


def ping_host(ip, timeout=2):
    """ping 으로 호스트 응답 확인"""
    try:
        result = subprocess.run(
            ["ping", "-n", "1", "-w", str(timeout * 1000), ip],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return result.returncode == 0
    except Exception:
        return False


def _ping_wait(ip, timeout_sec, log_callback):
    """timeout_sec 동안 5초 간격으로 ping 체크. 성공 시 True 반환."""
    elapsed = 0
    while elapsed < timeout_sec:
        if ping_host(ip, timeout=2):
            return True
        elapsed += 5
        if elapsed < timeout_sec:
            log_callback(f"[WOL] {ip} 응답 대기중... ({elapsed}/{timeout_sec}초)")
            time.sleep(5)
    return False


def wol_boot_thread(mac, ip, pc_name, log_callback):
    """WOL 부팅 스레드
    1차: WOL 전송 → 90초 ping 대기
    2차: 실패 시 WOL 재전송 → 90초 ping 대기
    최종 실패 시 텔레그램 알림
    """
    for attempt in range(1, 3):
        send_wol(mac)
        log_callback(f"[WOL] {pc_name}({ip}) 매직 패킷 전송 ({attempt}차)")
        time.sleep(3)  # WOL 수신 대기

        log_callback(f"[WOL] {pc_name}({ip}) 부팅 응답 대기중... (최대 90초)")
        if _ping_wait(ip, 90, log_callback):
            log_callback(f"[WOL] {pc_name}({ip}) 부팅 완료 ({attempt}차 시도)")
            return True

        if attempt == 1:
            log_callback(f"[WOL] {pc_name}({ip}) 1차 실패, 재시도...")

    # 2차까지 실패 → 텔레그램 알림
    msg = f"[AutoExec] {pc_name}({ip}) WOL 부팅 실패 (2회 시도, 총 180초 경과)"
    log_callback(f"[WOL] {msg}")
    send_telegram(msg)
    return False


# ═══════════════════════════════════════════════════════════
#  PC 편집 다이얼로그
# ═══════════════════════════════════════════════════════════
class PCEditDialog(tk.Toplevel):
    def __init__(self, parent, pc=None):
        super().__init__(parent)
        self.result = None
        self.pc = pc
        self.title("PC 편집" if pc else "PC 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        labels = ["PC 이름:", "IP 주소:", "MAC 주소:"]
        self.entries = {}
        for i, lbl in enumerate(labels):
            ttk.Label(frame, text=lbl).grid(row=i, column=0, sticky=tk.W, pady=3)
            ent = ttk.Entry(frame, width=30)
            ent.grid(row=i, column=1, columnspan=2, pady=3, padx=(5, 0))
            self.entries[i] = ent

        # 자동부팅 체크
        self.var_auto_boot = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="자동 부팅", variable=self.var_auto_boot).grid(
            row=3, column=0, columnspan=3, sticky=tk.W, pady=5
        )

        ttk.Label(frame, text="시작 시간:").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.ent_start = ttk.Entry(frame, width=8)
        self.ent_start.grid(row=4, column=1, sticky=tk.W, padx=(5, 0))
        self.ent_start.insert(0, "00:00")

        ttk.Label(frame, text="종료 시간:").grid(row=5, column=0, sticky=tk.W, pady=3)
        self.ent_end = ttk.Entry(frame, width=8)
        self.ent_end.grid(row=5, column=1, sticky=tk.W, padx=(5, 0))
        self.ent_end.insert(0, "00:00")

        # 휴장일 제외
        self.var_skip_holiday = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="휴장일 제외", variable=self.var_skip_holiday).grid(
            row=6, column=0, columnspan=3, sticky=tk.W, pady=3
        )

        # 기존 데이터 채우기
        if pc:
            self.entries[0].insert(0, pc["name"])
            self.entries[1].insert(0, pc["ip"])
            self.entries[2].insert(0, pc["mac"])
            self.var_auto_boot.set(bool(pc["auto_boot"]))
            self.ent_start.delete(0, tk.END)
            self.ent_start.insert(0, _to_hm(pc["boot_start"]))
            self.ent_end.delete(0, tk.END)
            self.ent_end.insert(0, _to_hm(pc["boot_end"]))
            self.var_skip_holiday.set(bool(pc.get("skip_holiday", 1)))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=7, column=0, columnspan=3, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.transient(parent)
        self.update_idletasks()
        # 부모 윈도우 중앙에 배치
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_width()
        dh = self.winfo_height()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.entries[0].focus_set()

    def _on_ok(self):
        name = self.entries[0].get().strip()
        ip = self.entries[1].get().strip()
        mac = self.entries[2].get().strip()
        if not name or not ip or not mac:
            messagebox.showwarning("입력 오류", "모든 필드를 입력하세요.", parent=self)
            return
        self.result = {
            "id": self.pc["id"] if self.pc else None,
            "name": name,
            "ip": ip,
            "mac": mac,
            "auto_boot": self.var_auto_boot.get(),
            "boot_start": self.ent_start.get().strip(),
            "boot_end": self.ent_end.get().strip(),
            "skip_holiday": self.var_skip_holiday.get(),
        }
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  자동실행 편집 다이얼로그
# ═══════════════════════════════════════════════════════════
class TaskEditDialog(tk.Toplevel):
    _REPEAT_MODES = {
        "매일 1회": "once", "N분 간격": "minutes", "N시간 간격": "hours",
        "요일 지정": "weekly", "매월 지정일": "monthly",
    }
    _REPEAT_LABELS = {
        "once": "매일 1회", "minutes": "N분 간격", "hours": "N시간 간격",
        "weekly": "요일 지정", "monthly": "매월 지정일",
    }
    _WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"]

    def __init__(self, parent, task=None):
        super().__init__(parent)
        self.result = None
        self.task = task
        self.title("자동실행 편집" if task else "자동실행 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # 작업 이름
        ttk.Label(frame, text="작업 이름:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_name = ttk.Entry(frame, width=35)
        self.ent_name.grid(row=0, column=1, columnspan=2, pady=3, padx=(5, 0))

        # 사용 여부
        self.var_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="사용", variable=self.var_enabled).grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=3
        )

        # 실행 모드
        ttk.Label(frame, text="실행 모드:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.var_repeat_mode = tk.StringVar(value="매일 1회")
        self.cmb_mode = ttk.Combobox(frame, textvariable=self.var_repeat_mode,
                                     values=list(self._REPEAT_MODES.keys()),
                                     state="readonly", width=12)
        self.cmb_mode.grid(row=2, column=1, sticky=tk.W, padx=(5, 0), pady=3)
        self.cmb_mode.bind("<<ComboboxSelected>>", self._on_mode_changed)

        # 반복 간격 (N분/N시간 모드에서만 표시)
        self.lbl_interval = ttk.Label(frame, text="반복 간격:")
        self.lbl_interval.grid(row=3, column=0, sticky=tk.W, pady=3)
        interval_frame = ttk.Frame(frame)
        interval_frame.grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=(5, 0), pady=3)
        self.var_interval = tk.IntVar(value=30)
        self.spn_interval = ttk.Spinbox(interval_frame, from_=1, to=1440, width=6,
                                        textvariable=self.var_interval)
        self.spn_interval.pack(side=tk.LEFT)
        self.lbl_interval_unit = ttk.Label(interval_frame, text="분")
        self.lbl_interval_unit.pack(side=tk.LEFT, padx=(3, 0))

        # 요일 선택 (weekly 모드에서만 표시)
        self.lbl_weekdays = ttk.Label(frame, text="실행 요일:")
        self.lbl_weekdays.grid(row=3, column=0, sticky=tk.W, pady=3)
        self.weekday_frame = ttk.Frame(frame)
        self.weekday_frame.grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=(5, 0), pady=3)
        self.var_weekdays = []
        for i, name in enumerate(self._WEEKDAY_NAMES):
            var = tk.BooleanVar(value=False)
            self.var_weekdays.append(var)
            ttk.Checkbutton(self.weekday_frame, text=name, variable=var).pack(side=tk.LEFT, padx=(0, 3))

        # 매월 날짜 선택 (monthly 모드에서만 표시)
        self.lbl_monthday = ttk.Label(frame, text="실행 날짜:")
        self.lbl_monthday.grid(row=3, column=0, sticky=tk.W, pady=3)
        monthday_frame = ttk.Frame(frame)
        monthday_frame.grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=(5, 0), pady=3)
        self.var_monthday = tk.IntVar(value=1)
        self.spn_monthday = ttk.Spinbox(monthday_frame, from_=1, to=31, width=4,
                                        textvariable=self.var_monthday)
        self.spn_monthday.pack(side=tk.LEFT)
        self.lbl_monthday_unit = ttk.Label(monthday_frame, text="일")
        self.lbl_monthday_unit.pack(side=tk.LEFT, padx=(3, 0))
        self.monthday_frame = monthday_frame

        # 실행 시간 (once: 실행 시간, repeat: 시작 시간)
        self.lbl_time = ttk.Label(frame, text="실행 시간:")
        self.lbl_time.grid(row=4, column=0, sticky=tk.W, pady=3)
        self.ent_time = ttk.Entry(frame, width=8)
        self.ent_time.grid(row=4, column=1, sticky=tk.W, padx=(5, 0))
        self.ent_time.insert(0, "00:00")

        # 종료 시간 (반복 모드에서만 표시)
        self.lbl_end_time = ttk.Label(frame, text="종료 시간:")
        self.lbl_end_time.grid(row=5, column=0, sticky=tk.W, pady=3)
        self.ent_end_time = ttk.Entry(frame, width=8)
        self.ent_end_time.grid(row=5, column=1, sticky=tk.W, padx=(5, 0))
        self.ent_end_time.insert(0, "23:59")

        # 실행 파일
        ttk.Label(frame, text="실행 파일:").grid(row=6, column=0, sticky=tk.W, pady=3)
        self.ent_exe = ttk.Entry(frame, width=35)
        self.ent_exe.grid(row=6, column=1, pady=3, padx=(5, 0))
        ttk.Button(frame, text="..", width=3, command=self._browse_exe).grid(row=6, column=2, padx=2)

        # 파라미터
        ttk.Label(frame, text="파라미터:").grid(row=7, column=0, sticky=tk.W, pady=3)
        self.ent_args = ttk.Entry(frame, width=35)
        self.ent_args.grid(row=7, column=1, columnspan=2, pady=3, padx=(5, 0))

        # Python 가상환경 경로
        ttk.Label(frame, text="Python 경로:").grid(row=8, column=0, sticky=tk.W, pady=3)
        self.ent_venv = ttk.Entry(frame, width=35)
        self.ent_venv.grid(row=8, column=1, pady=3, padx=(5, 0))
        ttk.Button(frame, text="..", width=3, command=self._browse_venv).grid(row=8, column=2, padx=2)
        ttk.Label(frame, text="(.py 파일인 경우 python.exe 경로)", foreground="gray").grid(
            row=9, column=0, columnspan=3, sticky=tk.W
        )

        # 휴장일 제외
        self.var_skip_holiday = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="휴장일 제외", variable=self.var_skip_holiday).grid(
            row=10, column=0, columnspan=3, sticky=tk.W, pady=3
        )

        # 기존 데이터 채우기
        if task:
            self.ent_name.insert(0, task["name"])
            self.var_enabled.set(bool(task["enabled"]))
            self.ent_time.delete(0, tk.END)
            self.ent_time.insert(0, _to_hm(task["run_time"]))
            self.ent_exe.insert(0, task["executable"])
            self.ent_args.insert(0, task.get("arguments", ""))
            self.ent_venv.insert(0, task["python_venv"])
            self.var_skip_holiday.set(bool(task["skip_holiday"]))
            # 반복 모드 복원
            mode = task.get("repeat_mode", "once")
            self.var_repeat_mode.set(self._REPEAT_LABELS.get(mode, "매일 1회"))
            interval_val = task.get("repeat_interval", 0) or 0
            if mode == "weekly":
                # 비트마스크에서 요일 체크박스 복원
                for i in range(7):
                    self.var_weekdays[i].set(bool(interval_val & (1 << i)))
            elif mode == "monthly":
                self.var_monthday.set(interval_val if 1 <= interval_val <= 31 else 1)
            else:
                self.var_interval.set(interval_val if interval_val > 0 else 30)
            self.ent_end_time.delete(0, tk.END)
            self.ent_end_time.insert(0, _to_hm(task.get("repeat_end_time", "23:59") or "23:59"))

        self._on_mode_changed()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=11, column=0, columnspan=3, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.transient(parent)
        self.update_idletasks()
        # 부모 윈도우 중앙에 배치
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_width()
        dh = self.winfo_height()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.ent_name.focus_set()

    def _on_mode_changed(self, event=None):
        """실행 모드 변경 시 UI 요소 표시/숨김"""
        mode_label = self.var_repeat_mode.get()
        mode = self._REPEAT_MODES.get(mode_label, "once")

        # 모든 row=3 위젯 숨김
        self.lbl_interval.grid_remove()
        self.spn_interval.master.grid_remove()
        self.lbl_weekdays.grid_remove()
        self.weekday_frame.grid_remove()
        self.lbl_monthday.grid_remove()
        self.monthday_frame.grid_remove()
        # 종료 시간 숨김
        self.lbl_end_time.grid_remove()
        self.ent_end_time.grid_remove()

        if mode in ("minutes", "hours"):
            self.lbl_interval.grid()
            self.spn_interval.master.grid()
            self.lbl_end_time.grid()
            self.ent_end_time.grid()
            self.lbl_time.config(text="시작 시간:")
            self.lbl_interval_unit.config(text="분" if mode == "minutes" else "시간")
        elif mode == "weekly":
            self.lbl_weekdays.grid()
            self.weekday_frame.grid()
            self.lbl_time.config(text="실행 시간:")
        elif mode == "monthly":
            self.lbl_monthday.grid()
            self.monthday_frame.grid()
            self.lbl_time.config(text="실행 시간:")
        else:
            self.lbl_time.config(text="실행 시간:")

    def _browse_exe(self):
        path = filedialog.askopenfilename(
            title="실행 파일 선택",
            filetypes=[("실행파일", "*.exe *.py *.pyw *.bat"), ("모든파일", "*.*")],
            parent=self,
        )
        if path:
            self.ent_exe.delete(0, tk.END)
            self.ent_exe.insert(0, path)

    def _browse_venv(self):
        path = filedialog.askopenfilename(
            title="Python 실행파일 선택",
            filetypes=[("Python", "python.exe pythonw.exe"), ("모든파일", "*.*")],
            parent=self,
        )
        if path:
            self.ent_venv.delete(0, tk.END)
            self.ent_venv.insert(0, path)

    def _on_ok(self):
        name = self.ent_name.get().strip()
        executable = self.ent_exe.get().strip()
        if not name or not executable:
            messagebox.showwarning("입력 오류", "작업 이름과 실행 파일을 입력하세요.", parent=self)
            return
        mode_label = self.var_repeat_mode.get()
        repeat_mode = self._REPEAT_MODES.get(mode_label, "once")

        # 모드별 interval 값 결정
        if repeat_mode == "weekly":
            # 비트마스크: bit0=월, bit1=화, ..., bit6=일
            bitmask = 0
            for i, var in enumerate(self.var_weekdays):
                if var.get():
                    bitmask |= (1 << i)
            if bitmask == 0:
                messagebox.showwarning("입력 오류", "실행할 요일을 하나 이상 선택하세요.", parent=self)
                return
            repeat_interval = bitmask
        elif repeat_mode == "monthly":
            repeat_interval = self.var_monthday.get()
        elif repeat_mode in ("minutes", "hours"):
            repeat_interval = self.var_interval.get()
        else:
            repeat_interval = 0

        self.result = {
            "id": self.task["id"] if self.task else None,
            "name": name,
            "enabled": self.var_enabled.get(),
            "run_time": self.ent_time.get().strip(),
            "executable": executable,
            "arguments": self.ent_args.get().strip(),
            "python_venv": self.ent_venv.get().strip(),
            "skip_holiday": self.var_skip_holiday.get(),
            "repeat_mode": repeat_mode,
            "repeat_interval": repeat_interval,
            "repeat_end_time": self.ent_end_time.get().strip() if repeat_mode in ("minutes", "hours") else "23:59",
        }
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  메인 윈도우
# ═══════════════════════════════════════════════════════════
class AutoExecApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("실행 관리 서버")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.settings = load_local_settings()
        self.closed_days = load_closed_days()
        self.today_str = ""
        self.booted_today = {}  # pc_id -> bool (당일 자동부팅 완료)
        self.tray_icon = None
        self.tray_thread = None
        self.hidden = False
        self._running_tasks = set()  # 실행중인 task ID (중복 실행 방지)
        self._task_processes = {}   # task_id → subprocess.Popen (프로세스 강제 종료용)
        self._booting_pcs = set()  # WOL 부팅중인 pc ID (중복 방지)

        self._build_ui()
        self._restore_window()
        self._refresh_pc_list()
        self._refresh_task_list()
        self._tick()
        # 트레이 아이콘은 mainloop 진입 후 생성 (윈도우 준비 완료 후)
        self.root.after(500, self._setup_tray)

    # ─── UI 구성 ──────────────────────────────────────────
    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)

        # ── row 0: 최상위 + GitHub 다운로드 (한 줄) ──
        top_frame = ttk.Frame(root)
        top_frame.grid(row=0, column=0, sticky=tk.EW, padx=8, pady=(6, 2))
        top_frame.columnconfigure(1, weight=1)

        self.var_topmost = tk.BooleanVar(value=self.settings["window"].get("topmost", False))
        ttk.Checkbutton(top_frame, text="최상위", variable=self.var_topmost,
                        command=self._toggle_topmost).grid(row=0, column=0, sticky=tk.W)
        self._apply_topmost()

        self.git_url_var = tk.StringVar()
        self.git_url_entry = ttk.Entry(top_frame, textvariable=self.git_url_var, font=("Consolas", 9))
        self.git_url_entry.grid(row=0, column=1, sticky=tk.EW, padx=(10, 3))
        self.git_url_entry.bind("<Return>", lambda e: self._git_download())
        self.git_url_entry.bind("<<Paste>>", self._on_git_paste)

        self.git_dl_btn = ttk.Button(top_frame, text="Git", width=4, command=self._git_download)
        self.git_dl_btn.grid(row=0, column=2)

        # ── row 1: 자동실행 (전체 폭) ──
        lf_task = ttk.LabelFrame(root, text="자동실행", padding=5)
        lf_task.grid(row=1, column=0, sticky=tk.NSEW, padx=8, pady=2)
        lf_task.columnconfigure(0, weight=1)
        lf_task.rowconfigure(0, weight=1)

        task_frame = ttk.Frame(lf_task)
        task_frame.grid(row=0, column=0, sticky=tk.NSEW)
        task_frame.columnconfigure(0, weight=1)
        task_frame.rowconfigure(0, weight=1)

        cols = ("사용", "이름", "시간", "실행파일")
        self.task_tree = ttk.Treeview(task_frame, columns=cols, show="headings", height=12)
        self.task_tree.heading("사용", text="사용")
        self.task_tree.heading("이름", text="이름")
        self.task_tree.heading("시간", text="시간")
        self.task_tree.heading("실행파일", text="실행파일")
        self.task_tree.column("사용", width=40, anchor=tk.CENTER)
        self.task_tree.column("이름", width=130)
        self.task_tree.column("시간", width=110, anchor=tk.CENTER)
        self.task_tree.column("실행파일", width=300)
        self.task_tree.grid(row=0, column=0, sticky=tk.NSEW)
        task_scroll = ttk.Scrollbar(task_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        task_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.task_tree.config(yscrollcommand=task_scroll.set)
        self.task_tree.bind("<Double-1>", self._on_task_double_click)

        task_btn_frame = ttk.Frame(lf_task)
        task_btn_frame.grid(row=1, column=0, sticky=tk.W, pady=(3, 0))
        ttk.Button(task_btn_frame, text="추가", width=6, command=self._add_task).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="편집", width=6, command=self._edit_task).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="삭제", width=6, command=self._delete_task).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="실행", width=6, command=self._run_task).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="중지", width=6, command=self._stop_task).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="\u25b2", width=3, command=lambda: self._move_task(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(task_btn_frame, text="\u25bc", width=3, command=lambda: self._move_task(1)).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(task_btn_frame, text="시작폴더", command=self._open_startup_folder).pack(side=tk.LEFT)

        # ── row 2: 로그 (좌) + PC 관리 (우) ──
        bottom_frame = ttk.Frame(root)
        bottom_frame.grid(row=2, column=0, sticky=tk.NSEW, padx=8, pady=(2, 8))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.rowconfigure(0, weight=1)

        # ── 로그 (좌측, 확장) ──
        lf_log = ttk.LabelFrame(bottom_frame, text="로그", padding=5)
        lf_log.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 4))
        lf_log.columnconfigure(0, weight=1)
        lf_log.rowconfigure(0, weight=1)

        self.log_text = tk.Text(lf_log, height=8, font=("Consolas", 9), state=tk.DISABLED, wrap=tk.NONE)
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        log_scroll = ttk.Scrollbar(lf_log, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.log_text.config(yscrollcommand=log_scroll.set)

        log_btn_frame = ttk.Frame(lf_log)
        log_btn_frame.grid(row=1, column=0, sticky=tk.W, pady=(3, 0))
        ttk.Button(log_btn_frame, text="로그 삭제", command=self._clear_log).pack(side=tk.LEFT)

        # ── PC 관리 (우측, 고정 폭) ──
        lf_pc = ttk.LabelFrame(bottom_frame, text="PC 관리", padding=5)
        lf_pc.grid(row=0, column=1, sticky=tk.NSEW)
        lf_pc.rowconfigure(0, weight=1)

        self.pc_listbox = tk.Listbox(lf_pc, height=6, width=14, font=("Consolas", 10))
        self.pc_listbox.grid(row=0, column=0, sticky=tk.NSEW)
        pc_scroll = ttk.Scrollbar(lf_pc, orient=tk.VERTICAL, command=self.pc_listbox.yview)
        pc_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.pc_listbox.config(yscrollcommand=pc_scroll.set)

        pc_side_frame = ttk.Frame(lf_pc)
        pc_side_frame.grid(row=0, column=2, sticky=tk.N, padx=(3, 0))
        ttk.Button(pc_side_frame, text="편집", width=5, command=self._edit_pc).pack(pady=1)
        ttk.Button(pc_side_frame, text="부팅", width=5, command=self._boot_pc).pack(pady=1)
        ttk.Button(pc_side_frame, text="핑", width=5, command=self._ping_pc).pack(pady=1)

        pc_btn_frame = ttk.Frame(lf_pc)
        pc_btn_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(3, 0))
        ttk.Button(pc_btn_frame, text="추가", width=5, command=self._add_pc).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(pc_btn_frame, text="삭제", width=5, command=self._delete_pc).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(pc_btn_frame, text="\u25b2", width=2, command=lambda: self._move_pc(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(pc_btn_frame, text="\u25bc", width=2, command=lambda: self._move_pc(1)).pack(side=tk.LEFT)

        # 행 가중치: row 1(자동실행) 고정, row 2(로그+PC) 확장
        root.rowconfigure(0, weight=0)
        root.rowconfigure(1, weight=0)
        root.rowconfigure(2, weight=1)

    # ─── 윈도우 좌표 ─────────────────────────────────────
    def _restore_window(self):
        w = self.settings["window"]
        self.root.geometry(f"{w['width']}x{w['height']}+{w['x']}+{w['y']}")

    def _save_window(self):
        try:
            geo = self.root.geometry()
            # WxH+X+Y
            size, pos = geo.split("+", 1)
            width, height = size.split("x")
            x, y = pos.split("+")
            self.settings["window"].update({"x": int(x), "y": int(y), "width": int(width), "height": int(height)})
            save_local_settings(self.settings)
        except Exception:
            pass

    # ─── 최상위 ──────────────────────────────────────────
    def _toggle_topmost(self):
        self.settings["window"]["topmost"] = self.var_topmost.get()
        self._apply_topmost()
        save_local_settings(self.settings)

    def _apply_topmost(self):
        self.root.attributes("-topmost", self.var_topmost.get())

    # ─── 로그 ────────────────────────────────────────────
    def _clear_log(self):
        """로그 텍스트 전체 삭제"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def log(self, msg):
        """로그 메시지 추가 (스레드 안전)"""
        def _append():
            now_str = datetime.now().strftime("%H:%M:%S")
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"[{now_str}] {msg}\n")
            self.log_text.see(tk.END)
            # 최대 500줄 유지
            lines = int(self.log_text.index("end-1c").split(".")[0])
            if lines > 500:
                self.log_text.delete("1.0", f"{lines - 500}.0")
            self.log_text.config(state=tk.DISABLED)

        self.root.after(0, _append)

    # ─── PC 리스트 ───────────────────────────────────────
    def _refresh_pc_list(self):
        self.pc_data = db_fetch_pcs()
        self.pc_listbox.delete(0, tk.END)
        for pc in self.pc_data:
            auto_mark = "[A]" if pc["auto_boot"] else "   "
            self.pc_listbox.insert(tk.END, f"{auto_mark} {pc['name']}")

    def _get_selected_pc(self):
        sel = self.pc_listbox.curselection()
        if not sel:
            messagebox.showinfo("알림", "PC를 선택하세요.", parent=self.root)
            return None
        return self.pc_data[sel[0]]

    def _add_pc(self):
        dlg = PCEditDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_pc(None, r["name"], r["ip"], r["mac"], r["auto_boot"], r["boot_start"], r["boot_end"], r["skip_holiday"])
            self._refresh_pc_list()
            self.log(f"PC 추가: {r['name']}")

    def _edit_pc(self):
        pc = self._get_selected_pc()
        if not pc:
            return
        dlg = PCEditDialog(self.root, pc)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_pc(r["id"], r["name"], r["ip"], r["mac"], r["auto_boot"], r["boot_start"], r["boot_end"], r["skip_holiday"])
            self._refresh_pc_list()
            self.log(f"PC 수정: {r['name']}")

    def _delete_pc(self):
        pc = self._get_selected_pc()
        if not pc:
            return
        if messagebox.askyesno("삭제 확인", f"'{pc['name']}' PC를 삭제하시겠습니까?", parent=self.root):
            db_delete_pc(pc["id"])
            self._refresh_pc_list()
            self.log(f"PC 삭제: {pc['name']}")

    def _move_pc(self, direction):
        """PC 순서 이동 (direction: -1=위, 1=아래)"""
        sel = self.pc_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.pc_data):
            return
        db_swap_sort_order("pcs", self.pc_data[idx]["id"], self.pc_data[new_idx]["id"])
        self._refresh_pc_list()
        self.pc_listbox.selection_set(new_idx)
        self.pc_listbox.see(new_idx)

    def _boot_pc(self):
        pc = self._get_selected_pc()
        if not pc:
            return
        pc_id = pc["id"]
        if pc_id in self._booting_pcs:
            self.log(f"{pc['name']} 이미 부팅 진행중 - 건너뜀")
            return
        self._booting_pcs.add(pc_id)
        self.log(f"{pc['name']} 부팅 시작 (MAC: {pc['mac']}, IP: {pc['ip']})")

        def _boot_and_cleanup():
            try:
                wol_boot_thread(pc["mac"], pc["ip"], pc["name"], self.log)
            finally:
                self._booting_pcs.discard(pc_id)

        threading.Thread(target=_boot_and_cleanup, daemon=True).start()

    def _ping_pc(self):
        pc = self._get_selected_pc()
        if not pc:
            return

        def _do_ping():
            self.log(f"{pc['name']} ({pc['ip']}) 핑 전송중...")
            if ping_host(pc["ip"]):
                self.log(f"{pc['name']} ({pc['ip']}) 응답 OK")
            else:
                self.log(f"{pc['name']} ({pc['ip']}) 응답 없음")

        threading.Thread(target=_do_ping, daemon=True).start()

    # ─── 자동실행 리스트 ─────────────────────────────────
    def _refresh_task_list(self):
        self.task_data = db_fetch_tasks()
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        for task in self.task_data:
            enabled_mark = "O" if task["enabled"] else ""
            exe_display = os.path.basename(task["executable"])
            if task["python_venv"]:
                exe_display += " (venv)"
            # 시간 표시: 반복 모드일 경우 간격 정보 포함
            repeat_mode = task.get("repeat_mode", "once")
            time_display = _to_hm(task["run_time"])
            interval_val = task.get("repeat_interval", 0) or 0
            if repeat_mode == "minutes":
                time_display += f" ({interval_val}분)"
            elif repeat_mode == "hours":
                time_display += f" ({interval_val}시간)"
            elif repeat_mode == "weekly":
                days = [n for i, n in enumerate(["월","화","수","목","금","토","일"]) if interval_val & (1 << i)]
                time_display += f" ({','.join(days)})"
            elif repeat_mode == "monthly":
                time_display += f" (매월 {interval_val}일)"
            self.task_tree.insert("", tk.END, iid=str(task["id"]),
                                  values=(enabled_mark, task["name"], time_display, exe_display))

    def _get_selected_task(self):
        sel = self.task_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "자동실행 항목을 선택하세요.", parent=self.root)
            return None
        task_id = int(sel[0])
        for t in self.task_data:
            if t["id"] == task_id:
                return t
        return None

    def _add_task(self):
        dlg = TaskEditDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_task(None, r["name"], r["enabled"], r["run_time"], r["executable"], r["arguments"],
                           r["python_venv"], r["skip_holiday"], r["repeat_mode"], r["repeat_interval"], r["repeat_end_time"])
            self._refresh_task_list()
            self.log(f"자동실행 추가: {r['name']}")

    def _on_task_double_click(self, event):
        """더블클릭: 자동실행 사용 → 편집, 아닌 경우 → 실행"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        task_id = int(item)
        task = next((t for t in self.task_data if t["id"] == task_id), None)
        if not task:
            return
        if task["enabled"]:
            self._edit_task()
        else:
            self._run_task()

    def _edit_task(self):
        task = self._get_selected_task()
        if not task:
            return
        dlg = TaskEditDialog(self.root, task)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_task(r["id"], r["name"], r["enabled"], r["run_time"], r["executable"], r["arguments"],
                           r["python_venv"], r["skip_holiday"], r["repeat_mode"], r["repeat_interval"], r["repeat_end_time"])
            self._refresh_task_list()
            self.log(f"자동실행 수정: {r['name']}")

    def _delete_task(self):
        task = self._get_selected_task()
        if not task:
            return
        if messagebox.askyesno("삭제 확인", f"'{task['name']}' 작업을 삭제하시겠습니까?", parent=self.root):
            db_delete_task(task["id"])
            self._refresh_task_list()
            self.log(f"자동실행 삭제: {task['name']}")

    def _move_task(self, direction):
        """자동실행 순서 이동 (direction: -1=위, 1=아래)"""
        sel = self.task_tree.selection()
        if not sel:
            return
        task_id = int(sel[0])
        # task_data 에서 현재 인덱스 찾기
        idx = next((i for i, t in enumerate(self.task_data) if t["id"] == task_id), None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.task_data):
            return
        db_swap_sort_order("tasks", self.task_data[idx]["id"], self.task_data[new_idx]["id"])
        self._refresh_task_list()
        # 선택 복원
        new_iid = str(task_id)
        if self.task_tree.exists(new_iid):
            self.task_tree.selection_set(new_iid)
            self.task_tree.see(new_iid)

    def _open_startup_folder(self):
        """윈도우 시작프로그램 폴더 열기"""
        startup = os.path.join(os.getenv("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
        os.startfile(startup)

    # ─── GitHub 다운로드 ──────────────────────────────────
    def _on_git_paste(self, event):
        """붙여넣기 후 URL이면 자동 다운로드 시작"""
        self.root.after(50, self._git_auto_download_on_paste)

    def _git_auto_download_on_paste(self):
        """붙여넣기된 텍스트가 GitHub URL이면 자동 다운로드"""
        url = self.git_url_var.get().strip()
        if url and "github.com/" in url:
            self._git_download()

    def _git_download(self):
        """GitHub 저장소 다운로드 (gitclone.py 호출)"""
        url = self.git_url_var.get().strip()
        if not url:
            return
        # 버튼 비활성화
        self.git_dl_btn.config(state=tk.DISABLED)
        self.git_url_entry.config(state=tk.DISABLED)
        self.log(f"[GitHub] 다운로드 시작: {url}")

        def _worker():
            try:
                gitclone_path = os.path.join(SCRIPT_DIR, "gitclone.py")
                result = subprocess.run(
                    ["python", gitclone_path, url],
                    cwd=SCRIPT_DIR,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=120,
                )
                success = result.returncode == 0
                output = (result.stdout + result.stderr).strip()
                self.root.after(0, lambda: self._git_download_done(success, output))
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self._git_download_done(False, "타임아웃 (120초 초과)"))
            except Exception as e:
                self.root.after(0, lambda: self._git_download_done(False, str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _git_download_done(self, success, output):
        """다운로드 완료 콜백 (메인 스레드)"""
        self.git_dl_btn.config(state=tk.NORMAL)
        self.git_url_entry.config(state=tk.NORMAL)
        if success:
            self.log(f"[GitHub] 다운로드 성공")
            self.git_url_var.set("")
        else:
            self.log(f"[GitHub] 다운로드 실패: {output[:200]}")
            messagebox.showerror("GitHub 다운로드 실패", output[:500], parent=self.root)
        self.git_url_entry.focus_set()

    def _run_task(self):
        """선택한 자동실행 작업을 즉시 테스트 실행"""
        task = self._get_selected_task()
        if not task:
            return
        repeat_mode = task.get("repeat_mode", "once")
        if repeat_mode == "once":
            run_stamp = datetime.now().strftime("%Y-%m-%d")
        else:
            run_stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # 수동 실행 시에도 last_run 메모리 갱신 → 자동 스케줄러 중복 실행 방지
        task["last_run"] = run_stamp
        self._execute_task(task, run_stamp)

    def _stop_task(self):
        """선택한 자동실행 작업의 실행중인 프로세스를 강제 종료"""
        task = self._get_selected_task()
        if not task:
            return
        task_id = task["id"]
        proc = self._task_processes.get(task_id)
        if proc is None or proc.poll() is not None:
            self.log(f"[자동실행] {task['name']} 실행중인 프로세스 없음")
            return
        try:
            proc.terminate()
            # 3초 후에도 안 죽으면 강제 kill
            try:
                proc.wait(timeout=3)
            except subprocess.TimeoutExpired:
                proc.kill()
            self.log(f"[자동실행] {task['name']} 프로세스 강제 종료됨")
        except Exception as e:
            self.log(f"[자동실행] {task['name']} 종료 실패: {e}")

    # ─── 1초 타이머 ──────────────────────────────────────
    def _tick(self):
        now = datetime.now()
        now_str = now.strftime("%Y-%m-%d %H:%M:%S")
        today_str = now.strftime("%Y-%m-%d")
        current_hm = now.strftime("%H:%M")

        # 타이틀바 갱신
        if not self.hidden:
            self.root.title(f"실행 관리 서버 - {now_str}")

        # 날짜 변경 시 초기화
        if self.today_str != today_str:
            self.today_str = today_str
            self.booted_today.clear()
            self.closed_days = load_closed_days()
            self.log(f"날짜 갱신: {today_str}")
            # DB에서 최신 데이터 리로드
            self._refresh_pc_list()
            self._refresh_task_list()

        is_closed = self._is_closed_day(now)

        # 자동 부팅 체크 (PC별 skip_holiday 개별 판단)
        self._check_auto_boot(current_hm, is_closed)

        # 자동실행 체크
        self._check_auto_tasks(current_hm, today_str, is_closed)

        self.root.after(1000, self._tick)

    def _is_closed_day(self, dt):
        """휴장일 판단: 토/일 + closed_days.db"""
        if dt.weekday() >= 5:  # 토(5), 일(6)
            return True
        date_int = int(dt.strftime("%Y%m%d"))
        return date_int in self.closed_days

    def _check_auto_boot(self, current_hm, is_closed):
        """자동 부팅 체크"""
        for pc in self.pc_data:
            if not pc["auto_boot"]:
                continue
            if pc.get("skip_holiday", 1) and is_closed:
                continue
            if self.booted_today.get(pc["id"]):
                continue
            start = _to_hm(pc["boot_start"])
            end = _to_hm(pc["boot_end"])
            # 자정 교차 범위 지원 (예: 23:00~01:00)
            if start <= end:
                in_range = start <= current_hm < end
            else:
                in_range = current_hm >= start or current_hm < end
            if in_range:
                pc_id = pc["id"]
                self.booted_today[pc_id] = True
                if pc_id in self._booting_pcs:
                    continue
                self._booting_pcs.add(pc_id)
                self.log(f"[자동부팅] {pc['name']} WOL 전송 시작")

                def _auto_boot(mac, ip, name, pid):
                    try:
                        wol_boot_thread(mac, ip, name, self.log)
                    finally:
                        self._booting_pcs.discard(pid)

                t = threading.Thread(
                    target=_auto_boot,
                    args=(pc["mac"], pc["ip"], pc["name"], pc_id),
                    daemon=True,
                )
                t.start()

    def _check_auto_tasks(self, current_hm, today_str, is_closed):
        """자동실행 체크 (매일 1회 + 반복 모드 지원)"""
        now = datetime.now()
        for task in self.task_data:
            if not task["enabled"]:
                continue
            if task["skip_holiday"] and is_closed:
                continue

            repeat_mode = task.get("repeat_mode", "once")

            if repeat_mode in ("once", "weekly", "monthly"):
                # 하루 1회 실행 계열: once(매일), weekly(지정 요일), monthly(지정일)
                if str(task["last_run"]) == today_str:
                    continue
                if _to_hm(task["run_time"]) != current_hm:
                    continue
                if repeat_mode == "weekly":
                    # 비트마스크: bit0=월(weekday 0), ..., bit6=일(weekday 6)
                    bitmask = task.get("repeat_interval", 0) or 0
                    if not (bitmask & (1 << now.weekday())):
                        continue
                elif repeat_mode == "monthly":
                    monthday = task.get("repeat_interval", 1) or 1
                    if now.day != monthday:
                        continue
                task["last_run"] = today_str
                self._execute_task(task, today_str)
            else:
                # 반복 모드: 시작 시간 기준 interval 배수 슬롯에서 실행
                # 예: start=00:05, interval=6h → 00:05, 06:05, 12:05, 18:05
                start_hm = _to_hm(task["run_time"])
                end_hm = _to_hm(task.get("repeat_end_time", "23:59") or "23:59")
                interval_min = task.get("repeat_interval", 1) or 1
                if repeat_mode == "hours":
                    interval_min *= 60  # 시간 → 분 변환

                # 시간 범위 확인
                if start_hm <= end_hm:
                    in_range = start_hm <= current_hm <= end_hm
                else:
                    in_range = current_hm >= start_hm or current_hm <= end_hm
                if not in_range:
                    continue

                # 오늘 시작 시간으로부터 경과 분 계산
                start_parts = start_hm.split(":")
                start_total_min = int(start_parts[0]) * 60 + int(start_parts[1])
                cur_parts = current_hm.split(":")
                cur_total_min = int(cur_parts[0]) * 60 + int(cur_parts[1])
                elapsed_min = cur_total_min - start_total_min
                if elapsed_min < 0:
                    elapsed_min += 1440  # 자정 교차

                # 현재 시각이 정확한 슬롯 시점인지 확인
                if elapsed_min % interval_min != 0:
                    continue

                # 이 슬롯에서 이미 실행했는지 확인
                last_run_str = str(task.get("last_run") or "")
                if last_run_str and last_run_str != "None":
                    try:
                        last_dt = datetime.strptime(last_run_str, "%Y-%m-%d %H:%M:%S")
                        last_hm = last_dt.strftime("%H:%M")
                        if last_dt.date() == now.date() and last_hm == current_hm:
                            continue  # 같은 날 같은 슬롯에서 이미 실행함
                    except ValueError:
                        try:
                            last_date = datetime.strptime(last_run_str[:10], "%Y-%m-%d").date()
                            if last_date == now.date():
                                continue  # 날짜만 있는 경우, 오늘 실행 기록 있으면 건너뜀
                        except ValueError:
                            pass

                now_dt_str = now.strftime("%Y-%m-%d %H:%M:%S")
                task["last_run"] = now_dt_str
                self._execute_task(task, now_dt_str)

    def _execute_task(self, task, today_str):
        """자동실행 작업 실행 (별도 프로세스)"""
        task_id = task["id"]

        # 중복 실행 방지: 이미 실행중인 작업이면 무시
        if task_id in self._running_tasks:
            self.log(f"[자동실행] {task['name']} 이미 실행중 - 건너뜀")
            return

        self._running_tasks.add(task_id)

        executable = task["executable"]
        arguments = task.get("arguments", "")
        python_venv = task["python_venv"]

        def _run():
            try:
                ext = os.path.splitext(executable)[1].lower()
                if ext in (".py", ".pyw") and python_venv:
                    cmd = [python_venv, executable]
                elif ext in (".py", ".pyw"):
                    # .pyw 호스트의 sys.executable은 pythonw.exe → python.exe로 대체
                    python_exe = os.path.join(os.path.dirname(sys.executable), "python.exe")
                    cmd = [python_exe, executable]
                else:
                    cmd = [executable]

                # 파라미터 추가
                if arguments:
                    import shlex
                    cmd.extend(shlex.split(arguments))

                self.log(f"[자동실행] {task['name']} 실행: {' '.join(cmd)}")
                t_start = time.time()
                proc = subprocess.Popen(
                    cmd,
                    cwd=os.path.dirname(executable) or None,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
                )
                self._task_processes[task_id] = proc
                db_update_task_last_run(task["id"], today_str)
                # 프로세스 종료 대기 (별도 스레드이므로 메인 루프 차단 없음)
                exit_code = proc.wait()
                elapsed = time.time() - t_start
                if elapsed < 60:
                    elapsed_str = f"{elapsed:.1f}초"
                else:
                    elapsed_str = f"{int(elapsed // 60)}분 {int(elapsed % 60)}초"
                if exit_code == 0:
                    self.log(f"[자동실행] {task['name']} 실행 완료 ({elapsed_str})")
                else:
                    self.log(f"[자동실행] {task['name']} 종료 (코드: {exit_code}, {elapsed_str})")
            except Exception as e:
                self.log(f"[자동실행] {task['name']} 실행 실패: {e}")
            finally:
                self._task_processes.pop(task_id, None)
                self._running_tasks.discard(task_id)

        threading.Thread(target=_run, daemon=True).start()

    # ─── 트레이 아이콘 ───────────────────────────────────
    def _setup_tray(self):
        """pystray 트레이 아이콘 설정 (백그라운드 스레드)"""
        try:
            import pystray
            from PIL import Image, ImageDraw

            img = Image.new("RGB", (64, 64), (30, 100, 200))
            draw = ImageDraw.Draw(img)
            draw.rectangle([16, 16, 48, 48], fill=(255, 255, 255))

            self.tray_icon = pystray.Icon("AutoExec", img, "실행 관리 서버")
            self.tray_icon.menu = pystray.Menu(
                pystray.MenuItem("열기", lambda icon, item: self.root.after(0, self._show_window), default=True),
                pystray.MenuItem("종료", lambda icon, item: self.root.after(0, self._force_quit)),
            )
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception:
            pass

    def _show_window(self):
        self.hidden = False
        self.root.deiconify()
        self.root.lift()

    def _force_quit(self):
        self._save_window()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()

    # ─── 종료 처리 ───────────────────────────────────────
    def _on_close(self):
        """닫기 버튼 → 트레이로 최소화 (종료하지 않음)"""
        self._save_window()
        self.hidden = True
        self.root.withdraw()

    def run(self):
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════
#  엔트리포인트
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    # 중복 실행 방지 (Windows Named Mutex)
    import ctypes
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "AutoExec_Python")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        ctypes.windll.kernel32.CloseHandle(_mutex)
        messagebox.showwarning("AutoExec", "이미 실행중입니다.")
        sys.exit(0)

    db_init()
    app = AutoExecApp()
    app.run()
    ctypes.windll.kernel32.CloseHandle(_mutex)
