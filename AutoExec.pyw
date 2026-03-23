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
import ctypes
import ctypes.wintypes
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date, timedelta

from dotenv import load_dotenv
import win11_setup
import win11_folder

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
    # move_targets 테이블에 maximize 필드 추가
    try:
        cur.execute("ALTER TABLE move_targets ADD COLUMN maximize INTEGER NOT NULL DEFAULT 1")
    except Exception:
        pass
    # tasks 테이블에 auto_move 필드 추가
    try:
        cur.execute("ALTER TABLE tasks ADD COLUMN auto_move INTEGER NOT NULL DEFAULT 0")
    except Exception:
        pass
    # tasks 테이블에 위치 필드 추가 (폴더 창 위치 지정용)
    for col, default in [("target_x", 0), ("target_y", 0), ("target_w", 0), ("target_h", 0)]:
        try:
            cur.execute(f"ALTER TABLE tasks ADD COLUMN {col} INTEGER NOT NULL DEFAULT {default}")
        except Exception:
            pass
    cur.execute("""
        CREATE TABLE IF NOT EXISTS closed_days (
            date_int INTEGER PRIMARY KEY,
            reason TEXT NOT NULL DEFAULT ''
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS move_targets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL DEFAULT '',
            exe_name TEXT NOT NULL DEFAULT '',
            enabled INTEGER NOT NULL DEFAULT 1,
            move_mode TEXT NOT NULL DEFAULT 'sub_monitor',
            target_x INTEGER NOT NULL DEFAULT 0,
            target_y INTEGER NOT NULL DEFAULT 0,
            target_w INTEGER NOT NULL DEFAULT 0,
            target_h INTEGER NOT NULL DEFAULT 0,
            sort_order INTEGER NOT NULL DEFAULT 0
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
    if table not in ("pcs", "tasks", "move_targets"):
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
                   repeat_mode="once", repeat_interval=0, repeat_end_time="23:59",
                   target_x=0, target_y=0, target_w=0, target_h=0, auto_move=0):
    """자동실행 추가 또는 수정"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if task_id:
            cur.execute(
                "UPDATE tasks SET name=?, enabled=?, run_time=?, executable=?, "
                "arguments=?, python_venv=?, skip_holiday=?, "
                "repeat_mode=?, repeat_interval=?, repeat_end_time=?, "
                "target_x=?, target_y=?, target_w=?, target_h=?, auto_move=? WHERE id=?",
                (name, int(enabled), run_time, executable, arguments, python_venv, int(skip_holiday),
                 repeat_mode, repeat_interval, repeat_end_time,
                 target_x, target_y, target_w, target_h, int(auto_move), task_id),
            )
        else:
            cur.execute(
                "INSERT INTO tasks (name, enabled, run_time, executable, arguments, python_venv, skip_holiday, "
                "repeat_mode, repeat_interval, repeat_end_time, target_x, target_y, target_w, target_h, auto_move, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM tasks))",
                (name, int(enabled), run_time, executable, arguments, python_venv, int(skip_holiday),
                 repeat_mode, repeat_interval, repeat_end_time,
                 target_x, target_y, target_w, target_h, int(auto_move)),
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
#  창 이동 대상 (move_targets 테이블)
# ═══════════════════════════════════════════════════════════
def db_fetch_move_targets():
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM move_targets ORDER BY sort_order, id")
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_move_target(target_id, name, exe_name, enabled=1, move_mode="sub_monitor",
                          target_x=0, target_y=0, target_w=0, target_h=0, maximize=1):
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if target_id:
            cur.execute(
                "UPDATE move_targets SET name=?, exe_name=?, enabled=?, move_mode=?, "
                "target_x=?, target_y=?, target_w=?, target_h=?, maximize=? WHERE id=?",
                (name, exe_name, int(enabled), move_mode, target_x, target_y, target_w, target_h, int(maximize), target_id),
            )
        else:
            cur.execute(
                "INSERT INTO move_targets (name, exe_name, enabled, move_mode, target_x, target_y, target_w, target_h, maximize, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM move_targets))",
                (name, exe_name, int(enabled), move_mode, target_x, target_y, target_w, target_h, int(maximize)),
            )
        conn.commit()
    finally:
        conn.close()


def db_delete_move_target(target_id):
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM move_targets WHERE id=?", (target_id,))
        conn.commit()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════
#  JSON 로컬 설정 (윈도우 좌표, 최상위)
# ═══════════════════════════════════════════════════════════
def load_local_settings():
    """로컬 UI 설정 로드"""
    defaults = {"window": {"x": 200, "y": 200, "width": 620, "height": 750, "topmost": False}}
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
#  탐색기 창 검색
# ═══════════════════════════════════════════════════════════
def _find_explorer_window_by_title(folder_name):
    """CabinetWClass(탐색기) 창 중 제목이 folder_name으로 시작하는 핸들 반환.
    Windows 탐색기 창 제목은 '폴더명 - 파일 탐색기' 형식."""
    result = [None]
    GetClassNameW = ctypes.windll.user32.GetClassNameW

    def enum_callback(hwnd, lParam):
        if not ctypes.windll.user32.IsWindowVisible(hwnd):
            return True
        cls_buf = ctypes.create_unicode_buffer(64)
        GetClassNameW(hwnd, cls_buf, 64)
        if cls_buf.value != "CabinetWClass":
            return True
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value
        # "폴더명" 또는 "폴더명 - 파일 탐색기" 형식 모두 매칭
        if title == folder_name or title.startswith(folder_name + " -"):
            result[0] = hwnd
            return False
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return result[0]


def _get_process_cmdline(pid):
    """NtQueryInformationProcess로 프로세스 커맨드라인을 읽어 반환 (Win32 API 직접 호출, <0.1ms)"""
    kernel32 = ctypes.windll.kernel32
    ntdll = ctypes.windll.ntdll
    kernel32.ReadProcessMemory.argtypes = [
        ctypes.wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)
    ]
    kernel32.ReadProcessMemory.restype = ctypes.wintypes.BOOL

    class PBI(ctypes.Structure):
        _fields_ = [("R1", ctypes.c_void_p), ("PebBaseAddress", ctypes.c_void_p),
                     ("R2", ctypes.c_void_p * 2), ("UniqueProcessId", ctypes.c_void_p), ("R3", ctypes.c_void_p)]

    class US(ctypes.Structure):
        _fields_ = [("Length", ctypes.c_ushort), ("MaxLength", ctypes.c_ushort),
                     ("_pad", ctypes.c_uint), ("Buffer", ctypes.c_void_p)]

    h = kernel32.OpenProcess(0x0400 | 0x0010, False, pid)  # QUERY_INFORMATION | VM_READ
    if not h:
        return None
    try:
        pbi = PBI()
        ret = ctypes.c_ulong()
        if ntdll.NtQueryInformationProcess(h, 0, ctypes.byref(pbi), ctypes.sizeof(pbi), ctypes.byref(ret)) != 0:
            return None
        rd = ctypes.c_size_t()
        pp = ctypes.c_void_p()
        kernel32.ReadProcessMemory(h, pbi.PebBaseAddress + 0x20, ctypes.byref(pp), ctypes.sizeof(pp), ctypes.byref(rd))
        us = US()
        kernel32.ReadProcessMemory(h, pp.value + 0x70, ctypes.byref(us), ctypes.sizeof(us), ctypes.byref(rd))
        buf = ctypes.create_unicode_buffer(us.Length // 2 + 1)
        kernel32.ReadProcessMemory(h, us.Buffer, buf, us.Length, ctypes.byref(rd))
        return buf.value
    except Exception:
        return None
    finally:
        kernel32.CloseHandle(h)


def _find_window_by_exe_name(exe_name_lower):
    """exe 파일명(소문자)으로 보이는 창 핸들 반환. .py/.pyw는 커맨드라인 매칭."""
    result = [None]
    is_python_script = exe_name_lower.endswith((".py", ".pyw"))

    def enum_callback(hwnd, lParam):
        if not ctypes.windll.user32.IsWindowVisible(hwnd):
            return True
        if ctypes.windll.user32.GetWindowTextLengthW(hwnd) == 0:
            return True
        pid = ctypes.wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)
        if handle:
            try:
                buf = ctypes.create_unicode_buffer(260)
                size = ctypes.wintypes.DWORD(260)
                if ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                    proc_exe = os.path.basename(buf.value).lower()
                    if proc_exe == exe_name_lower:
                        result[0] = hwnd
                        return False
                    # python.exe/pythonw.exe가 실행한 스크립트인지 커맨드라인으로 확인
                    if is_python_script and proc_exe in ("python.exe", "pythonw.exe"):
                        cmdline = _get_process_cmdline(pid.value)
                        if cmdline and exe_name_lower in cmdline.lower():
                            result[0] = hwnd
                            return False
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return result[0]


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
        self.var_skip_holiday = tk.BooleanVar(value=False)
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

        self.bind("<Escape>", lambda e: self.destroy())
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
        "매일 1회": "once", "부팅시 1회": "boot", "N분 간격": "minutes", "N시간 간격": "hours",
        "요일 지정": "weekly", "매월 지정일": "monthly",
    }
    _REPEAT_LABELS = {
        "once": "매일 1회", "boot": "부팅시 1회", "minutes": "N분 간격", "hours": "N시간 간격",
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
        self.var_repeat_mode = tk.StringVar(value="부팅시 1회")
        self.cmb_mode = ttk.Combobox(frame, textvariable=self.var_repeat_mode,
                                     values=list(self._REPEAT_MODES.keys()),
                                     state="readonly", width=12)
        self.cmb_mode.grid(row=2, column=1, sticky=tk.W, padx=(5, 0), pady=3)
        self.cmb_mode.bind("<<ComboboxSelected>>", self._on_mode_changed)

        # 반복 간격 (N분/N시간 모드에서만 표시)
        self.lbl_interval = ttk.Label(frame, text="반복 간격:")
        self.lbl_interval.grid(row=3, column=0, sticky=tk.W, pady=3)
        self.interval_frame = ttk.Frame(frame)
        self.interval_frame.grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=(5, 0), pady=3)
        self.var_interval = tk.IntVar(value=30)
        self.spn_interval = ttk.Spinbox(self.interval_frame, from_=1, to=1440, width=6,
                                        textvariable=self.var_interval)
        self.spn_interval.pack(side=tk.LEFT)
        self.lbl_interval_unit = ttk.Label(self.interval_frame, text="분")
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
        exe_btn_frame = ttk.Frame(frame)
        exe_btn_frame.grid(row=6, column=2, padx=2)
        ttk.Button(exe_btn_frame, text="..", width=3, command=self._browse_exe).pack(side=tk.LEFT)
        ttk.Button(exe_btn_frame, text="폴더", width=4, command=self._browse_folder).pack(side=tk.LEFT, padx=(2, 0))

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
        self.var_skip_holiday = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="휴장일 제외", variable=self.var_skip_holiday).grid(
            row=10, column=0, columnspan=3, sticky=tk.W, pady=3
        )

        # 창위치 자동이동 체크박스
        self.var_auto_move = tk.BooleanVar(value=False)
        self.chk_auto_move = ttk.Checkbutton(
            frame, text="창위치 자동이동", variable=self.var_auto_move,
            command=self._on_auto_move_toggled
        )
        self.chk_auto_move.grid(row=11, column=0, columnspan=3, sticky=tk.W, pady=3)

        # 창 위치 (폴더 열기 시 탐색기 창 위치 지정)
        self.lbl_pos = ttk.Label(frame, text="창 위치:")
        self.lbl_pos.grid(row=12, column=0, sticky=tk.W, pady=3)
        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=12, column=1, columnspan=2, sticky=tk.W, padx=(5, 0), pady=3)
        self.pos_frame = pos_frame
        ttk.Label(pos_frame, text="X").pack(side=tk.LEFT)
        self.var_tx = tk.IntVar(value=0)
        self.ent_tx = ttk.Entry(pos_frame, textvariable=self.var_tx, width=6)
        self.ent_tx.pack(side=tk.LEFT, padx=(2, 5))
        ttk.Label(pos_frame, text="Y").pack(side=tk.LEFT)
        self.var_ty = tk.IntVar(value=0)
        self.ent_ty = ttk.Entry(pos_frame, textvariable=self.var_ty, width=6)
        self.ent_ty.pack(side=tk.LEFT, padx=(2, 5))
        ttk.Label(pos_frame, text="W").pack(side=tk.LEFT)
        self.var_tw = tk.IntVar(value=0)
        self.ent_tw = ttk.Entry(pos_frame, textvariable=self.var_tw, width=6)
        self.ent_tw.pack(side=tk.LEFT, padx=(2, 5))
        ttk.Label(pos_frame, text="H").pack(side=tk.LEFT)
        self.var_th = tk.IntVar(value=0)
        self.ent_th = ttk.Entry(pos_frame, textvariable=self.var_th, width=6)
        self.ent_th.pack(side=tk.LEFT, padx=(2, 5))
        self.btn_capture_pos = ttk.Button(pos_frame, text="위치 캡처", command=self._capture_folder_pos)
        self.btn_capture_pos.pack(side=tk.LEFT, padx=(5, 0))
        self.lbl_pos_hint = ttk.Label(frame, text="(폴더를 열어 원하는 위치에 놓은 뒤 캡처, 0=지정 안 함)", foreground="gray")
        self.lbl_pos_hint.grid(row=13, column=0, columnspan=3, sticky=tk.W)

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
            # 위치 복원
            self.var_auto_move.set(bool(task.get("auto_move", 0)))
            self.var_tx.set(task.get("target_x", 0) or 0)
            self.var_ty.set(task.get("target_y", 0) or 0)
            self.var_tw.set(task.get("target_w", 0) or 0)
            self.var_th.set(task.get("target_h", 0) or 0)

        self._on_mode_changed()
        self._update_pos_visibility()
        self._on_auto_move_toggled()
        # 실행파일 변경 시 위치 필드 표시/숨김 업데이트
        self.ent_exe.bind("<FocusOut>", lambda e: self._update_pos_visibility())

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=14, column=0, columnspan=3, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.bind("<Escape>", lambda e: self.destroy())
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
        self.interval_frame.grid_remove()
        self.lbl_weekdays.grid_remove()
        self.weekday_frame.grid_remove()
        self.lbl_monthday.grid_remove()
        self.monthday_frame.grid_remove()
        # 종료 시간 숨김
        self.lbl_end_time.grid_remove()
        self.ent_end_time.grid_remove()

        if mode == "boot":
            self.lbl_time.grid_remove()
            self.ent_time.grid_remove()
        elif mode in ("minutes", "hours"):
            self.lbl_time.grid()
            self.ent_time.grid()
            self.lbl_interval.grid()
            self.interval_frame.grid()
            self.lbl_end_time.grid()
            self.ent_end_time.grid()
            self.lbl_time.config(text="시작 시간:")
            self.lbl_interval_unit.config(text="분" if mode == "minutes" else "시간")
        elif mode == "weekly":
            self.lbl_time.grid()
            self.ent_time.grid()
            self.lbl_weekdays.grid()
            self.weekday_frame.grid()
            self.lbl_time.config(text="실행 시간:")
        elif mode == "monthly":
            self.lbl_time.grid()
            self.ent_time.grid()
            self.lbl_monthday.grid()
            self.monthday_frame.grid()
            self.lbl_time.config(text="실행 시간:")
        else:
            self.lbl_time.grid()
            self.ent_time.grid()
            self.lbl_time.config(text="실행 시간:")

    def _on_auto_move_toggled(self):
        """창위치 자동이동 체크 시 위치 입력 필드 활성/비활성"""
        enabled = self.var_auto_move.get()
        state = "normal" if enabled else "disabled"
        self.ent_tx.config(state=state)
        self.ent_ty.config(state=state)
        self.ent_tw.config(state=state)
        self.ent_th.config(state=state)
        self.btn_capture_pos.config(state=state)

    def _update_pos_visibility(self):
        """실행 파일이 지정되어 있으면 창 위치 필드 표시"""
        exe = self.ent_exe.get().strip()
        is_folder = os.path.isdir(exe)
        has_exe = bool(exe) and not is_folder
        if is_folder or has_exe:
            self.chk_auto_move.grid()
            self.lbl_pos.grid()
            self.pos_frame.grid()
            if is_folder:
                self.lbl_pos_hint.config(text="(폴더를 열어 원하는 위치에 놓은 뒤 캡처)")
            else:
                self.lbl_pos_hint.config(text="(프로그램을 실행하고 원하는 위치에 놓은 뒤 캡처)")
            self.lbl_pos_hint.grid()
        else:
            self.chk_auto_move.grid_remove()
            self.lbl_pos.grid_remove()
            self.pos_frame.grid_remove()
            self.lbl_pos_hint.grid_remove()

    def _browse_exe(self):
        path = filedialog.askopenfilename(
            title="실행 파일 선택",
            filetypes=[("실행파일", "*.exe *.py *.pyw *.bat"), ("모든파일", "*.*")],
            parent=self,
        )
        if path:
            self.ent_exe.delete(0, tk.END)
            self.ent_exe.insert(0, path)
            self._update_pos_visibility()
            self._auto_fill_python_path(path)

    def _auto_fill_python_path(self, exe_path):
        """Python 스크립트 선택 시 Python 경로 자동 입력"""
        if self.ent_venv.get().strip():
            return  # 이미 지정되어 있으면 건드리지 않음
        ext = os.path.splitext(exe_path)[1].lower()
        if ext == ".pyw":
            python_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        elif ext == ".py":
            python_path = os.path.join(os.path.dirname(sys.executable), "python.exe")
        else:
            return
        if os.path.isfile(python_path):
            self.ent_venv.delete(0, tk.END)
            self.ent_venv.insert(0, python_path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="폴더 선택", parent=self)
        if path:
            self.ent_exe.delete(0, tk.END)
            self.ent_exe.insert(0, path)
            self._update_pos_visibility()

    def _capture_folder_pos(self):
        """열려 있는 창의 위치/크기를 캡처 (폴더: 탐색기, 실행파일: exe명으로 검색)"""
        exe_path = self.ent_exe.get().strip()
        if not exe_path:
            messagebox.showwarning("알림", "먼저 실행 파일 또는 폴더를 선택하세요.", parent=self)
            return

        target_hwnd = None
        if os.path.isdir(exe_path):
            folder_name = os.path.basename(exe_path.rstrip("\\/"))
            target_hwnd = _find_explorer_window_by_title(folder_name)
            if not target_hwnd:
                messagebox.showinfo("알림", f"'{folder_name}' 탐색기 창을 찾을 수 없습니다.\n폴더를 먼저 열어주세요.", parent=self)
                return
        else:
            exe_name = os.path.basename(exe_path).lower()
            target_hwnd = _find_window_by_exe_name(exe_name)
            if not target_hwnd:
                messagebox.showinfo("알림", f"'{os.path.basename(exe_path)}' 창을 찾을 수 없습니다.\n프로그램을 먼저 실행해주세요.", parent=self)
                return

        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(target_hwnd, ctypes.byref(rect))
        self.var_tx.set(rect.left)
        self.var_ty.set(rect.top)
        self.var_tw.set(rect.right - rect.left)
        self.var_th.set(rect.bottom - rect.top)

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
            "auto_move": self.var_auto_move.get(),
            "target_x": self.var_tx.get(),
            "target_y": self.var_ty.get(),
            "target_w": self.var_tw.get(),
            "target_h": self.var_th.get(),
        }
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  메인 윈도우
# ═══════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════
#  창 이동 대상 편집 다이얼로그
# ═══════════════════════════════════════════════════════════
class MoveTargetEditDialog(tk.Toplevel):
    _MODE_LABELS = {"서브 모니터": "sub_monitor", "현재 위치 저장": "save_position", "좌표 직접 입력": "custom"}
    _MODE_KEYS = {"sub_monitor": "서브 모니터", "save_position": "현재 위치 저장", "custom": "좌표 직접 입력"}

    def __init__(self, parent, target=None):
        super().__init__(parent)
        self.result = None
        self.target = target
        self.title("이동 대상 편집" if target else "이동 대상 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="이름:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_name = ttk.Entry(frame, width=30)
        self.ent_name.grid(row=0, column=1, columnspan=3, pady=3, padx=(5, 0))

        ttk.Label(frame, text="프로세스명:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_exe = ttk.Entry(frame, width=30)
        self.ent_exe.grid(row=1, column=1, columnspan=3, pady=3, padx=(5, 0))
        ttk.Label(frame, text="(예: chrome.exe)", foreground="gray").grid(
            row=2, column=0, columnspan=4, sticky=tk.W
        )

        self.var_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="자동 이동 사용", variable=self.var_enabled).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=3
        )

        self.var_maximize = tk.BooleanVar(value=True)
        self.chk_maximize = ttk.Checkbutton(frame, text="최대화", variable=self.var_maximize)
        self.chk_maximize.grid(row=3, column=2, columnspan=2, sticky=tk.W, pady=3)

        ttk.Label(frame, text="이동 위치:").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.var_mode = tk.StringVar(value="서브 모니터")
        self.cmb_mode = ttk.Combobox(frame, textvariable=self.var_mode,
                                     values=list(self._MODE_LABELS.keys()),
                                     state="readonly", width=15)
        self.cmb_mode.grid(row=4, column=1, columnspan=3, sticky=tk.W, padx=(5, 0), pady=3)
        self.cmb_mode.bind("<<ComboboxSelected>>", self._on_mode_changed)

        # 좌표 입력 (custom 모드에서만 표시)
        self.coord_frame = ttk.Frame(frame)
        self.coord_frame.grid(row=5, column=0, columnspan=4, sticky=tk.W, pady=3)
        ttk.Label(self.coord_frame, text="X:").pack(side=tk.LEFT)
        self.ent_x = ttk.Entry(self.coord_frame, width=6)
        self.ent_x.pack(side=tk.LEFT, padx=(2, 6))
        self.ent_x.insert(0, "0")
        ttk.Label(self.coord_frame, text="Y:").pack(side=tk.LEFT)
        self.ent_y = ttk.Entry(self.coord_frame, width=6)
        self.ent_y.pack(side=tk.LEFT, padx=(2, 6))
        self.ent_y.insert(0, "0")
        ttk.Label(self.coord_frame, text="W:").pack(side=tk.LEFT)
        self.ent_w = ttk.Entry(self.coord_frame, width=6)
        self.ent_w.pack(side=tk.LEFT, padx=(2, 6))
        self.ent_w.insert(0, "0")
        ttk.Label(self.coord_frame, text="H:").pack(side=tk.LEFT)
        self.ent_h = ttk.Entry(self.coord_frame, width=6)
        self.ent_h.pack(side=tk.LEFT, padx=(2, 0))
        self.ent_h.insert(0, "0")
        ttk.Label(self.coord_frame, text="(W,H=0: 현재 크기 유지)", foreground="gray").pack(side=tk.LEFT, padx=(8, 0))

        # 위치 저장 프레임 (save_position 모드에서만 표시)
        self.save_pos_frame = ttk.Frame(frame)
        self.save_pos_frame.grid(row=6, column=0, columnspan=4, sticky=tk.W, pady=3)
        ttk.Button(self.save_pos_frame, text="현재 위치 캡처", command=self._capture_position).pack(side=tk.LEFT)
        self.lbl_saved_pos = ttk.Label(self.save_pos_frame, text="", foreground="gray")
        self.lbl_saved_pos.pack(side=tk.LEFT, padx=(8, 0))

        if target:
            self.ent_name.insert(0, target["name"])
            self.ent_exe.insert(0, target["exe_name"])
            self.var_enabled.set(bool(target.get("enabled", 1)))
            self.var_maximize.set(bool(target.get("maximize", 1)))
            mode = target.get("move_mode", "sub_monitor")
            self.var_mode.set(self._MODE_KEYS.get(mode, "서브 모니터"))
            if mode == "custom":
                self.ent_x.delete(0, tk.END)
                self.ent_x.insert(0, str(target.get("target_x", 0)))
                self.ent_y.delete(0, tk.END)
                self.ent_y.insert(0, str(target.get("target_y", 0)))
                self.ent_w.delete(0, tk.END)
                self.ent_w.insert(0, str(target.get("target_w", 0)))
                self.ent_h.delete(0, tk.END)
                self.ent_h.insert(0, str(target.get("target_h", 0)))

        self._on_mode_changed()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=7, column=0, columnspan=4, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_x()
        py = parent.winfo_y()
        dw = self.winfo_width()
        dh = self.winfo_height()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.ent_name.focus_set()

    def _on_mode_changed(self, event=None):
        mode_label = self.var_mode.get()
        mode = self._MODE_LABELS.get(mode_label, "sub_monitor")
        self.coord_frame.grid_remove()
        self.save_pos_frame.grid_remove()
        if mode == "custom":
            self.coord_frame.grid()
        elif mode == "save_position":
            self.save_pos_frame.grid()
            self._update_saved_pos_label()
        # 모드 변경 시 최대화 기본값 (사용자가 직접 변경한 경우가 아닐 때만)
        if event is not None:
            self.var_maximize.set(mode == "sub_monitor")

    def _update_saved_pos_label(self):
        """저장된 위치 표시"""
        src = getattr(self, "_captured", None) or (self.target if self.target else None)
        if src:
            x, y = src.get("target_x", 0), src.get("target_y", 0)
            w, h = src.get("target_w", 0), src.get("target_h", 0)
            if x or y or w or h:
                size = f" {w}x{h}" if w and h else ""
                self.lbl_saved_pos.config(text=f"저장된 위치: {x},{y}{size}")
                return
        self.lbl_saved_pos.config(text="저장된 위치 없음 - 캡처하세요")

    def _capture_position(self):
        """현재 실행중인 프로세스의 창 위치를 캡처"""
        exe_name = self.ent_exe.get().strip().lower()
        if not exe_name:
            messagebox.showwarning("입력 오류", "프로세스명을 먼저 입력하세요.", parent=self)
            return

        hwnds = []

        def enum_callback(hwnd, lParam):
            if not ctypes.windll.user32.IsWindowVisible(hwnd):
                return True
            if ctypes.windll.user32.GetWindowTextLengthW(hwnd) == 0:
                return True
            pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)
            if handle:
                try:
                    buf = ctypes.create_unicode_buffer(260)
                    size = ctypes.wintypes.DWORD(260)
                    if ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                        if os.path.basename(buf.value).lower() == exe_name:
                            hwnds.append(hwnd)
                finally:
                    ctypes.windll.kernel32.CloseHandle(handle)
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)

        if not hwnds:
            messagebox.showinfo("알림", f"{exe_name} 실행중인 창을 찾을 수 없습니다.", parent=self)
            return

        # 첫 번째 창의 위치 캡처
        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnds[0], ctypes.byref(rect))
        x, y = rect.left, rect.top
        w, h = rect.right - rect.left, rect.bottom - rect.top

        # 캡처 좌표 저장
        self._captured = {"target_x": x, "target_y": y, "target_w": w, "target_h": h}

        self.lbl_saved_pos.config(text=f"캡처 완료: {x},{y} {w}x{h}")

    def _on_ok(self):
        name = self.ent_name.get().strip()
        exe_name = self.ent_exe.get().strip()
        if not name or not exe_name:
            messagebox.showwarning("입력 오류", "이름과 프로세스명을 입력하세요.", parent=self)
            return
        mode = self._MODE_LABELS.get(self.var_mode.get(), "sub_monitor")
        target_x = target_y = target_w = target_h = 0
        if mode == "custom":
            try:
                target_x = int(self.ent_x.get())
                target_y = int(self.ent_y.get())
                target_w = int(self.ent_w.get())
                target_h = int(self.ent_h.get())
            except ValueError:
                messagebox.showwarning("입력 오류", "좌표는 정수로 입력하세요.", parent=self)
                return
        elif mode == "save_position":
            src = getattr(self, "_captured", None) or (self.target if self.target else {})
            target_x = src.get("target_x", 0)
            target_y = src.get("target_y", 0)
            target_w = src.get("target_w", 0)
            target_h = src.get("target_h", 0)
        self.result = {
            "id": self.target.get("id") if self.target else None,
            "name": name,
            "exe_name": exe_name,
            "enabled": self.var_enabled.get(),
            "move_mode": mode,
            "target_x": target_x, "target_y": target_y,
            "target_w": target_w, "target_h": target_h,
            "maximize": self.var_maximize.get(),
        }
        self.destroy()


# ═══════════════════════════════════════════════════════════
#  레지스트리 관리 다이얼로그
# ═══════════════════════════════════════════════════════════
class RegistryItemEditDialog(tk.Toplevel):
    """레지스트리 항목 추가/편집 다이얼로그"""
    _TYPES = ["REG_DWORD", "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_QWORD", "REG_MULTI_SZ"]

    def __init__(self, parent, item=None):
        super().__init__(parent)
        self.result = None
        self.title("레지스트리 편집" if item else "레지스트리 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="설명:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_desc = ttk.Entry(frame, width=50)
        self.ent_desc.grid(row=0, column=1, columnspan=2, pady=3, padx=(5, 0))

        ttk.Label(frame, text="경로:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_path = ttk.Entry(frame, width=50)
        self.ent_path.grid(row=1, column=1, columnspan=2, pady=3, padx=(5, 0))

        ttk.Label(frame, text="이름:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_name = ttk.Entry(frame, width=30)
        self.ent_name.grid(row=2, column=1, columnspan=2, pady=3, padx=(5, 0), sticky=tk.W)

        ttk.Label(frame, text="유형:").grid(row=3, column=0, sticky=tk.W, pady=3)
        self.var_type = tk.StringVar(value="REG_DWORD")
        ttk.Combobox(frame, textvariable=self.var_type, values=self._TYPES,
                     state="readonly", width=15).grid(row=3, column=1, sticky=tk.W, padx=(5, 0), pady=3)

        ttk.Label(frame, text="값:").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.ent_value = ttk.Entry(frame, width=50)
        self.ent_value.grid(row=4, column=1, columnspan=2, pady=3, padx=(5, 0))

        if item:
            self.ent_desc.insert(0, item.get("description", ""))
            self.ent_path.insert(0, item.get("path", ""))
            self.ent_name.insert(0, item.get("name", ""))
            self.var_type.set(item.get("type", "REG_DWORD"))
            self.ent_value.insert(0, str(item.get("value", "")))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        dw, dh = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.focus_force()
        self.ent_desc.focus_set()

    def _on_ok(self):
        path = self.ent_path.get().strip()
        name = self.ent_name.get().strip()
        if not path or not name:
            messagebox.showwarning("입력 오류", "경로와 이름을 입력하세요.", parent=self)
            return
        reg_type = self.var_type.get()
        value_str = self.ent_value.get().strip()
        # 값 변환
        if reg_type in ("REG_DWORD", "REG_QWORD"):
            try:
                value = int(value_str)
            except ValueError:
                messagebox.showwarning("입력 오류", "정수 값을 입력하세요.", parent=self)
                return
        else:
            value = value_str
        self.result = {
            "path": path, "name": name, "type": reg_type,
            "value": value, "description": self.ent_desc.get().strip(),
        }
        self.destroy()


class RegistryDialog(tk.Toplevel):
    def __init__(self, parent, log_callback):
        super().__init__(parent)
        self.log_callback = log_callback
        self.title("레지스트리 관리")
        self.geometry("700x450")
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("설명", "경로", "이름", "값")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=15)
        self.tree.heading("설명", text="설명")
        self.tree.heading("경로", text="레지스트리 경로")
        self.tree.heading("이름", text="이름")
        self.tree.heading("값", text="값")
        self.tree.column("설명", width=250)
        self.tree.column("경로", width=200)
        self.tree.column("이름", width=120)
        self.tree.column("값", width=80)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky=tk.NS)
        self.tree.config(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", self._apply_one)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        ttk.Button(btn_frame, text="추가", command=self._add_item).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="편집", command=self._edit_item).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="삭제", command=self._delete_item).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(btn_frame, text="전체 적용", command=self._apply_all).pack(side=tk.LEFT)

        self._load_items()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())
        self.transient(parent)
        self.update_idletasks()
        saved = _dialog_positions.get("RegistryDialog")
        if saved:
            self.geometry(saved)
        else:
            pw, ph = parent.winfo_width(), parent.winfo_height()
            px, py = parent.winfo_x(), parent.winfo_y()
            dw, dh = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.focus_force()

    def _on_close(self):
        _dialog_positions["RegistryDialog"] = self.geometry()
        self.destroy()

    def _load_items(self):
        for c in self.tree.get_children():
            self.tree.delete(c)
        self.items = win11_setup.load_registry_items()
        for i, item in enumerate(self.items):
            desc = item.get("description", item["name"])
            short_path = item["path"].split("\\")[-1] if "\\" in item["path"] else item["path"]
            self.tree.insert("", tk.END, iid=str(i),
                             values=(desc, short_path, item["name"], str(item["value"])[:30]))

    def _apply_one(self, event):
        """더블클릭: 해당 항목 1개 적용"""
        row = self.tree.identify_row(event.y)
        if not row:
            return
        item = self.items[int(row)]
        desc = item.get("description", item["name"])
        self.log_callback(f"[레지스트리] 적용: {desc}")

        def _worker():
            win11_setup.apply_registry_items([item], self.log_callback)

        threading.Thread(target=_worker, daemon=True).start()

    def _add_item(self):
        dlg = RegistryItemEditDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self.items.append(dlg.result)
            win11_setup.save_registry_items(self.items)
            self._load_items()

    def _edit_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        dlg = RegistryItemEditDialog(self, self.items[idx])
        self.wait_window(dlg)
        if dlg.result:
            self.items[idx] = dlg.result
            win11_setup.save_registry_items(self.items)
            self._load_items()

    def _delete_item(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        desc = self.items[idx].get("description", self.items[idx]["name"])
        if messagebox.askyesno("삭제 확인", f"'{desc}' 항목을 삭제하시겠습니까?", parent=self):
            del self.items[idx]
            win11_setup.save_registry_items(self.items)
            self._load_items()

    def _apply_all(self):
        if not self.items:
            return
        self.log_callback(f"[레지스트리] {len(self.items)}개 항목 전체 적용 시작")

        def _worker():
            success, fail = win11_setup.apply_registry_items(self.items, self.log_callback)
            self.log_callback(f"[레지스트리] 완료: 성공 {success}개, 실패 {fail}개")

        threading.Thread(target=_worker, daemon=True).start()


# 다이얼로그 위치 기억 저장소 (클래스명 → geometry 문자열)
_dialog_positions: dict[str, str] = {}


# ═══════════════════════════════════════════════════════════
#  백업 경로 관리 다이얼로그
# ═══════════════════════════════════════════════════════════
class BackupPathsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("백업 경로 관리")
        self.geometry("650x400")
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("경로", "목적지", "서비스", "제외")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=15)
        self.tree.heading("경로", text="원본 경로")
        self.tree.heading("목적지", text="백업 폴더명")
        self.tree.heading("서비스", text="서비스")
        self.tree.heading("제외", text="제외 항목")
        self.tree.column("경로", width=300)
        self.tree.column("목적지", width=120)
        self.tree.column("서비스", width=70)
        self.tree.column("제외", width=120)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky=tk.NS)
        self.tree.config(yscrollcommand=scroll.set)

        # 백업 저장 위치
        dest_frame = ttk.Frame(frame)
        dest_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6, 0))
        ttk.Label(dest_frame, text="백업 위치:").pack(side=tk.LEFT)
        self.var_dest = tk.StringVar(value=win11_folder.get_last_backup_destination() or "")
        self.ent_dest = ttk.Entry(dest_frame, textvariable=self.var_dest, width=50)
        self.ent_dest.pack(side=tk.LEFT, padx=(5, 3), fill=tk.X, expand=True)
        ttk.Button(dest_frame, text="..", width=3, command=self._browse_dest).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        ttk.Button(btn_frame, text="추가", command=self._add_path).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="삭제", command=self._remove_path).pack(side=tk.LEFT)

        self._load_items()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Escape>", lambda e: self._on_close())
        self.transient(parent)
        self.update_idletasks()
        saved = _dialog_positions.get("BackupPathsDialog")
        if saved:
            self.geometry(saved)
        else:
            pw, ph = parent.winfo_width(), parent.winfo_height()
            px, py = parent.winfo_x(), parent.winfo_y()
            dw, dh = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.focus_force()

    def _on_close(self):
        # 백업 위치 저장
        dest = self.var_dest.get().strip()
        if dest:
            win11_folder.save_last_backup_destination(dest)
        _dialog_positions["BackupPathsDialog"] = self.geometry()
        self.destroy()

    def _browse_dest(self):
        path = filedialog.askdirectory(title="백업 저장 위치 선택", initialdir=self.var_dest.get() or None, parent=self)
        if path:
            self.var_dest.set(os.path.normpath(path))

    def _load_items(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        config = win11_folder.load_config()
        self.paths = config.get("backup_paths", [])
        for i, item in enumerate(self.paths):
            info = win11_folder.normalize_path_item(item)
            expanded = win11_folder.expand_path(info["path"])
            dest = info.get("destination") or ""
            svc = info.get("service") or ""
            exclude = ", ".join(info.get("exclude") or [])
            self.tree.insert("", tk.END, iid=str(i), values=(expanded, dest, svc, exclude))

    def _add_path(self):
        path = filedialog.askdirectory(title="백업할 폴더/파일 선택", parent=self)
        if not path:
            return

        path = os.path.normpath(path)
        # 목적지 이름 입력
        dest_name = os.path.basename(path)

        # 중복 체크
        config = win11_folder.load_config()
        for item in config.get("backup_paths", []):
            info = win11_folder.normalize_path_item(item)
            if os.path.normpath(win11_folder.expand_path(info["path"])) == path:
                messagebox.showinfo("알림", "이미 등록된 경로입니다.", parent=self)
                return

        config["backup_paths"].append({"path": path, "destination": dest_name})
        win11_folder.save_config(config)
        self._load_items()

    def _remove_path(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("알림", "삭제할 항목을 선택하세요.", parent=self)
            return
        idx = int(sel[0])
        info = win11_folder.normalize_path_item(self.paths[idx])
        if not messagebox.askyesno("삭제 확인", f"'{info['path']}' 경로를 제거하시겠습니까?", parent=self):
            return
        config = win11_folder.load_config()
        del config["backup_paths"][idx]
        win11_folder.save_config(config)
        self._load_items()


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
        self._last_monitor_count = ctypes.windll.user32.GetSystemMetrics(80)  # SM_CMONITORS
        self._monitor_check_counter = 0

        self._build_ui()
        self._restore_window()
        self._refresh_pc_list()
        self._refresh_task_list()
        self._refresh_move_list()
        self._tick()
        # 트레이 아이콘은 mainloop 진입 후 생성 (윈도우 준비 완료 후)
        self.root.after(500, self._setup_tray)
        # 부팅시 1회 실행 태스크 처리
        self.root.after(1000, self._run_boot_tasks)

    # ─── UI 구성 ──────────────────────────────────────────
    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)

        # ── 메뉴바 ──
        self.var_topmost = tk.BooleanVar(value=self.settings["window"].get("topmost", False))
        self._build_menubar()
        self._apply_topmost()

        # ── row 1: GitHub 다운로드 ──
        git_frame = ttk.Frame(root)
        git_frame.grid(row=1, column=0, sticky=tk.EW, padx=8, pady=(3, 2))
        git_frame.columnconfigure(1, weight=1)

        ttk.Label(git_frame, text="깃허브:").grid(row=0, column=0, padx=(0, 3))

        self.git_url_var = tk.StringVar()
        self.git_url_entry = ttk.Entry(git_frame, textvariable=self.git_url_var, font=("Consolas", 9))
        self.git_url_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0, 3))
        self.git_url_entry.bind("<Return>", lambda e: self._git_download())
        self.git_url_entry.bind("<<Paste>>", self._on_git_paste)

        self.git_dl_btn = ttk.Button(git_frame, text="Git", width=4, command=self._git_download)
        self.git_dl_btn.grid(row=0, column=2)

        self.var_git_open_folder = tk.BooleanVar(value=self.settings.get("git_open_folder", False))
        ttk.Checkbutton(git_frame, text="완료시 폴더 열기", variable=self.var_git_open_folder,
                        command=self._save_git_open_folder).grid(row=0, column=3, padx=(3, 0))

        # ── row 2: 자동실행 (전체 폭) ──
        lf_task = ttk.LabelFrame(root, text="자동실행", padding=5)
        lf_task.grid(row=2, column=0, sticky=tk.NSEW, padx=8, pady=2)
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
        self.task_tree.bind("<Button-3>", self._on_task_right_click)

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

        # ── row 3: 창 이동 대상 (좌) + PC 관리 (우) ──
        mid_frame = ttk.Frame(root)
        mid_frame.grid(row=3, column=0, sticky=tk.NSEW, padx=8, pady=2)
        mid_frame.columnconfigure(0, weight=1)
        mid_frame.rowconfigure(0, weight=1)

        lf_move = ttk.LabelFrame(mid_frame, text="창 이동 대상 (듀얼 모니터 감지 시)", padding=5)
        lf_move.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 4))
        lf_move.columnconfigure(0, weight=1)
        lf_move.rowconfigure(0, weight=1)

        move_frame = ttk.Frame(lf_move)
        move_frame.grid(row=0, column=0, sticky=tk.NSEW)
        move_frame.columnconfigure(0, weight=1)
        move_frame.rowconfigure(0, weight=1)

        move_cols = ("사용", "이름", "프로세스명", "이동위치")
        self.move_tree = ttk.Treeview(move_frame, columns=move_cols, show="headings", height=4)
        self.move_tree.heading("사용", text="사용")
        self.move_tree.heading("이름", text="이름")
        self.move_tree.heading("프로세스명", text="프로세스명")
        self.move_tree.heading("이동위치", text="이동 위치")
        self.move_tree.column("사용", width=40, anchor=tk.CENTER)
        self.move_tree.column("이름", width=130)
        self.move_tree.column("프로세스명", width=110)
        self.move_tree.column("이동위치", width=150)
        self.move_tree.grid(row=0, column=0, sticky=tk.NSEW)
        move_scroll = ttk.Scrollbar(move_frame, orient=tk.VERTICAL, command=self.move_tree.yview)
        move_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.move_tree.config(yscrollcommand=move_scroll.set)
        self.move_tree.bind("<Double-1>", self._on_move_double_click)

        move_btn_frame = ttk.Frame(lf_move)
        move_btn_frame.grid(row=1, column=0, sticky=tk.W, pady=(3, 0))
        ttk.Button(move_btn_frame, text="추가", width=6, command=self._add_move_target).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(move_btn_frame, text="편집", width=6, command=self._edit_move_target).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(move_btn_frame, text="삭제", width=6, command=self._delete_move_target).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(move_btn_frame, text="이동", width=6, command=self._manual_move_target).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(move_btn_frame, text="\u25b2", width=3, command=lambda: self._reorder_move_target(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(move_btn_frame, text="\u25bc", width=3, command=lambda: self._reorder_move_target(1)).pack(side=tk.LEFT)

        # ── PC 관리 (우측, 고정 폭) ──
        lf_pc = ttk.LabelFrame(mid_frame, text="PC 관리", padding=5)
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

        # ── row 4: 로그 (전체 폭) ──
        lf_log = ttk.LabelFrame(root, text="로그", padding=5)
        lf_log.grid(row=4, column=0, sticky=tk.NSEW, padx=8, pady=(2, 8))
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

        # 행 가중치: row 0,1(체크박스,Git) 고정, row 2(자동실행) 고정, row 3(이동대상+PC) 고정, row 4(로그) 확장
        root.rowconfigure(0, weight=0)
        root.rowconfigure(1, weight=0)
        root.rowconfigure(2, weight=0)
        root.rowconfigure(3, weight=0)
        root.rowconfigure(4, weight=1)

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

    def _save_git_open_folder(self):
        self.settings["git_open_folder"] = self.var_git_open_folder.get()
        save_local_settings(self.settings)

    # ─── 최상위 ──────────────────────────────────────────
    def _toggle_topmost(self):
        self.settings["window"]["topmost"] = self.var_topmost.get()
        self._apply_topmost()
        save_local_settings(self.settings)

    def _apply_topmost(self):
        self.root.attributes("-topmost", self.var_topmost.get())

    # ─── 메뉴바 ─────────────────────────────────────────
    def _build_menubar(self):
        menubar = tk.Menu(self.root)

        # 설정 메뉴
        menu_settings = tk.Menu(menubar, tearoff=0)
        menu_settings.add_checkbutton(label="최상위", variable=self.var_topmost, command=self._toggle_topmost)
        menu_settings.add_command(label="인트라넷 등록", command=self._menu_intranet)
        menubar.add_cascade(label="설정", menu=menu_settings)

        # 관리 메뉴
        menu_backup = tk.Menu(menubar, tearoff=0)
        menu_backup.add_command(label="레지스트리 관리...", command=self._menu_registry)
        menu_backup.add_command(label="백업 경로 관리...", command=self._menu_backup_paths)
        menu_backup.add_separator()
        menu_backup.add_command(label="폴더 백업...", command=self._menu_backup)
        menu_backup.add_command(label="폴더 복구...", command=self._menu_restore)
        menubar.add_cascade(label="관리", menu=menu_backup)

        self.root.config(menu=menubar)

    def _menu_intranet(self):
        def _worker():
            win11_setup.setup_intranet_zone(self.log)
        threading.Thread(target=_worker, daemon=True).start()

    def _menu_registry(self):
        RegistryDialog(self.root, self.log)

    def _menu_backup_paths(self):
        BackupPathsDialog(self.root)

    def _menu_backup(self):
        dest = win11_folder.get_last_backup_destination()
        if not dest or not os.path.isdir(dest):
            dest = filedialog.askdirectory(title="백업 폴더 선택", initialdir=dest, parent=self.root)
            if not dest:
                return

        def _worker():
            win11_folder.backup(dest, self.log)

        threading.Thread(target=_worker, daemon=True).start()

    def _menu_restore(self):
        last_dest = win11_folder.get_last_backup_destination()
        src = filedialog.askdirectory(title="복구할 백업 폴더 선택", initialdir=last_dest, parent=self.root)
        if not src:
            return
        if not messagebox.askyesno("복구 확인", "기존 파일이 덮어쓰기될 수 있습니다.\n계속하시겠습니까?", parent=self.root):
            return

        def _worker():
            win11_folder.restore(src, self.log)

        threading.Thread(target=_worker, daemon=True).start()

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
            is_folder = os.path.isdir(task["executable"])
            exe_display = os.path.basename(task["executable"]) if not is_folder else task["executable"]
            if not is_folder and task["python_venv"]:
                exe_display += " (venv)"
            # 시간 표시
            repeat_mode = task.get("repeat_mode", "once")
            if repeat_mode == "boot":
                time_display = "부팅시" if not is_folder else "부팅시 폴더"
            elif is_folder:
                time_display = "폴더"
            else:
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
                           r["python_venv"], r["skip_holiday"], r["repeat_mode"], r["repeat_interval"], r["repeat_end_time"],
                           r.get("target_x", 0), r.get("target_y", 0), r.get("target_w", 0), r.get("target_h", 0),
                           r.get("auto_move", 0))
            self._refresh_task_list()
            self.log(f"자동실행 추가: {r['name']}")

    @staticmethod
    def _force_foreground(hwnd):
        """창을 강제로 포그라운드로 활성화"""
        if ctypes.windll.user32.IsIconic(hwnd):
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        # AttachThreadInput 트릭: 포그라운드 스레드에 붙어서 SetForegroundWindow 허용
        fore_hwnd = ctypes.windll.user32.GetForegroundWindow()
        fore_tid = ctypes.windll.user32.GetWindowThreadProcessId(fore_hwnd, None)
        target_tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        if fore_tid != target_tid:
            ctypes.windll.user32.AttachThreadInput(fore_tid, target_tid, True)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.AttachThreadInput(fore_tid, target_tid, False)
        else:
            ctypes.windll.user32.SetForegroundWindow(hwnd)

    def _activate_window_by_pid(self, pid):
        """PID로 창을 찾아 포그라운드로 활성화. 성공 시 True."""
        target_hwnd = None

        def enum_callback(hwnd, lParam):
            nonlocal target_hwnd
            if not ctypes.windll.user32.IsWindowVisible(hwnd):
                return True
            w_pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(w_pid))
            if w_pid.value == pid:
                target_hwnd = hwnd
                return False
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
        if not target_hwnd:
            self.log(f"[디버그] PID {pid} 에 해당하는 창을 찾지 못함")
            return False
        self._force_foreground(target_hwnd)
        return True

    def _activate_window_by_exe(self, exe_name):
        """exe 이름으로 창을 찾아 포그라운드로 활성화. 성공 시 True."""
        hwnds = self._find_windows_by_exe(exe_name.lower())
        if not hwnds:
            return False
        self._force_foreground(hwnds[0])
        return True

    def _on_task_double_click(self, event):
        """더블클릭: 실행중이면 활성화, 폴더 → 열기, 사용 → 편집, 아닌 경우 → 실행"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        task_id = int(item)
        task = next((t for t in self.task_data if t["id"] == task_id), None)
        if not task:
            return
        executable = task["executable"]
        if os.path.isdir(executable):
            self._open_folder_task(task)
            return
        # 1) AutoExec이 실행한 프로세스 → PID로 창 활성화
        if task_id in self._running_tasks:
            proc = self._task_processes.get(task_id)
            if proc and proc.poll() is None and self._activate_window_by_pid(proc.pid):
                self.log(f"[자동실행] {task['name']} 창 활성화")
                return
        # 2) 외부에서 실행중인 프로세스 → exe/스크립트명으로 창 활성화
        if self._activate_window_by_exe(os.path.basename(executable)):
            self.log(f"[자동실행] {task['name']} 창 활성화")
            return
        if task["enabled"]:
            self._edit_task()
        else:
            self._run_task()

    def _on_task_right_click(self, event):
        """우클릭: 컨텍스트 메뉴 표시"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        self.task_tree.selection_set(item)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="실행", command=self._run_task)
        menu.add_command(label="편집", command=self._edit_task)
        menu.add_command(label="삭제", command=self._delete_task)
        menu.add_separator()
        menu.add_command(label="위로", command=lambda: self._move_task(-1))
        menu.add_command(label="아래로", command=lambda: self._move_task(1))
        menu.tk_popup(event.x_root, event.y_root)

    def _edit_task(self):
        task = self._get_selected_task()
        if not task:
            return
        dlg = TaskEditDialog(self.root, task)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_task(r["id"], r["name"], r["enabled"], r["run_time"], r["executable"], r["arguments"],
                           r["python_venv"], r["skip_holiday"], r["repeat_mode"], r["repeat_interval"], r["repeat_end_time"],
                           r.get("target_x", 0), r.get("target_y", 0), r.get("target_w", 0), r.get("target_h", 0),
                           r.get("auto_move", 0))
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

    def _open_folder_task(self, task):
        """폴더 태스크를 열고 auto_move가 켜져 있고 위치 지정이 있으면 이동"""
        executable = task["executable"]
        folder_path = os.path.normpath(executable)  # 포워드 슬래시 → 백슬래시 변환

        subprocess.Popen(["explorer.exe", folder_path])
        self.log(f"[폴더] {task['name']} 열기: {folder_path}")

        if task.get("auto_move", 0):
            tx = task.get("target_x", 0) or 0
            ty = task.get("target_y", 0) or 0
            tw = task.get("target_w", 0) or 0
            th = task.get("target_h", 0) or 0
            has_pos = tx or ty or tw or th
            if has_pos:
                folder_name = os.path.basename(folder_path)
                def _move():
                    time.sleep(1.5)
                    self._move_explorer_window(folder_name, tx, ty, tw, th)
                threading.Thread(target=_move, daemon=True).start()

    def _open_startup_folder(self):
        """윈도우 시작프로그램 폴더 열기"""
        startup = os.path.join(os.getenv("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
        os.startfile(startup)

    # ─── 창 이동 대상 관리 ─────────────────────────────────
    _MOVE_MODE_DISPLAY = {"sub_monitor": "서브 모니터", "save_position": "저장 위치", "custom": "좌표 직접"}

    def _refresh_move_list(self):
        self.move_data = db_fetch_move_targets()
        for item in self.move_tree.get_children():
            self.move_tree.delete(item)
        for t in self.move_data:
            enabled_mark = "O" if t.get("enabled", 1) else ""
            mode = t.get("move_mode", "sub_monitor")
            mode_display = self._MOVE_MODE_DISPLAY.get(mode) or str(mode)
            if mode in ("custom", "save_position"):
                x, y = t.get("target_x", 0), t.get("target_y", 0)
                w, h = t.get("target_w", 0), t.get("target_h", 0)
                if x or y or w or h:
                    size = f" {w}x{h}" if w and h else ""
                    mode_display += f" ({x},{y}{size})"
            self.move_tree.insert("", tk.END, iid=str(t["id"]),
                                  values=(enabled_mark, t["name"], t["exe_name"], mode_display))

    def _get_selected_move_target(self):
        sel = self.move_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "이동 대상을 선택하세요.", parent=self.root)
            return None
        target_id = int(sel[0])
        for t in self.move_data:
            if t["id"] == target_id:
                return t
        return None

    def _add_move_target(self):
        dlg = MoveTargetEditDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_move_target(None, r["name"], r["exe_name"], r["enabled"], r["move_mode"],
                                  r["target_x"], r["target_y"], r["target_w"], r["target_h"], r["maximize"])
            self._refresh_move_list()
            self.log(f"[이동대상] 추가: {r['name']} ({r['exe_name']})")

    def _edit_move_target(self):
        target = self._get_selected_move_target()
        if not target:
            return
        dlg = MoveTargetEditDialog(self.root, target)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_move_target(r["id"], r["name"], r["exe_name"], r["enabled"], r["move_mode"],
                                  r["target_x"], r["target_y"], r["target_w"], r["target_h"], r["maximize"])
            self._refresh_move_list()
            self.log(f"[이동대상] 수정: {r['name']} ({r['exe_name']})")

    def _delete_move_target(self):
        target = self._get_selected_move_target()
        if not target:
            return
        if messagebox.askyesno("삭제 확인", f"'{target['name']}' 이동 대상을 삭제하시겠습니까?", parent=self.root):
            db_delete_move_target(target["id"])
            self._refresh_move_list()
            self.log(f"[이동대상] 삭제: {target['name']}")

    def _reorder_move_target(self, direction):
        sel = self.move_tree.selection()
        if not sel:
            return
        target_id = int(sel[0])
        idx = next((i for i, t in enumerate(self.move_data) if t["id"] == target_id), None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.move_data):
            return
        db_swap_sort_order("move_targets", self.move_data[idx]["id"], self.move_data[new_idx]["id"])
        self._refresh_move_list()
        new_iid = str(target_id)
        if self.move_tree.exists(new_iid):
            self.move_tree.selection_set(new_iid)
            self.move_tree.see(new_iid)

    def _on_move_double_click(self, event):
        """더블클릭: 해당 프로세스 창을 포그라운드로 활성화"""
        item = self.move_tree.identify_row(event.y)
        if not item:
            return
        target_id = int(item)
        target = next((t for t in self.move_data if t["id"] == target_id), None)
        if not target:
            return
        exe_name = target["exe_name"]
        if self._activate_window_by_exe(exe_name):
            self.log(f"[이동대상] {target['name']} 창 활성화")
        else:
            self.log(f"[이동대상] {target['name']} ({exe_name}) 실행중인 창 없음")

    def _manual_move_target(self):
        """선택한 이동 대상의 창을 지정 위치로 즉시 이동"""
        target = self._get_selected_move_target()
        if not target:
            return
        exe_name = target["exe_name"].lower()
        move_mode = target.get("move_mode", "sub_monitor")
        maximize = target.get("maximize", 1)

        def _do_move():
            hwnds = self._find_windows_by_exe(exe_name)
            if not hwnds:
                self.log(f"[이동대상] {target['name']} ({exe_name}) 실행중인 창 없음")
                return

            tx = ty = tw = th = 0
            if move_mode == "sub_monitor":
                monitors = self._get_monitors_info()
                sub = None
                for m in monitors:
                    if not m[4]:
                        sub = m
                        break
                if not sub:
                    self.log("[이동대상] 서브 모니터를 찾을 수 없음")
                    return
                tx, ty, tw, th = sub[:4]
            elif move_mode in ("custom", "save_position"):
                tx = target.get("target_x", 0)
                ty = target.get("target_y", 0)
                tw = target.get("target_w", 0)
                th = target.get("target_h", 0)

            moved = 0
            for hwnd in hwnds:
                try:
                    if ctypes.windll.user32.IsIconic(hwnd):
                        continue  # 최소화 상태면 건너뜀
                    if move_mode == "sub_monitor":
                        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                        time.sleep(0.1)
                        ctypes.windll.user32.SetWindowPos(
                            hwnd, 0, tx + 100, ty + 100, tw - 200, th - 200, 0x0004
                        )
                        if maximize:
                            time.sleep(0.1)
                            ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                    else:  # custom, save_position
                        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                        time.sleep(0.1)
                        if tw > 0 and th > 0:
                            ctypes.windll.user32.SetWindowPos(
                                hwnd, 0, tx, ty, tw, th, 0x0004
                            )
                        else:
                            rect = ctypes.wintypes.RECT()
                            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                            cw = rect.right - rect.left
                            ch = rect.bottom - rect.top
                            ctypes.windll.user32.SetWindowPos(
                                hwnd, 0, tx, ty, cw, ch, 0x0004
                            )
                    moved += 1
                except Exception:
                    pass

            mode_str = self._MOVE_MODE_DISPLAY.get(move_mode, move_mode)
            self.log(f"[이동대상] {target['name']} {moved}개 창 → {mode_str} 이동 완료")

        threading.Thread(target=_do_move, daemon=True).start()

    def _move_explorer_window(self, folder_name, tx, ty, tw, th):
        """폴더명으로 탐색기 창을 찾아 지정 위치로 이동"""
        target_hwnd = _find_explorer_window_by_title(folder_name)

        if target_hwnd:
            ctypes.windll.user32.ShowWindow(target_hwnd, 9)  # SW_RESTORE
            time.sleep(0.1)
            if tw > 0 and th > 0:
                ctypes.windll.user32.SetWindowPos(target_hwnd, 0, tx, ty, tw, th, 0x0004)
            else:
                rect = ctypes.wintypes.RECT()
                ctypes.windll.user32.GetWindowRect(target_hwnd, ctypes.byref(rect))
                cw = rect.right - rect.left
                ch = rect.bottom - rect.top
                ctypes.windll.user32.SetWindowPos(target_hwnd, 0, tx, ty, cw, ch, 0x0004)
            self.log(f"[폴더] {folder_name} 창 위치 이동: ({tx},{ty} {tw}x{th})")
        else:
            self.log(f"[폴더] {folder_name} 창을 찾지 못함")

    def _wait_and_move_exe_window(self, pid, executable, tx, ty, tw, th):
        """프로세스 창이 나타날 때까지 대기한 뒤 위치 이동"""
        exe_name = os.path.basename(executable).lower()
        user32 = ctypes.windll.user32
        SWP_NOZORDER = 0x0004

        for _ in range(20):  # 최대 10초 대기 (0.5초 × 20)
            time.sleep(0.5)
            hwnd = _find_window_by_exe_name(exe_name)
            if hwnd:
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                time.sleep(0.1)
                if tw > 0 and th > 0:
                    user32.SetWindowPos(hwnd, 0, tx, ty, tw, th, SWP_NOZORDER)
                else:
                    rect = ctypes.wintypes.RECT()
                    user32.GetWindowRect(hwnd, ctypes.byref(rect))
                    user32.SetWindowPos(hwnd, 0, tx, ty, rect.right - rect.left, rect.bottom - rect.top, SWP_NOZORDER)
                self.log(f"[자동실행] {os.path.basename(executable)} 창 위치 이동: ({tx},{ty} {tw}x{th})")
                return
        self.log(f"[자동실행] {os.path.basename(executable)} 창을 찾지 못함 (위치 이동 실패)")

    @staticmethod
    def _find_pids_by_script(script_name_lower):
        """Python 스크립트를 실행중인 프로세스 PID set 반환 (Win32 API 직접 호출)"""
        pids = set()
        try:
            # python/pythonw 프로세스를 tasklist로 빠르게 찾기
            result = subprocess.run(
                ["tasklist", "/fi", "imagename eq pythonw.exe", "/fo", "csv", "/nh"],
                capture_output=True, text=True, timeout=5,
                creationflags=0x08000000,
            )
            result2 = subprocess.run(
                ["tasklist", "/fi", "imagename eq python.exe", "/fo", "csv", "/nh"],
                capture_output=True, text=True, timeout=5,
                creationflags=0x08000000,
            )
            candidate_pids = set()
            for line in (result.stdout + result2.stdout).splitlines():
                parts = line.strip().strip('"').split('","')
                if len(parts) >= 2:
                    try:
                        candidate_pids.add(int(parts[1]))
                    except ValueError:
                        pass
            # 각 PID의 커맨드라인을 Win32 API로 확인
            for pid in candidate_pids:
                cmdline = _get_process_cmdline(pid)
                if cmdline and script_name_lower in cmdline.lower():
                    pids.add(pid)
        except Exception:
            pass
        return pids

    def _find_windows_by_exe(self, exe_name_lower):
        """특정 프로세스명의 창 핸들 목록 반환 (.py/.pyw는 커맨드라인으로 매칭)"""
        hwnds = []
        is_python_script = exe_name_lower.endswith((".py", ".pyw"))
        script_pids = self._find_pids_by_script(exe_name_lower) if is_python_script else set()

        def enum_callback(hwnd, lParam):
            if not ctypes.windll.user32.IsWindowVisible(hwnd):
                return True
            if ctypes.windll.user32.GetWindowTextLengthW(hwnd) == 0:
                return True
            pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if is_python_script and pid.value in script_pids:
                hwnds.append(hwnd)
                return True
            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)
            if handle:
                try:
                    buf = ctypes.create_unicode_buffer(260)
                    size = ctypes.wintypes.DWORD(260)
                    if ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                        if os.path.basename(buf.value).lower() == exe_name_lower:
                            hwnds.append(hwnd)
                finally:
                    ctypes.windll.kernel32.CloseHandle(handle)
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)

        # OpenProcess 실패 대비: 프로세스 목록에서 PID를 찾아 창 매칭
        if not hwnds and not is_python_script:
            try:
                result = subprocess.run(
                    ["tasklist", "/fi", f"imagename eq {exe_name_lower}", "/fo", "csv", "/nh"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=0x08000000,
                )
                fallback_pids = set()
                for line in result.stdout.splitlines():
                    parts = line.strip().strip('"').split('","')
                    if len(parts) >= 2:
                        try:
                            fallback_pids.add(int(parts[1]))
                        except ValueError:
                            pass
                if fallback_pids:
                    def enum_fallback(hwnd, lParam):
                        if not ctypes.windll.user32.IsWindowVisible(hwnd):
                            return True
                        if ctypes.windll.user32.GetWindowTextLengthW(hwnd) == 0:
                            return True
                        w_pid = ctypes.wintypes.DWORD()
                        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(w_pid))
                        if w_pid.value in fallback_pids:
                            hwnds.append(hwnd)
                        return True
                    ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_fallback), 0)
            except Exception:
                pass

        return hwnds

    # ─── GitHub 다운로드 ──────────────────────────────────
    def _on_git_paste(self, event):
        """붙여넣기 후 URL이면 자동 다운로드 시작"""
        self.root.after(50, self._git_auto_download_on_paste)

    def _git_auto_download_on_paste(self):
        """붙여넣기된 텍스트가 GitHub URL이면 자동 다운로드"""
        url = self.git_url_var.get().strip()
        if url and "github.com/" in url:
            if not self._is_valid_git_url(url):
                self.git_url_var.set("")
                self.log("[GitHub] 잘못된 붙여넣기 감지 (URL이 아닌 텍스트)")
                return
            self._git_download()

    @staticmethod
    def _is_valid_git_url(text: str) -> bool:
        """GitHub URL 유효성 검증 (여러 줄이거나 URL 형식이 아니면 False)"""
        if "\n" in text or "\r" in text:
            return False
        import re
        return bool(re.match(r'^https?://github\.com/[\w.\-]+/[\w.\-]+/?$', text))

    def _git_download(self):
        """GitHub 저장소 다운로드 (gitclone.py 호출)"""
        url = self.git_url_var.get().strip()
        if not url:
            return
        if not self._is_valid_git_url(url):
            self.log("[GitHub] 잘못된 URL 형식입니다")
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
        # 출력에서 저장 경로 추출 (##CLONE_PATH: 마커 사용)
        saved_path = ""
        for line in output.splitlines():
            if line.startswith("##CLONE_PATH:"):
                saved_path = line[len("##CLONE_PATH:"):].strip()
        if success:
            self.log(f"[GitHub] 다운로드 성공: {saved_path}" if saved_path else "[GitHub] 다운로드 성공")
            self.git_url_var.set("")
            if saved_path and self.var_git_open_folder.get() and os.path.isdir(saved_path):
                subprocess.Popen(["explorer.exe", os.path.normpath(saved_path)])
        else:
            if saved_path:
                self.log(f"[GitHub] 이미 존재하는 경로: {saved_path}")
            else:
                self.log(f"[GitHub] 다운로드 실패: {output[:200]}")
                messagebox.showerror("GitHub 다운로드 실패", output[:500], parent=self.root)
        self.git_url_entry.focus_set()

    def _run_task(self):
        """선택한 자동실행 작업을 즉시 테스트 실행 (폴더면 열기, 실행중이면 활성화)"""
        task = self._get_selected_task()
        if not task:
            return
        if os.path.isdir(task["executable"]):
            self._open_folder_task(task)
            return
        # 이미 실행중인 창이 있으면 활성화
        task_id = task["id"]
        if task_id in self._running_tasks:
            proc = self._task_processes.get(task_id)
            if proc and proc.poll() is None and self._activate_window_by_pid(proc.pid):
                self.log(f"[자동실행] {task['name']} 창 활성화")
                return
        if self._activate_window_by_exe(os.path.basename(task["executable"])):
            self.log(f"[자동실행] {task['name']} 창 활성화")
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

        # 듀얼 모니터 감지 (3초마다)
        self._check_monitor_change()

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

    def _run_boot_tasks(self):
        """Windows 부팅 후 1회 실행 태스크 처리"""
        now = datetime.now()
        # Windows 부팅 시각 계산
        uptime_ms = ctypes.windll.kernel32.GetTickCount64()
        boot_time = now - timedelta(milliseconds=uptime_ms)
        now_str = now.strftime("%Y-%m-%d %H:%M:%S")

        # 부팅 후 10분 이상 경과하면 실행하지 않음
        uptime_min = uptime_ms / 60000
        if uptime_min > 10:
            self.log(f"[부팅] 부팅 후 {int(uptime_min)}분 경과 - 부팅 태스크 건너뜀")
            return

        boot_count = 0
        for task in self.task_data:
            if not task["enabled"]:
                continue
            if task.get("repeat_mode") != "boot":
                continue
            # 마지막 실행이 이번 부팅 이후였으면 건너뜀
            last_run = str(task.get("last_run", "") or "")
            if last_run:
                try:
                    last_dt = datetime.strptime(last_run[:19], "%Y-%m-%d %H:%M:%S")
                    if last_dt > boot_time:
                        continue
                except ValueError:
                    try:
                        last_dt = datetime.strptime(last_run[:10], "%Y-%m-%d")
                        if last_dt.date() == now.date():
                            continue  # 날짜만 있는 레거시 데이터: 오늘이면 건너뜀
                    except ValueError:
                        pass
            if task["skip_holiday"] and self._is_closed_day(now):
                continue
            task["last_run"] = now_str
            db_update_task_last_run(task["id"], now_str)
            self._execute_task(task, now_str)
            boot_count += 1
        if boot_count:
            self.log(f"[부팅] {boot_count}개 태스크 실행 (부팅 시각: {boot_time.strftime('%H:%M:%S')})")

    def _check_auto_tasks(self, current_hm, today_str, is_closed):
        """자동실행 체크 (매일 1회 + 반복 모드 지원)"""
        now = datetime.now()
        for task in self.task_data:
            if not task["enabled"]:
                continue
            if task["skip_holiday"] and is_closed:
                continue

            repeat_mode = task.get("repeat_mode", "once")

            if repeat_mode == "boot":
                continue  # boot 모드는 앱 시작 시 _run_boot_tasks에서 처리

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
        """자동실행 작업 실행 (별도 프로세스, 폴더는 탐색기로 열기)"""
        task_id = task["id"]
        executable = task["executable"]

        # 폴더인 경우 탐색기로 열기
        if os.path.isdir(executable):
            self._open_folder_task(task)
            db_update_task_last_run(task_id, today_str)
            return

        # 중복 실행 방지: 이미 실행중인 작업이면 무시
        if task_id in self._running_tasks:
            self.log(f"[자동실행] {task['name']} 이미 실행중 - 건너뜀")
            return

        self._running_tasks.add(task_id)

        arguments = task.get("arguments", "")
        python_venv = task["python_venv"]

        def _run():
            try:
                ext = os.path.splitext(executable)[1].lower()
                if ext in (".py", ".pyw") and python_venv:
                    cmd = [python_venv, executable]
                elif ext == ".pyw":
                    python_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                    cmd = [python_exe, executable]
                elif ext == ".py":
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
                # auto_move가 켜져 있으면 창이 생길 때까지 대기 후 위치 이동
                if task.get("auto_move", 0):
                    tx = task.get("target_x", 0) or 0
                    ty = task.get("target_y", 0) or 0
                    tw = task.get("target_w", 0) or 0
                    th = task.get("target_h", 0) or 0
                    if tx or ty or tw or th:
                        self._wait_and_move_exe_window(proc.pid, executable, tx, ty, tw, th)
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

    # ─── 듀얼 모니터 감지 → 브라우저 이동 ─────────────────
    def _check_monitor_change(self):
        """모니터 수 변화 감지 (3초 주기)"""
        self._monitor_check_counter += 1
        if self._monitor_check_counter < 3:
            return
        self._monitor_check_counter = 0

        count = ctypes.windll.user32.GetSystemMetrics(80)  # SM_CMONITORS
        prev = self._last_monitor_count
        self._last_monitor_count = count

        if prev <= 1 and count >= 2:
            self.log("[모니터] 듀얼 모니터 감지 → 브라우저를 서브 모니터로 이동")
            threading.Thread(target=self._move_browsers_to_sub_monitor, daemon=True).start()

    def _get_monitors_info(self):
        """모니터 정보 반환: [(x, y, w, h, is_primary), ...]"""
        monitors = []

        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.wintypes.DWORD),
                ("rcMonitor", ctypes.wintypes.RECT),
                ("rcWork", ctypes.wintypes.RECT),
                ("dwFlags", ctypes.wintypes.DWORD),
            ]

        def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
            info = MONITORINFO()
            info.cbSize = ctypes.sizeof(MONITORINFO)
            ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(info))
            rc = info.rcMonitor
            is_primary = bool(info.dwFlags & 0x01)  # MONITORINFOF_PRIMARY
            monitors.append((rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, is_primary))
            return True

        MONITORENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool,
            ctypes.wintypes.HMONITOR,
            ctypes.wintypes.HDC,
            ctypes.POINTER(ctypes.wintypes.RECT),
            ctypes.wintypes.LPARAM,
        )
        ctypes.windll.user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0)
        return monitors

    def _find_movable_windows(self):
        """DB에 등록된 이동 대상(사용 중)의 창 핸들 목록 반환"""
        targets = db_fetch_move_targets()
        if not targets:
            return []
        target_exes = {t["exe_name"].lower() for t in targets if t.get("enabled", 1)}
        hwnds = []

        def enum_callback(hwnd, lParam):
            if not ctypes.windll.user32.IsWindowVisible(hwnd):
                return True
            if ctypes.windll.user32.GetWindowTextLengthW(hwnd) == 0:
                return True
            pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)
            if handle:
                try:
                    buf = ctypes.create_unicode_buffer(260)
                    size = ctypes.wintypes.DWORD(260)
                    if ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                        exe_name = os.path.basename(buf.value).lower()
                        if exe_name in target_exes:
                            hwnds.append(hwnd)
                finally:
                    ctypes.windll.kernel32.CloseHandle(handle)
            return True

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
        return hwnds

    def _move_browsers_to_sub_monitor(self):
        """DB 등록 대상을 각자 설정된 위치로 이동"""
        try:
            time.sleep(2)  # 모니터 인식 안정화 대기

            monitors = self._get_monitors_info()
            sub = None
            for m in monitors:
                if not m[4]:
                    sub = m
                    break

            targets = db_fetch_move_targets()
            if not targets:
                self.log("[모니터] 등록된 이동 대상 없음")
                return

            moved = 0
            for target in targets:
                if not target.get("enabled", 1):
                    continue
                exe_name = target["exe_name"].lower()
                move_mode = target.get("move_mode", "sub_monitor")
                maximize = target.get("maximize", 1)
                hwnds = self._find_windows_by_exe(exe_name)
                if not hwnds:
                    continue

                tx = ty = tw = th = 0
                if move_mode == "sub_monitor":
                    if not sub:
                        continue
                    tx, ty, tw, th = sub[:4]
                elif move_mode in ("custom", "save_position"):
                    tx = target.get("target_x", 0)
                    ty = target.get("target_y", 0)
                    tw = target.get("target_w", 0)
                    th = target.get("target_h", 0)

                for hwnd in hwnds:
                    try:
                        if ctypes.windll.user32.IsIconic(hwnd):
                            continue  # 최소화 상태면 건너뜀
                        if move_mode == "sub_monitor":
                            ctypes.windll.user32.ShowWindow(hwnd, 9)
                            time.sleep(0.1)
                            ctypes.windll.user32.SetWindowPos(
                                hwnd, 0, tx + 100, ty + 100, tw - 200, th - 200, 0x0004
                            )
                            if maximize:
                                time.sleep(0.1)
                                ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                        else:  # custom, save_position
                            ctypes.windll.user32.ShowWindow(hwnd, 9)
                            time.sleep(0.1)
                            if tw > 0 and th > 0:
                                ctypes.windll.user32.SetWindowPos(
                                    hwnd, 0, tx, ty, tw, th, 0x0004
                                )
                            else:
                                rect = ctypes.wintypes.RECT()
                                ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                                cw = rect.right - rect.left
                                ch = rect.bottom - rect.top
                                ctypes.windll.user32.SetWindowPos(
                                    hwnd, 0, tx, ty, cw, ch, 0x0004
                                )
                        moved += 1
                    except Exception:
                        pass

            self.log(f"[모니터] {moved}개 창 이동 완료")
        except Exception as e:
            self.log(f"[모니터] 창 이동 실패: {e}")

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
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "AutoExec_Python")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        ctypes.windll.kernel32.CloseHandle(_mutex)
        # 이미 실행중인 AutoExec 창을 찾아 활성화
        EnumWindowsProc = ctypes.WINFUNCTYPE(
            ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM
        )
        user32 = ctypes.windll.user32
        found = [None]

        def _enum_cb(hwnd, _lp):
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            if buf.value.startswith("실행 관리 서버"):
                found[0] = hwnd
                return False  # 찾았으므로 열거 중단
            return True

        user32.EnumWindows(EnumWindowsProc(_enum_cb), 0)
        if found[0]:
            SW_RESTORE = 9
            if user32.IsIconic(found[0]):
                user32.ShowWindow(found[0], SW_RESTORE)
            user32.SetForegroundWindow(found[0])
        sys.exit(0)

    db_init()
    app = AutoExecApp()
    app.run()
    ctypes.windll.kernel32.CloseHandle(_mutex)
