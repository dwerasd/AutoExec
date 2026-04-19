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
    # ── 창 배치 프로파일 시스템 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS window_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL DEFAULT '',
            exe_name TEXT NOT NULL DEFAULT '',
            enabled INTEGER NOT NULL DEFAULT 1,
            trigger_mode TEXT NOT NULL DEFAULT 'monitor_change',
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS window_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            profile_id INTEGER NOT NULL,
            window_class TEXT NOT NULL DEFAULT '',
            title_pattern TEXT NOT NULL DEFAULT '',
            move_mode TEXT NOT NULL DEFAULT 'custom',
            target_x INTEGER NOT NULL DEFAULT 0,
            target_y INTEGER NOT NULL DEFAULT 0,
            target_w INTEGER NOT NULL DEFAULT 0,
            target_h INTEGER NOT NULL DEFAULT 0,
            maximize INTEGER NOT NULL DEFAULT 0,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    # ── 과제 (루틴) 테이블 ──
    cur.execute("""
        CREATE TABLE IF NOT EXISTS routines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL DEFAULT '',
            daily_count INTEGER NOT NULL DEFAULT 1,
            enabled INTEGER NOT NULL DEFAULT 1,
            start_date TEXT NOT NULL DEFAULT '',
            repeat_type TEXT NOT NULL DEFAULT 'once',
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)
    try:
        cur.execute("ALTER TABLE routines ADD COLUMN start_date TEXT NOT NULL DEFAULT ''")
    except Exception:
        pass
    # repeat_type: 'once'(1회) / 'daily'(매일)
    try:
        cur.execute("ALTER TABLE routines ADD COLUMN repeat_type TEXT NOT NULL DEFAULT 'once'")
    except Exception:
        pass
    cur.execute("""
        CREATE TABLE IF NOT EXISTS routine_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            routine_id INTEGER NOT NULL,
            log_date TEXT NOT NULL,
            seq INTEGER NOT NULL,
            done_time TEXT NOT NULL DEFAULT '',
            UNIQUE(routine_id, log_date, seq)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS routine_hidden (
            routine_id INTEGER NOT NULL,
            hidden_date TEXT NOT NULL,
            UNIQUE(routine_id, hidden_date)
        )
    """)
    # move_targets → window_profiles + window_rules 마이그레이션
    try:
        cur.execute("SELECT COUNT(*) as cnt FROM window_profiles")
        profile_count = cur.fetchone()["cnt"]
        if profile_count == 0:
            cur.execute("SELECT * FROM move_targets ORDER BY sort_order, id")
            old_targets = cur.fetchall()
            for ot in old_targets:
                cur.execute(
                    "INSERT INTO window_profiles (name, exe_name, enabled, trigger_mode, sort_order) "
                    "VALUES (?, ?, ?, 'monitor_change', ?)",
                    (ot["name"], ot["exe_name"], ot.get("enabled", 1), ot.get("sort_order", 0)),
                )
                new_profile_id = cur.lastrowid
                cur.execute(
                    "INSERT INTO window_rules (profile_id, window_class, title_pattern, move_mode, "
                    "target_x, target_y, target_w, target_h, maximize, sort_order) "
                    "VALUES (?, '', '', ?, ?, ?, ?, ?, ?, 0)",
                    (new_profile_id, ot.get("move_mode", "sub_monitor"),
                     ot.get("target_x", 0), ot.get("target_y", 0),
                     ot.get("target_w", 0), ot.get("target_h", 0),
                     ot.get("maximize", 1)),
                )
    except Exception:
        pass
    # routine_logs의 구형 done_time(HH:MM:SS) → 풀 datetime(YYYY-MM-DD HH:MM:SS) 마이그레이션
    try:
        cur.execute("SELECT id, log_date, done_time FROM routine_logs WHERE LENGTH(done_time) <= 8 AND done_time != ''")
        old_rows = cur.fetchall()
        for r in old_rows:
            new_time = f"{r['log_date']} {r['done_time']}"
            cur.execute("UPDATE routine_logs SET done_time=? WHERE id=?", (new_time, r["id"]))
    except Exception:
        pass
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
    if table not in ("pcs", "tasks", "move_targets", "window_profiles", "window_rules", "routines"):
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


def parse_done_datetime(log_date, done_time):
    """done_time을 datetime으로 변환. 풀 datetime 또는 시간만 있는 구형 포맷 모두 지원."""
    if not done_time:
        return None
    try:
        if len(done_time) > 8:
            # 풀 datetime: "YYYY-MM-DD HH:MM:SS"
            return datetime.strptime(done_time, "%Y-%m-%d %H:%M:%S")
        # 구형: 시간만 "HH:MM:SS" → log_date와 조합
        return datetime.strptime(f"{log_date} {done_time}", "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def format_done_time_display(done_time):
    """done_time을 표시용 문자열로 변환 (풀 datetime 그대로 표시)"""
    if not done_time:
        return ""
    return done_time


def db_get_prev_routine_done_time(routine_id, before_date):
    """해당 날짜 이전의 가장 최근 완료 시각 반환 (datetime 또는 None)"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT log_date, done_time FROM routine_logs "
            "WHERE routine_id=? AND log_date<? ORDER BY log_date DESC, seq DESC LIMIT 1",
            (routine_id, before_date),
        )
        row = cur.fetchone()
        if row and row["done_time"]:
            return parse_done_datetime(row["log_date"], row["done_time"])
        return None
    finally:
        conn.close()


def format_elapsed(seconds):
    """초 단위를 사람이 읽기 좋은 경과 문자열로 변환"""
    if seconds < 0:
        return ""
    minutes = int(seconds // 60)
    hours = minutes // 60
    if hours > 0:
        return f"{hours}시간 {minutes % 60}분"
    return f"{minutes}분"



# ═══════════════════════════════════════════════════════════
#  과제 (루틴) DB 헬퍼
# ═══════════════════════════════════════════════════════════
def db_fetch_routines():
    """과제 목록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM routines ORDER BY sort_order, id")
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_routine(routine_id, name, daily_count, enabled=1, start_date="", repeat_type="once"):
    """과제 추가 또는 수정"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if routine_id:
            cur.execute(
                "UPDATE routines SET name=?, daily_count=?, enabled=?, start_date=?, repeat_type=? WHERE id=?",
                (name, daily_count, int(enabled), start_date, repeat_type, routine_id),
            )
        else:
            cur.execute(
                "INSERT INTO routines (name, daily_count, enabled, start_date, repeat_type, sort_order) "
                "VALUES (?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM routines))",
                (name, daily_count, int(enabled), start_date, repeat_type),
            )
        conn.commit()
    finally:
        conn.close()


def db_delete_routine(routine_id):
    """과제 비활성화 (과거 기록 보존, 일정만 중단)"""
    conn = get_db_connection()
    try:
        conn.execute("UPDATE routines SET enabled=0 WHERE id=?", (routine_id,))
        conn.commit()
    finally:
        conn.close()


def db_hide_routine_date(routine_id, date_str):
    """일정 날짜 숨김 저장"""
    conn = get_db_connection()
    try:
        conn.execute(
            "INSERT OR IGNORE INTO routine_hidden (routine_id, hidden_date) VALUES (?, ?)",
            (routine_id, date_str),
        )
        conn.commit()
    finally:
        conn.close()


def db_fetch_hidden_routine_dates():
    """숨김 처리된 (routine_id, date) 세트 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT routine_id, hidden_date FROM routine_hidden")
        return {(row["routine_id"], row["hidden_date"]) for row in cur.fetchall()}
    finally:
        conn.close()


def db_fetch_routine_logs(routine_id, log_date):
    """특정 과제의 특정 날짜 완료 기록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM routine_logs WHERE routine_id=? AND log_date=? ORDER BY seq",
            (routine_id, log_date),
        )
        return cur.fetchall()
    finally:
        conn.close()


def db_add_routine_log(routine_id, log_date, seq):
    """과제 완료 기록 추가"""
    conn = get_db_connection()
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT OR IGNORE INTO routine_logs (routine_id, log_date, seq, done_time) "
            "VALUES (?, ?, ?, ?)",
            (routine_id, log_date, seq, now_str),
        )
        conn.commit()
    finally:
        conn.close()


def db_remove_last_routine_log(routine_id, log_date):
    """특정 과제의 특정 날짜 마지막 완료 기록 삭제"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id FROM routine_logs WHERE routine_id=? AND log_date=? ORDER BY seq DESC LIMIT 1",
            (routine_id, log_date),
        )
        row = cur.fetchone()
        if row:
            conn.execute("DELETE FROM routine_logs WHERE id=?", (row["id"],))
            conn.commit()
            return True
        return False
    finally:
        conn.close()


def db_get_routine_display_dates(routine_id, daily_count, start_date, repeat_type):
    """일정의 표시할 날짜 목록 반환. [(date_str, done_count, is_past), ...]
    매일 반복: start_date~오늘 중 미완료 날짜(최대 7일) + 오늘.
    1회: 해당 날짜만."""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        today = date.today()
        today_str = today.strftime("%Y-%m-%d")

        if repeat_type == "once":
            target = start_date if start_date else today_str
            cur.execute(
                "SELECT COUNT(*) as cnt FROM routine_logs WHERE routine_id=? AND log_date=?",
                (routine_id, target),
            )
            return [(target, cur.fetchone()["cnt"], target < today_str)]

        # 매일 반복 — 시작일 결정
        if start_date:
            start = datetime.strptime(start_date, "%Y-%m-%d").date()
        else:
            cur.execute(
                "SELECT MIN(log_date) as first_date FROM routine_logs WHERE routine_id=?",
                (routine_id,),
            )
            row = cur.fetchone()
            start = datetime.strptime(row["first_date"], "%Y-%m-%d").date() if row["first_date"] else today

        # 최대 7일 전까지만
        week_ago = today - timedelta(days=7)
        start = max(start, week_ago)

        results = []
        now = datetime.now()
        current = start
        while current <= today:
            d_str = current.strftime("%Y-%m-%d")
            cur.execute(
                "SELECT COUNT(*) as cnt FROM routine_logs WHERE routine_id=? AND log_date=?",
                (routine_id, d_str),
            )
            done = cur.fetchone()["cnt"]
            if current == today:
                # 오늘은 항상 표시
                results.append((d_str, done, False))
            elif done < daily_count:
                # 과거 미완료 → 표시
                results.append((d_str, done, True))
            else:
                # 과거 완료 → 마지막 완료 시각이 24시간 이내면 표시
                cur.execute(
                    "SELECT done_time FROM routine_logs "
                    "WHERE routine_id=? AND log_date=? ORDER BY seq DESC LIMIT 1",
                    (routine_id, d_str),
                )
                row = cur.fetchone()
                if row and row["done_time"]:
                    done_dt = parse_done_datetime(d_str, row["done_time"])
                    if done_dt and (now - done_dt).total_seconds() < 86400:
                        results.append((d_str, done, True))
            current += timedelta(days=1)

        return results
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
#  창 배치 프로파일 (window_profiles + window_rules 테이블)
# ═══════════════════════════════════════════════════════════
def db_fetch_profiles():
    """창 배치 프로파일 목록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM window_profiles ORDER BY sort_order, id")
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_profile(profile_id, name, exe_name, enabled=1, trigger_mode="monitor_change"):
    """프로파일 추가 또는 수정. 반환: id"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if profile_id:
            cur.execute(
                "UPDATE window_profiles SET name=?, exe_name=?, enabled=?, trigger_mode=? WHERE id=?",
                (name, exe_name, int(enabled), trigger_mode, profile_id),
            )
        else:
            cur.execute(
                "INSERT INTO window_profiles (name, exe_name, enabled, trigger_mode, sort_order) "
                "VALUES (?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM window_profiles))",
                (name, exe_name, int(enabled), trigger_mode),
            )
            profile_id = cur.lastrowid
        conn.commit()
        return profile_id
    finally:
        conn.close()


def db_delete_profile(profile_id):
    """프로파일 + 소속 규칙 삭제"""
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM window_rules WHERE profile_id=?", (profile_id,))
        conn.execute("DELETE FROM window_profiles WHERE id=?", (profile_id,))
        conn.commit()
    finally:
        conn.close()


def db_fetch_rules(profile_id):
    """특정 프로파일의 규칙 목록 조회"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM window_rules WHERE profile_id=? ORDER BY sort_order, id", (profile_id,))
        return cur.fetchall()
    finally:
        conn.close()


def db_upsert_rule(rule_id, profile_id, window_class="", title_pattern="",
                   move_mode="custom", target_x=0, target_y=0, target_w=0, target_h=0, maximize=0):
    """규칙 추가 또는 수정. 반환: id"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        if rule_id:
            cur.execute(
                "UPDATE window_rules SET profile_id=?, window_class=?, title_pattern=?, move_mode=?, "
                "target_x=?, target_y=?, target_w=?, target_h=?, maximize=? WHERE id=?",
                (profile_id, window_class, title_pattern, move_mode,
                 target_x, target_y, target_w, target_h, int(maximize), rule_id),
            )
        else:
            cur.execute(
                "INSERT INTO window_rules (profile_id, window_class, title_pattern, move_mode, "
                "target_x, target_y, target_w, target_h, maximize, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?, (SELECT IFNULL(MAX(sort_order),0)+1 FROM window_rules WHERE profile_id=?))",
                (profile_id, window_class, title_pattern, move_mode,
                 target_x, target_y, target_w, target_h, int(maximize), profile_id),
            )
            rule_id = cur.lastrowid
        conn.commit()
        return rule_id
    finally:
        conn.close()


def db_delete_rule(rule_id):
    """규칙 삭제"""
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM window_rules WHERE id=?", (rule_id,))
        conn.commit()
    finally:
        conn.close()


def db_replace_rules(profile_id, rules_list):
    """프로파일의 모든 규칙을 교체 (일괄 캡처용)"""
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM window_rules WHERE profile_id=?", (profile_id,))
        cur = conn.cursor()
        for idx, r in enumerate(rules_list):
            cur.execute(
                "INSERT INTO window_rules (profile_id, window_class, title_pattern, move_mode, "
                "target_x, target_y, target_w, target_h, maximize, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?,?)",
                (profile_id, r.get("window_class", ""), r.get("title_pattern", ""),
                 r.get("move_mode", "custom"), r.get("target_x", 0), r.get("target_y", 0),
                 r.get("target_w", 0), r.get("target_h", 0), int(r.get("maximize", 0)), idx),
            )
        conn.commit()
    finally:
        conn.close()


def db_count_rules(profile_id):
    """프로파일의 규칙 수 반환"""
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) as cnt FROM window_rules WHERE profile_id=?", (profile_id,))
        return cur.fetchone()["cnt"]
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
        if pp.value is None:
            return None
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


def _enumerate_process_windows(exe_name_lower):
    """프로세스명(소문자)의 모든 visible 창 정보 반환: [{"hwnd", "class_name", "title", "x", "y", "w", "h"}, ...]"""
    windows = []
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    GetClassNameW = user32.GetClassNameW

    def enum_callback(hwnd, lParam):
        if not user32.IsWindowVisible(hwnd):
            return True
        if user32.GetWindowTextLengthW(hwnd) == 0:
            return True
        pid = ctypes.wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        handle = kernel32.OpenProcess(0x1000, False, pid.value)
        if handle:
            try:
                buf = ctypes.create_unicode_buffer(260)
                size = ctypes.wintypes.DWORD(260)
                if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                    if os.path.basename(buf.value).lower() == exe_name_lower:
                        cls_buf = ctypes.create_unicode_buffer(256)
                        GetClassNameW(hwnd, cls_buf, 256)
                        length = user32.GetWindowTextLengthW(hwnd)
                        title_buf = ctypes.create_unicode_buffer(length + 1)
                        user32.GetWindowTextW(hwnd, title_buf, length + 1)
                        rect = ctypes.wintypes.RECT()
                        user32.GetWindowRect(hwnd, ctypes.byref(rect))
                        windows.append({
                            "hwnd": hwnd,
                            "class_name": cls_buf.value,
                            "title": title_buf.value,
                            "x": rect.left, "y": rect.top,
                            "w": rect.right - rect.left, "h": rect.bottom - rect.top,
                        })
            finally:
                kernel32.CloseHandle(handle)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return windows


def _find_single_window(exe_name_lower, class_filter="", title_filter=""):
    """class+title 조건으로 단일 창 핸들 반환"""
    windows = _enumerate_process_windows(exe_name_lower)
    for w in windows:
        if class_filter and w["class_name"] != class_filter:
            continue
        if title_filter:
            if title_filter.endswith("*"):
                if not w["title"].startswith(title_filter[:-1]):
                    continue
            elif title_filter not in w["title"]:
                continue
        return w["hwnd"]
    return None


def _match_window_to_rules(class_name, title, rules):
    """창의 class/title에 매칭되는 첫 번째 규칙 반환. 없으면 None."""
    for rule in rules:
        rc = rule.get("window_class", "")
        rt = rule.get("title_pattern", "")
        # class 매칭
        if rc and rc != class_name:
            continue
        # title 매칭
        if rt:
            if rt.endswith("*"):
                if not title.startswith(rt[:-1]):
                    continue
            elif rt not in title:
                continue
        return rule
    return None


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
#  과제 편집 다이얼로그
# ═══════════════════════════════════════════════════════════
class RoutineEditDialog(tk.Toplevel):
    def __init__(self, parent, routine=None):
        super().__init__(parent)
        self.result = None
        self.routine = routine
        self.title("일정 편집" if routine else "일정 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="내용:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_name = ttk.Entry(frame, width=25)
        self.ent_name.grid(row=0, column=1, pady=3, padx=(5, 0))

        ttk.Label(frame, text="반복:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.cmb_repeat = ttk.Combobox(frame, values=["1회", "매일"],
                                        width=8, state="readonly")
        self.cmb_repeat.grid(row=1, column=1, sticky=tk.W, pady=3, padx=(5, 0))
        self.cmb_repeat.set("1회")

        ttk.Label(frame, text="시작일:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_start_date = ttk.Entry(frame, width=12)
        self.ent_start_date.grid(row=2, column=1, sticky=tk.W, pady=3, padx=(5, 0))
        self.ent_start_date.insert(0, date.today().strftime("%Y-%m-%d"))

        self.var_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="사용", variable=self.var_enabled).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=3
        )

        if routine:
            self.ent_name.insert(0, routine["name"])
            repeat_label = "1회" if routine.get("repeat_type", "once") == "once" else "매일"
            self.cmb_repeat.set(repeat_label)
            self.ent_start_date.delete(0, tk.END)
            self.ent_start_date.insert(0, routine.get("start_date", ""))
            self.var_enabled.set(bool(routine["enabled"]))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        dw, dh = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
        self.ent_name.focus_set()

    def _on_ok(self):
        name = self.ent_name.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "내용을 입력하세요.", parent=self)
            return
        repeat_type = "once" if self.cmb_repeat.get() == "1회" else "daily"
        self.result = {
            "id": self.routine["id"] if self.routine else None,
            "name": name,
            "daily_count": 1,
            "repeat_type": repeat_type,
            "start_date": self.ent_start_date.get().strip(),
            "enabled": self.var_enabled.get(),
        }
        self.destroy()


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
class RuleEditDialog(tk.Toplevel):
    """창 배치 규칙 편집 다이얼로그"""
    _MODE_LABELS = {"서브 모니터": "sub_monitor", "현재 위치 저장": "save_position", "좌표 직접 입력": "custom"}
    _MODE_KEYS = {"sub_monitor": "서브 모니터", "save_position": "현재 위치 저장", "custom": "좌표 직접 입력"}

    def __init__(self, parent, rule=None, exe_name=""):
        super().__init__(parent)
        self.result = None
        self.rule = rule
        self._exe_name = exe_name
        self.title("규칙 편집" if rule else "규칙 추가")
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Window Class:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_class = ttk.Entry(frame, width=30)
        self.ent_class.grid(row=0, column=1, columnspan=3, pady=3, padx=(5, 0))
        ttk.Label(frame, text="(비어있으면 모두 매칭)", foreground="gray").grid(row=1, column=0, columnspan=4, sticky=tk.W)

        ttk.Label(frame, text="Title 패턴:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.ent_title = ttk.Entry(frame, width=30)
        self.ent_title.grid(row=2, column=1, columnspan=3, pady=3, padx=(5, 0))
        ttk.Label(frame, text="(비어있으면 모두 매칭, * = 와일드카드)", foreground="gray").grid(
            row=3, column=0, columnspan=4, sticky=tk.W)

        ttk.Label(frame, text="이동 위치:").grid(row=4, column=0, sticky=tk.W, pady=3)
        self.var_mode = tk.StringVar(value="좌표 직접 입력")
        self.cmb_mode = ttk.Combobox(frame, textvariable=self.var_mode,
                                     values=list(self._MODE_LABELS.keys()),
                                     state="readonly", width=15)
        self.cmb_mode.grid(row=4, column=1, columnspan=3, sticky=tk.W, padx=(5, 0), pady=3)
        self.cmb_mode.bind("<<ComboboxSelected>>", self._on_mode_changed)

        # 좌표 입력
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

        # 위치 캡처 프레임
        self.save_pos_frame = ttk.Frame(frame)
        self.save_pos_frame.grid(row=6, column=0, columnspan=4, sticky=tk.W, pady=3)
        ttk.Button(self.save_pos_frame, text="현재 위치 캡처", command=self._capture_position).pack(side=tk.LEFT)
        self.lbl_saved_pos = ttk.Label(self.save_pos_frame, text="", foreground="gray")
        self.lbl_saved_pos.pack(side=tk.LEFT, padx=(8, 0))

        self.var_maximize = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="최대화", variable=self.var_maximize).grid(
            row=7, column=0, columnspan=2, sticky=tk.W, pady=3)

        if rule:
            self.ent_class.insert(0, rule.get("window_class", ""))
            self.ent_title.insert(0, rule.get("title_pattern", ""))
            self.var_maximize.set(bool(rule.get("maximize", 0)))
            mode = rule.get("move_mode", "custom")
            self.var_mode.set(self._MODE_KEYS.get(mode, "좌표 직접 입력"))
            if mode in ("custom", "save_position"):
                self.ent_x.delete(0, tk.END)
                self.ent_x.insert(0, str(rule.get("target_x", 0)))
                self.ent_y.delete(0, tk.END)
                self.ent_y.insert(0, str(rule.get("target_y", 0)))
                self.ent_w.delete(0, tk.END)
                self.ent_w.insert(0, str(rule.get("target_w", 0)))
                self.ent_h.delete(0, tk.END)
                self.ent_h.insert(0, str(rule.get("target_h", 0)))

        self._on_mode_changed()

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=8, column=0, columnspan=4, pady=(10, 0))
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
        self.ent_class.focus_set()

    def _on_mode_changed(self, event=None):
        mode = self._MODE_LABELS.get(self.var_mode.get(), "custom")
        self.coord_frame.grid_remove()
        self.save_pos_frame.grid_remove()
        if mode == "custom":
            self.coord_frame.grid()
        elif mode == "save_position":
            self.save_pos_frame.grid()
            self._update_saved_pos_label()
        if event is not None:
            self.var_maximize.set(mode == "sub_monitor")

    def _update_saved_pos_label(self):
        src = getattr(self, "_captured", None) or (self.rule if self.rule else None)
        if src:
            x, y = src.get("target_x", 0), src.get("target_y", 0)
            w, h = src.get("target_w", 0), src.get("target_h", 0)
            if x or y or w or h:
                size = f" {w}x{h}" if w and h else ""
                self.lbl_saved_pos.config(text=f"저장된 위치: {x},{y}{size}")
                return
        self.lbl_saved_pos.config(text="저장된 위치 없음 - 캡처하세요")

    def _capture_position(self):
        """class+title로 매칭되는 창의 현재 위치를 캡처"""
        cls_filter = self.ent_class.get().strip()
        title_filter = self.ent_title.get().strip()
        exe_name = self._exe_name.lower()
        if not exe_name:
            messagebox.showwarning("입력 오류", "프로파일의 프로세스명이 없습니다.", parent=self)
            return

        found_hwnd = _find_single_window(exe_name, cls_filter, title_filter)
        if not found_hwnd:
            messagebox.showinfo("알림", "매칭되는 창을 찾을 수 없습니다.", parent=self)
            return

        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(found_hwnd, ctypes.byref(rect))
        x, y = rect.left, rect.top
        w, h = rect.right - rect.left, rect.bottom - rect.top
        self._captured = {"target_x": x, "target_y": y, "target_w": w, "target_h": h}
        self.ent_x.delete(0, tk.END)
        self.ent_x.insert(0, str(x))
        self.ent_y.delete(0, tk.END)
        self.ent_y.insert(0, str(y))
        self.ent_w.delete(0, tk.END)
        self.ent_w.insert(0, str(w))
        self.ent_h.delete(0, tk.END)
        self.ent_h.insert(0, str(h))
        self.lbl_saved_pos.config(text=f"캡처 완료: {x},{y} {w}x{h}")

    def _on_ok(self):
        mode = self._MODE_LABELS.get(self.var_mode.get(), "custom")
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
            src = getattr(self, "_captured", None) or (self.rule if self.rule else {})
            target_x = src.get("target_x", 0)
            target_y = src.get("target_y", 0)
            target_w = src.get("target_w", 0)
            target_h = src.get("target_h", 0)
        self.result = {
            "id": self.rule.get("id") if self.rule else None,
            "window_class": self.ent_class.get().strip(),
            "title_pattern": self.ent_title.get().strip(),
            "move_mode": mode,
            "target_x": target_x, "target_y": target_y,
            "target_w": target_w, "target_h": target_h,
            "maximize": self.var_maximize.get(),
        }
        self.destroy()


class ProfileEditDialog(tk.Toplevel):
    """창 배치 프로파일 편집 다이얼로그"""
    _TRIGGER_LABELS = {"듀얼모니터 감지": "monitor_change", "프로세스 감지": "process_start", "둘 다": "both"}
    _TRIGGER_KEYS = {"monitor_change": "듀얼모니터 감지", "process_start": "프로세스 감지", "both": "둘 다"}
    _MODE_DISPLAY = {"sub_monitor": "서브모니터", "save_position": "저장위치", "custom": "좌표직접"}

    def __init__(self, parent, profile=None):
        super().__init__(parent)
        self.result = None
        self.profile = profile
        self.rules_data = []  # 현재 편집중인 규칙 목록
        self.title("프로파일 편집" if profile else "프로파일 추가")
        self.resizable(True, True)
        self.grab_set()
        self.minsize(580, 400)

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        # ── 상단: 프로파일 기본 정보 ──
        ttk.Label(frame, text="이름:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.ent_name = ttk.Entry(frame, width=30)
        self.ent_name.grid(row=0, column=1, columnspan=3, sticky=tk.EW, pady=3, padx=(5, 0))

        ttk.Label(frame, text="프로세스명:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.ent_exe = ttk.Entry(frame, width=30)
        self.ent_exe.grid(row=1, column=1, columnspan=3, sticky=tk.EW, pady=3, padx=(5, 0))

        info_frame = ttk.Frame(frame)
        info_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W)
        self.var_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(info_frame, text="사용", variable=self.var_enabled).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(info_frame, text="트리거:").pack(side=tk.LEFT)
        self.var_trigger = tk.StringVar(value="듀얼모니터 감지")
        self.cmb_trigger = ttk.Combobox(info_frame, textvariable=self.var_trigger,
                                        values=list(self._TRIGGER_LABELS.keys()),
                                        state="readonly", width=14)
        self.cmb_trigger.pack(side=tk.LEFT, padx=(5, 0))

        # ── 중단: 창 규칙 목록 ──
        lf_rules = ttk.LabelFrame(frame, text="창 규칙", padding=5)
        lf_rules.grid(row=3, column=0, columnspan=4, sticky=tk.NSEW, pady=(8, 0))
        lf_rules.columnconfigure(0, weight=1)
        lf_rules.rowconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        rule_cols = ("Class", "Title 패턴", "이동 위치", "최대화")
        self.rule_tree = ttk.Treeview(lf_rules, columns=rule_cols, show="headings", height=8)
        self.rule_tree.heading("Class", text="Class")
        self.rule_tree.heading("Title 패턴", text="Title 패턴")
        self.rule_tree.heading("이동 위치", text="이동 위치")
        self.rule_tree.heading("최대화", text="Max")
        self.rule_tree.column("Class", width=140)
        self.rule_tree.column("Title 패턴", width=120)
        self.rule_tree.column("이동 위치", width=180)
        self.rule_tree.column("최대화", width=40, anchor=tk.CENTER)
        self.rule_tree.grid(row=0, column=0, sticky=tk.NSEW)
        rule_scroll = ttk.Scrollbar(lf_rules, orient=tk.VERTICAL, command=self.rule_tree.yview)
        rule_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.rule_tree.config(yscrollcommand=rule_scroll.set)

        rule_btn_frame = ttk.Frame(lf_rules)
        rule_btn_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(3, 0))
        ttk.Button(rule_btn_frame, text="추가", width=6, command=self._add_rule).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(rule_btn_frame, text="편집", width=6, command=self._edit_rule).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(rule_btn_frame, text="삭제", width=6, command=self._delete_rule).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(rule_btn_frame, text="\u25b2", width=3, command=lambda: self._reorder_rule(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(rule_btn_frame, text="\u25bc", width=3, command=lambda: self._reorder_rule(1)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(rule_btn_frame, text="현재 창 일괄 캡처", command=self._capture_all_windows).pack(side=tk.LEFT)

        # ── 하단: 확인/취소 ──
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=4, pady=(10, 0))
        ttk.Button(btn_frame, text="확인", command=self._on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=self.destroy).pack(side=tk.LEFT, padx=5)

        # 기존 데이터 로드
        if profile:
            self.ent_name.insert(0, profile["name"])
            self.ent_exe.insert(0, profile["exe_name"])
            self.var_enabled.set(bool(profile.get("enabled", 1)))
            trigger = profile.get("trigger_mode", "monitor_change")
            self.var_trigger.set(self._TRIGGER_KEYS.get(trigger, "듀얼모니터 감지"))
            self.rules_data = db_fetch_rules(profile["id"])
        self._refresh_rules_tree()

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

    def _refresh_rules_tree(self):
        for item in self.rule_tree.get_children():
            self.rule_tree.delete(item)
        for idx, r in enumerate(self.rules_data):
            mode = r.get("move_mode", "custom")
            mode_str = self._MODE_DISPLAY.get(mode, mode) or mode
            x, y = r.get("target_x", 0), r.get("target_y", 0)
            w, h = r.get("target_w", 0), r.get("target_h", 0)
            if mode in ("custom", "save_position") and (x or y or w or h):
                size = f" {w}x{h}" if w and h else ""
                mode_str += f" ({x},{y}{size})"
            max_mark = "O" if r.get("maximize", 0) else ""
            cls = r.get("window_class", "") or "(전체)"
            title = r.get("title_pattern", "") or "(전체)"
            self.rule_tree.insert("", tk.END, iid=str(idx), values=(cls, title, mode_str, max_mark))

    def _get_selected_rule_idx(self):
        sel = self.rule_tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _add_rule(self):
        exe = self.ent_exe.get().strip()
        dlg = RuleEditDialog(self, exe_name=exe)
        self.wait_window(dlg)
        if dlg.result:
            self.rules_data.append(dlg.result)
            self._refresh_rules_tree()

    def _edit_rule(self):
        idx = self._get_selected_rule_idx()
        if idx is None:
            messagebox.showinfo("알림", "규칙을 선택하세요.", parent=self)
            return
        exe = self.ent_exe.get().strip()
        dlg = RuleEditDialog(self, rule=self.rules_data[idx], exe_name=exe)
        self.wait_window(dlg)
        if dlg.result:
            dlg.result["id"] = self.rules_data[idx].get("id")
            self.rules_data[idx] = dlg.result
            self._refresh_rules_tree()

    def _delete_rule(self):
        idx = self._get_selected_rule_idx()
        if idx is None:
            return
        self.rules_data.pop(idx)
        self._refresh_rules_tree()

    def _reorder_rule(self, direction):
        idx = self._get_selected_rule_idx()
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.rules_data):
            return
        self.rules_data[idx], self.rules_data[new_idx] = self.rules_data[new_idx], self.rules_data[idx]
        self._refresh_rules_tree()
        self.rule_tree.selection_set(str(new_idx))

    def _capture_all_windows(self):
        """프로세스의 모든 visible 창 위치를 일괄 캡처하여 규칙 생성"""
        exe_name = self.ent_exe.get().strip().lower()
        if not exe_name:
            messagebox.showwarning("입력 오류", "프로세스명을 먼저 입력하세요.", parent=self)
            return

        windows = _enumerate_process_windows(exe_name)
        if not windows:
            messagebox.showinfo("알림", f"{exe_name} 실행중인 창을 찾을 수 없습니다.", parent=self)
            return

        if self.rules_data:
            if not messagebox.askyesno("확인", f"{len(windows)}개 창이 발견되었습니다.\n"
                                       "기존 규칙을 모두 교체하시겠습니까?", parent=self):
                return

        self.rules_data = []
        for w in windows:
            self.rules_data.append({
                "id": None,
                "window_class": w["class_name"],
                "title_pattern": w["title"],
                "move_mode": "custom",
                "target_x": w["x"], "target_y": w["y"],
                "target_w": w["w"], "target_h": w["h"],
                "maximize": 0,
            })
        self._refresh_rules_tree()
        messagebox.showinfo("완료", f"{len(windows)}개 창 규칙이 캡처되었습니다.", parent=self)

    def _on_ok(self):
        name = self.ent_name.get().strip()
        exe_name = self.ent_exe.get().strip()
        if not name or not exe_name:
            messagebox.showwarning("입력 오류", "이름과 프로세스명을 입력하세요.", parent=self)
            return
        trigger = self._TRIGGER_LABELS.get(self.var_trigger.get(), "monitor_change")
        self.result = {
            "id": self.profile.get("id") if self.profile else None,
            "name": name,
            "exe_name": exe_name,
            "enabled": self.var_enabled.get(),
            "trigger_mode": trigger,
            "rules": self.rules_data,
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
        self._process_check_counter = 0
        self._profile_moved_pids: dict[int, set[int]] = {}  # profile_id → 이동 완료된 PID set

        self._build_ui()
        self._restore_window()
        self._refresh_pc_list()
        self._refresh_task_list()
        self._refresh_profile_list()
        self._refresh_routine_list()
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
        self.var_git_open_folder = tk.BooleanVar(value=bool(self.settings.get("git_open_folder", False)))
        self._build_menubar()
        self._apply_topmost()

        # ── row 1: GitHub 다운로드 ──
        git_frame = ttk.Frame(root)
        git_frame.grid(row=1, column=0, sticky=tk.EW, padx=8, pady=(3, 2))
        git_frame.columnconfigure(1, weight=1)

        ttk.Label(git_frame, text="깃허브:").grid(row=0, column=0, padx=(0, 3))

        self.git_url_var = tk.StringVar()
        self._git_download_queue: list[str] = []
        self.git_url_entry = ttk.Entry(git_frame, textvariable=self.git_url_var, font=("Consolas", 9))
        self.git_url_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0, 3))
        self.git_url_entry.bind("<Return>", lambda e: self._git_download())
        self.git_url_entry.bind("<<Paste>>", self._on_git_paste)

        self.git_dl_btn = ttk.Button(git_frame, text="Git", width=4, command=self._git_download)
        self.git_dl_btn.grid(row=0, column=2)

        ttk.Button(git_frame, text="깃 다운 폴더 열기", command=self._open_git_root_folder).grid(row=0, column=3, padx=(3, 0))

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
        self.task_tree.bind("<Delete>", self._delete_task)

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

        lf_move = ttk.LabelFrame(mid_frame, text="창 배치 프로파일", padding=5)
        lf_move.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 4))
        lf_move.columnconfigure(0, weight=1)
        lf_move.rowconfigure(0, weight=1)

        move_frame = ttk.Frame(lf_move)
        move_frame.grid(row=0, column=0, sticky=tk.NSEW)
        move_frame.columnconfigure(0, weight=1)
        move_frame.rowconfigure(0, weight=1)

        prof_cols = ("사용", "이름", "프로세스명", "트리거", "규칙수")
        self.profile_tree = ttk.Treeview(move_frame, columns=prof_cols, show="headings", height=4)
        self.profile_tree.heading("사용", text="사용")
        self.profile_tree.heading("이름", text="이름")
        self.profile_tree.heading("프로세스명", text="프로세스명")
        self.profile_tree.heading("트리거", text="트리거")
        self.profile_tree.heading("규칙수", text="규칙")
        self.profile_tree.column("사용", width=40, anchor=tk.CENTER)
        self.profile_tree.column("이름", width=100)
        self.profile_tree.column("프로세스명", width=100)
        self.profile_tree.column("트리거", width=90)
        self.profile_tree.column("규칙수", width=40, anchor=tk.CENTER)
        self.profile_tree.grid(row=0, column=0, sticky=tk.NSEW)
        move_scroll = ttk.Scrollbar(move_frame, orient=tk.VERTICAL, command=self.profile_tree.yview)
        move_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.profile_tree.config(yscrollcommand=move_scroll.set)
        self.profile_tree.bind("<Double-1>", self._on_profile_double_click)
        self.profile_tree.bind("<Button-3>", self._on_profile_right_click)
        self.profile_tree.bind("<Delete>", self._delete_profile)

        move_btn_frame = ttk.Frame(lf_move)
        move_btn_frame.grid(row=1, column=0, sticky=tk.W, pady=(3, 0))
        ttk.Button(move_btn_frame, text="추가", width=6, command=self._add_profile).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(move_btn_frame, text="\u25b2", width=3, command=lambda: self._reorder_profile(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(move_btn_frame, text="\u25bc", width=3, command=lambda: self._reorder_profile(1)).pack(side=tk.LEFT)

        # ── PC 관리 (우측, 고정 폭) ──
        lf_pc = ttk.LabelFrame(mid_frame, text="PC 관리", padding=5)
        lf_pc.grid(row=0, column=1, sticky=tk.NSEW)
        lf_pc.rowconfigure(0, weight=1)

        self.pc_listbox = tk.Listbox(lf_pc, height=6, width=14, font=("Consolas", 10), activestyle="none")
        self.pc_listbox.grid(row=0, column=0, sticky=tk.NSEW)
        pc_scroll = ttk.Scrollbar(lf_pc, orient=tk.VERTICAL, command=self.pc_listbox.yview)
        pc_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.pc_listbox.config(yscrollcommand=pc_scroll.set)
        self.pc_listbox.bind("<Double-1>", lambda e: self._boot_pc())
        self.pc_listbox.bind("<Button-3>", self._on_pc_right_click)

        pc_btn_frame = ttk.Frame(lf_pc)
        pc_btn_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(3, 0))
        ttk.Button(pc_btn_frame, text="추가", width=5, command=self._add_pc).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(pc_btn_frame, text="삭제", width=5, command=self._delete_pc).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(pc_btn_frame, text="\u25b2", width=2, command=lambda: self._move_pc(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(pc_btn_frame, text="\u25bc", width=2, command=lambda: self._move_pc(1)).pack(side=tk.LEFT)

        # ── row 4: 과제 (전체 폭) ──
        lf_routine = ttk.LabelFrame(root, text="일정", padding=5)
        lf_routine.grid(row=4, column=0, sticky=tk.NSEW, padx=8, pady=2)
        lf_routine.columnconfigure(0, weight=1)
        lf_routine.rowconfigure(0, weight=1)

        routine_frame = ttk.Frame(lf_routine)
        routine_frame.grid(row=0, column=0, sticky=tk.NSEW)
        routine_frame.columnconfigure(0, weight=1)
        routine_frame.rowconfigure(0, weight=1)

        rt_cols = ("날짜", "내용", "완료시간", "경과")
        self._hidden_routine_dates = db_fetch_hidden_routine_dates()
        self.routine_tree = ttk.Treeview(routine_frame, columns=rt_cols, show="headings", height=8)
        self.routine_tree.heading("날짜", text="날짜")
        self.routine_tree.heading("내용", text="내용")
        self.routine_tree.heading("완료시간", text="완료시간")
        self.routine_tree.heading("경과", text="경과")
        self.routine_tree.column("날짜", width=80, anchor=tk.CENTER)
        self.routine_tree.column("내용", width=110)
        self.routine_tree.column("완료시간", width=130, anchor=tk.CENTER)
        self.routine_tree.column("경과", width=70, anchor=tk.CENTER)
        self.routine_tree.tag_configure("missed", foreground="gray")
        self.routine_tree.tag_configure("done", foreground="gray")
        self.routine_tree.grid(row=0, column=0, sticky=tk.NSEW)
        rt_scroll = ttk.Scrollbar(routine_frame, orient=tk.VERTICAL, command=self.routine_tree.yview)
        rt_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.routine_tree.config(yscrollcommand=rt_scroll.set)
        self.routine_tree.bind("<Double-1>", self._on_routine_double_click)
        self.routine_tree.bind("<Button-3>", self._on_routine_right_click)
        self.routine_tree.bind("<Delete>", self._hide_routine_date)

        routine_btn_frame = ttk.Frame(lf_routine)
        routine_btn_frame.grid(row=1, column=0, sticky=tk.W, pady=(3, 0))
        ttk.Button(routine_btn_frame, text="추가", width=6, command=self._add_routine).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Button(routine_btn_frame, text="▲", width=3, command=lambda: self._move_routine(-1)).pack(side=tk.LEFT, padx=(0, 1))
        ttk.Button(routine_btn_frame, text="▼", width=3, command=lambda: self._move_routine(1)).pack(side=tk.LEFT)

        # ── row 5: 로그 (전체 폭) ──
        lf_log = ttk.LabelFrame(root, text="로그", padding=5)
        lf_log.grid(row=5, column=0, sticky=tk.NSEW, padx=8, pady=(2, 8))
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

        # 행 가중치: row 0,1 고정, row 2(자동실행) 고정, row 3(이동대상+PC) 고정, row 4(일정) 고정, row 5(로그) 확장
        root.rowconfigure(0, weight=0)
        root.rowconfigure(1, weight=0)
        root.rowconfigure(2, weight=0)
        root.rowconfigure(3, weight=0)
        root.rowconfigure(4, weight=0)
        root.rowconfigure(5, weight=1)

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
        self.settings["git_open_folder"] = self.var_git_open_folder.get()  # type: ignore[assignment]
        save_local_settings(self.settings)

    def _open_git_root_folder(self):
        """깃허브 다운로드 루트 폴더를 연다 (.env의 CLONE_BASE_PATH 우선)"""
        clone_base = os.getenv("CLONE_BASE_PATH", "").strip()
        if not clone_base:
            clone_base = os.path.join(SCRIPT_DIR, "data")
        if not os.path.isdir(clone_base):
            try:
                os.makedirs(clone_base, exist_ok=True)
            except OSError as e:
                self.log(f"[GitHub] 루트 폴더 생성 실패: {e}")
                return
        subprocess.Popen(["explorer.exe", os.path.normpath(clone_base)])

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
        menu_settings.add_checkbutton(label="깃 다운시 폴더 열기", variable=self.var_git_open_folder,
                                      command=self._save_git_open_folder)
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

    # ─── 과제 (루틴) ──────────────────────────────────────
    def _refresh_routine_list(self):
        """과제 Treeview 갱신 — 매일 반복 일정은 과거 미완료 날짜도 회색으로 표시"""
        self.routine_data = db_fetch_routines()
        self.routine_tree.delete(*self.routine_tree.get_children())
        today_str = date.today().strftime("%Y-%m-%d")
        # 모든 행을 수집 후 날짜 → 정렬순으로 정렬하여 삽입
        rows = []
        for rt in self.routine_data:
            if not rt["enabled"]:
                continue
            start = rt.get("start_date", "")
            repeat_type = rt.get("repeat_type", "once")
            if start and start > today_str:
                continue
            if repeat_type == "once" and start and start != today_str:
                continue
            display_dates = db_get_routine_display_dates(
                rt["id"], 1, start, repeat_type)
            for date_str, done, is_past in display_dates:
                logs = db_fetch_routine_logs(rt["id"], date_str)
                raw_time = logs[-1]["done_time"] if logs else ""
                display_time = format_done_time_display(raw_time)
                # 경과시간: 완료 후 현재까지, 미완료 시 이전 완료로부터 현재까지
                elapsed = ""
                now = datetime.now()
                if raw_time:
                    done_dt = parse_done_datetime(date_str, raw_time)
                    if done_dt:
                        elapsed = format_elapsed((now - done_dt).total_seconds())
                elif done == 0:
                    prev_dt = db_get_prev_routine_done_time(rt["id"], date_str)
                    if prev_dt and prev_dt.date() < now.date():
                        elapsed = format_elapsed((now - prev_dt).total_seconds())
                iid = f"{rt['id']}:{date_str}"
                if (rt["id"], date_str) in self._hidden_routine_dates:
                    continue
                tag = ("done",) if done >= 1 else ("missed",) if is_past else ()
                rows.append((date_str, rt.get("sort_order", 0), iid,
                             (date_str, rt["name"], display_time, elapsed), tag))
        rows.sort(key=lambda r: (r[0], r[1]))
        for _, _, iid, values, tag in rows:
            self.routine_tree.insert("", tk.END, iid=iid, values=values, tags=tag)

    def _refresh_routine_elapsed(self):
        """일정 경과시간 컬럼만 갱신 (전체 리프레시 없이)"""
        now = datetime.now()
        for iid in self.routine_tree.get_children():
            parts = iid.split(":")
            rt_id = int(parts[0])
            date_str = parts[1]
            vals = list(self.routine_tree.item(iid, "values"))
            display_time = vals[2] if len(vals) > 2 else ""
            elapsed = ""
            if display_time:
                logs = db_fetch_routine_logs(rt_id, date_str)
                raw_time = logs[-1]["done_time"] if logs else ""
                done_dt = parse_done_datetime(date_str, raw_time)
                if done_dt:
                    elapsed = format_elapsed((now - done_dt).total_seconds())
            else:
                prev_dt = db_get_prev_routine_done_time(rt_id, date_str)
                if prev_dt and prev_dt.date() < now.date():
                    elapsed = format_elapsed((now - prev_dt).total_seconds())
            self.routine_tree.set(iid, "경과", elapsed)

    def _get_selected_routine(self):
        sel = self.routine_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "일정을 선택하세요.", parent=self.root)
            return None
        rt_id = int(sel[0].split(":")[0])
        for rt in self.routine_data:
            if rt["id"] == rt_id:
                return rt
        return None

    def _get_selected_routine_date(self):
        """선택된 행의 날짜 반환"""
        sel = self.routine_tree.selection()
        if not sel:
            return None
        parts = sel[0].split(":")
        return parts[1] if len(parts) > 1 else None

    def _on_routine_double_click(self, event):
        """더블클릭: 완료 처리"""
        self._complete_routine()

    def _on_routine_right_click(self, event):
        """우클릭: 과제 컨텍스트 메뉴"""
        item = self.routine_tree.identify_row(event.y)
        if not item:
            return
        if item not in self.routine_tree.selection():
            self.routine_tree.selection_set(item)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="완료 취소", command=self._undo_routine)
        menu.add_command(label="목록에서 삭제", command=self._hide_routine_date)
        menu.add_separator()
        menu.add_command(label="일정 편집", command=self._edit_routine)
        menu.add_command(label="일정 삭제", command=self._delete_routine)
        menu.tk_popup(event.x_root, event.y_root)

    def _add_routine(self):
        dlg = RoutineEditDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_routine(None, r["name"], r["daily_count"], r["enabled"], r["start_date"], r["repeat_type"])
            self._refresh_routine_list()
            self.log(f"일정 추가: {r['name']}")

    def _edit_routine(self):
        rt = self._get_selected_routine()
        if not rt:
            return
        dlg = RoutineEditDialog(self.root, routine=rt)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_routine(r["id"], r["name"], r["daily_count"], r["enabled"], r["start_date"], r["repeat_type"])
            self._refresh_routine_list()
            self.log(f"일정 수정: {r['name']}")

    def _hide_routine_date(self, event=None):
        """선택한 날짜 행을 UI 목록에서만 숨김 (DB 삭제 없음)"""
        sel = self.routine_tree.selection()
        if not sel:
            return
        hidden = []
        for iid in sel:
            parts = iid.split(":")
            if len(parts) < 2:
                continue
            rt_id = int(parts[0])
            active_date = parts[1]
            rt = next((r for r in self.routine_data if r["id"] == rt_id), None)
            if not rt:
                continue
            self._hidden_routine_dates.add((rt_id, active_date))
            db_hide_routine_date(rt_id, active_date)
            self.routine_tree.delete(iid)
            hidden.append(f"{rt['name']}({active_date})")
        if hidden:
            self.log(f"목록에서 숨김: {', '.join(hidden)}")

    def _delete_routine(self, event=None):
        sel = self.routine_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "일정을 선택하세요.", parent=self.root)
            return
        # 선택된 항목에서 고유 일정 ID 추출
        rt_ids_seen = set()
        routines = []
        for iid in sel:
            rt_id = int(iid.split(":")[0])
            if rt_id in rt_ids_seen:
                continue
            rt_ids_seen.add(rt_id)
            for rt in self.routine_data:
                if rt["id"] == rt_id:
                    routines.append(rt)
                    break
        if not routines:
            return
        if len(routines) == 1:
            msg = f"'{routines[0]['name']}' 일정을 중단하시겠습니까?\n(과거 기록은 보존됩니다)"
        else:
            msg = f"{len(routines)}개 일정을 중단하시겠습니까?\n(과거 기록은 보존됩니다)\n" + ", ".join(r["name"] for r in routines)
        if not messagebox.askyesno("확인", msg, parent=self.root):
            return
        for rt in routines:
            db_delete_routine(rt["id"])
        self._refresh_routine_list()
        names = ", ".join(rt["name"] for rt in routines)
        self.log(f"일정 삭제: {names}")

    def _complete_routine(self):
        """선택한 일정의 해당 날짜에 완료 처리"""
        rt = self._get_selected_routine()
        if not rt:
            return
        active_date = self._get_selected_routine_date()
        if not active_date:
            return
        logs = db_fetch_routine_logs(rt["id"], active_date)
        if logs:
            messagebox.showinfo("알림", f"'{rt['name']}' 은(는) {active_date} 이미 완료되었습니다.",
                                parent=self.root)
            return
        db_add_routine_log(rt["id"], active_date, 1)
        self.log(f"일정 완료: {rt['name']} ({active_date})")
        # 1회 일정은 완료 시 비활성화
        if rt.get("repeat_type", "once") == "once":
            db_upsert_routine(rt["id"], rt["name"], 1, 0,
                              rt.get("start_date", ""), "once")
            self.log(f"1회 일정 완료 → 비활성화: {rt['name']}")
        self._refresh_routine_list()

    def _undo_routine(self):
        """잘못 누른 확인을 되돌리기 (마지막 1회 취소)"""
        rt = self._get_selected_routine()
        if not rt:
            return
        active_date = self._get_selected_routine_date()
        if not active_date:
            return
        logs = db_fetch_routine_logs(rt["id"], active_date)
        if not logs:
            messagebox.showinfo("알림", f"'{rt['name']}' 되돌릴 기록이 없습니다.",
                                parent=self.root)
            return
        db_remove_last_routine_log(rt["id"], active_date)
        done = len(logs) - 1
        self._refresh_routine_list()
        self.log(f"일정 확인 취소: {rt['name']} ({active_date} {done}/{rt['daily_count']})")

    def _move_routine(self, direction):
        """과제 순서 이동"""
        rt = self._get_selected_routine()
        if not rt:
            return
        enabled_data = [r for r in self.routine_data if r["enabled"]]
        idx = next((i for i, r in enumerate(enabled_data) if r["id"] == rt["id"]), -1)
        target = idx + direction
        if target < 0 or target >= len(enabled_data):
            return
        db_swap_sort_order("routines", rt["id"], enabled_data[target]["id"])
        self._refresh_routine_list()
        # 이동 후 오늘 날짜 행 선택
        today_iid = f"{rt['id']}:{date.today().strftime('%Y-%m-%d')}"
        if self.routine_tree.exists(today_iid):
            self.routine_tree.selection_set(today_iid)

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

    def _on_pc_right_click(self, event):
        """우클릭: PC 컨텍스트 메뉴"""
        idx = self.pc_listbox.nearest(event.y)
        if idx < 0:
            return
        self.pc_listbox.selection_clear(0, tk.END)
        self.pc_listbox.selection_set(idx)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="편집", command=self._edit_pc)
        menu.add_command(label="부팅", command=self._boot_pc)
        menu.add_command(label="핑", command=self._ping_pc)
        menu.tk_popup(event.x_root, event.y_root)

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
        """더블클릭: 기존 동작 (폴더 열기 / 창 활성화)"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        # 더블클릭한 항목을 선택 상태로 동기화
        self.task_tree.selection_set(item)
        task_id = int(item)
        task = next((t for t in self.task_data if t["id"] == task_id), None)
        if not task:
            return
        executable = task["executable"]
        if os.path.isdir(executable):
            self._open_folder_task(task)
            return
        # AutoExec이 실행한 프로세스 → PID로 창 활성화 (빠름)
        if task_id in self._running_tasks:
            proc = self._task_processes.get(task_id)
            if proc and proc.poll() is None and self._activate_window_by_pid(proc.pid):
                self.log(f"[자동실행] {task['name']} 창 활성화")
                return
        if executable.lower().endswith(".pyw"):
            self._run_task(skip_activation=True)
        elif task["enabled"]:
            self._edit_task()
        else:
            self._run_task(skip_activation=True)

    def _on_task_right_click(self, event):
        """우클릭: 컨텍스트 메뉴 표시"""
        item = self.task_tree.identify_row(event.y)
        if not item:
            return
        if item not in self.task_tree.selection():
            self.task_tree.selection_set(item)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="실행", command=self._run_task)
        menu.add_command(label="폴더 열기", command=self._open_task_folder)
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

    def _delete_task(self, event=None):
        sel = self.task_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "자동실행 항목을 선택하세요.", parent=self.root)
            return
        tasks = []
        for iid in sel:
            task_id = int(iid)
            for t in self.task_data:
                if t["id"] == task_id:
                    tasks.append(t)
                    break
        if not tasks:
            return
        if len(tasks) == 1:
            msg = f"'{tasks[0]['name']}' 작업을 삭제하시겠습니까?"
        else:
            msg = f"{len(tasks)}개 작업을 삭제하시겠습니까?\n" + ", ".join(t["name"] for t in tasks)
        if messagebox.askyesno("삭제 확인", msg, parent=self.root):
            for t in tasks:
                db_delete_task(t["id"])
            self._refresh_task_list()
            names = ", ".join(t["name"] for t in tasks)
            self.log(f"자동실행 삭제: {names}")

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

    def _open_task_folder(self):
        """선택한 자동실행 항목의 경로를 탐색기로 열기"""
        task = self._get_selected_task()
        if not task:
            return
        executable = task["executable"]
        if not executable:
            return
        path = os.path.normpath(executable)
        if os.path.isdir(path):
            subprocess.Popen(["explorer.exe", path])
            self.log(f"[자동실행] 폴더 열기: {path}")
        elif os.path.isfile(path):
            folder = os.path.dirname(path)
            subprocess.Popen(["explorer.exe", "/select,", path])
            self.log(f"[자동실행] 폴더 열기: {folder}")
        else:
            messagebox.showinfo("알림", f"경로를 찾을 수 없습니다:\n{path}", parent=self.root)

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

    # ─── 창 배치 프로파일 관리 ─────────────────────────────
    _TRIGGER_DISPLAY = {"monitor_change": "듀얼모니터", "process_start": "프로세스", "both": "둘 다"}

    def _refresh_profile_list(self):
        self.profile_data = db_fetch_profiles()
        for item in self.profile_tree.get_children():
            self.profile_tree.delete(item)
        for p in self.profile_data:
            enabled_mark = "O" if p.get("enabled", 1) else ""
            trigger = self._TRIGGER_DISPLAY.get(p.get("trigger_mode", "monitor_change"), "")
            rule_count = db_count_rules(p["id"])
            self.profile_tree.insert("", tk.END, iid=str(p["id"]),
                                     values=(enabled_mark, p["name"], p["exe_name"], trigger, rule_count))

    def _get_selected_profile(self):
        sel = self.profile_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "프로파일을 선택하세요.", parent=self.root)
            return None
        profile_id = int(sel[0])
        for p in self.profile_data:
            if p["id"] == profile_id:
                return p
        return None

    def _add_profile(self):
        dlg = ProfileEditDialog(self.root)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            pid = db_upsert_profile(None, r["name"], r["exe_name"], r["enabled"], r["trigger_mode"])
            db_replace_rules(pid, r["rules"])
            self._refresh_profile_list()
            self.log(f"[프로파일] 추가: {r['name']} ({r['exe_name']})")

    def _edit_profile(self):
        profile = self._get_selected_profile()
        if not profile:
            return
        dlg = ProfileEditDialog(self.root, profile)
        self.root.wait_window(dlg)
        if dlg.result:
            r = dlg.result
            db_upsert_profile(r["id"], r["name"], r["exe_name"], r["enabled"], r["trigger_mode"])
            db_replace_rules(r["id"], r["rules"])
            self._refresh_profile_list()
            self.log(f"[프로파일] 수정: {r['name']} ({r['exe_name']})")

    def _delete_profile(self, event=None):
        sel = self.profile_tree.selection()
        if not sel:
            messagebox.showinfo("알림", "프로파일을 선택하세요.", parent=self.root)
            return
        profiles = []
        for iid in sel:
            profile_id = int(iid)
            for p in self.profile_data:
                if p["id"] == profile_id:
                    profiles.append(p)
                    break
        if not profiles:
            return
        if len(profiles) == 1:
            msg = f"'{profiles[0]['name']}' 프로파일을 삭제하시겠습니까?"
        else:
            msg = f"{len(profiles)}개 프로파일을 삭제하시겠습니까?\n" + ", ".join(p["name"] for p in profiles)
        if messagebox.askyesno("삭제 확인", msg, parent=self.root):
            for p in profiles:
                db_delete_profile(p["id"])
            self._refresh_profile_list()
            names = ", ".join(p["name"] for p in profiles)
            self.log(f"[프로파일] 삭제: {names}")

    def _reorder_profile(self, direction):
        sel = self.profile_tree.selection()
        if not sel:
            return
        profile_id = int(sel[0])
        idx = next((i for i, p in enumerate(self.profile_data) if p["id"] == profile_id), None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.profile_data):
            return
        db_swap_sort_order("window_profiles", self.profile_data[idx]["id"], self.profile_data[new_idx]["id"])
        self._refresh_profile_list()
        new_iid = str(profile_id)
        if self.profile_tree.exists(new_iid):
            self.profile_tree.selection_set(new_iid)
            self.profile_tree.see(new_iid)

    def _on_profile_right_click(self, event):
        """우클릭: 프로파일 컨텍스트 메뉴"""
        item = self.profile_tree.identify_row(event.y)
        if not item:
            return
        if item not in self.profile_tree.selection():
            self.profile_tree.selection_set(item)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="편집", command=self._edit_profile)
        menu.add_command(label="삭제", command=self._delete_profile)
        menu.tk_popup(event.x_root, event.y_root)

    def _on_profile_double_click(self, event):
        """더블클릭: 창 이동"""
        item = self.profile_tree.identify_row(event.y)
        if not item:
            return
        profile_id = int(item)
        profile = next((p for p in self.profile_data if p["id"] == profile_id), None)
        if not profile:
            return
        rules = db_fetch_rules(profile["id"])
        if not rules:
            self.log(f"[프로파일] {profile['name']} 규칙 없음")
            return

        def _do_move():
            moved = self._apply_profile_rules(profile, rules)
            self.log(f"[프로파일] {profile['name']} {moved}개 창 이동 완료")

        threading.Thread(target=_do_move, daemon=True).start()

    def _manual_move_profile(self):
        """선택한 프로파일의 규칙에 따라 창을 즉시 이동"""
        profile = self._get_selected_profile()
        if not profile:
            return
        rules = db_fetch_rules(profile["id"])
        if not rules:
            self.log(f"[프로파일] {profile['name']} 규칙 없음")
            return

        def _do_move():
            moved = self._apply_profile_rules(profile, rules)
            self.log(f"[프로파일] {profile['name']} {moved}개 창 이동 완료")

        threading.Thread(target=_do_move, daemon=True).start()

    def _apply_profile_rules(self, profile, rules, sub_monitor_info=None):
        """프로파일의 규칙에 따라 창 이동. 반환: 이동된 창 수."""
        exe_name = profile["exe_name"].lower()
        windows = _enumerate_process_windows(exe_name)
        if not windows:
            return 0

        # 서브 모니터 정보 (필요 시)
        sub = sub_monitor_info
        if sub is None:
            for rule in rules:
                if rule.get("move_mode") == "sub_monitor":
                    monitors = self._get_monitors_info()
                    for m in monitors:
                        if not m[4]:
                            sub = m
                            break
                    break

        moved = 0
        user32 = ctypes.windll.user32
        for w in windows:
            rule = _match_window_to_rules(w["class_name"], w["title"], rules)
            if not rule:
                continue
            hwnd = w["hwnd"]
            try:
                if user32.IsIconic(hwnd):
                    continue
                move_mode = rule.get("move_mode", "custom")
                maximize = rule.get("maximize", 0)
                if move_mode == "sub_monitor":
                    if not sub:
                        continue
                    tx, ty, tw, th = sub[:4]
                    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                    time.sleep(0.1)
                    user32.SetWindowPos(hwnd, 0, tx + 100, ty + 100, tw - 200, th - 200, 0x0004)
                    if maximize:
                        time.sleep(0.1)
                        user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                else:  # custom, save_position
                    tx = rule.get("target_x", 0)
                    ty = rule.get("target_y", 0)
                    tw = rule.get("target_w", 0)
                    th = rule.get("target_h", 0)
                    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                    time.sleep(0.1)
                    if tw > 0 and th > 0:
                        user32.SetWindowPos(hwnd, 0, tx, ty, tw, th, 0x0004)
                    else:
                        rect = ctypes.wintypes.RECT()
                        user32.GetWindowRect(hwnd, ctypes.byref(rect))
                        cw = rect.right - rect.left
                        ch = rect.bottom - rect.top
                        user32.SetWindowPos(hwnd, 0, tx, ty, cw, ch, 0x0004)
                    if maximize:
                        time.sleep(0.1)
                        user32.ShowWindow(hwnd, 3)
                moved += 1
            except Exception:
                pass
        return moved

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
        """붙여넣기된 텍스트에서 GitHub URL을 추출하여 자동 다운로드"""
        text = self.git_url_var.get().strip()
        if not text or "github.com/" not in text:
            return
        urls = self._extract_git_urls(text)
        if not urls:
            self.git_url_var.set("")
            self.log("[GitHub] 잘못된 붙여넣기 감지 (GitHub URL을 찾을 수 없음)")
            return
        if len(urls) == 1:
            self.git_url_var.set(urls[0])
            self._git_download()
        else:
            self.git_url_var.set("")
            self._git_download_queue = list(urls)
            self.log(f"[GitHub] {len(urls)}개 저장소 일괄 다운로드 시작")
            self._git_download_next()

    def _git_download_next(self):
        """큐에서 다음 URL을 꺼내 다운로드"""
        if not self._git_download_queue:
            self.git_url_var.set("")
            self.git_url_entry.focus_set()
            self.log("[GitHub] 일괄 다운로드 완료")
            return
        url = self._git_download_queue.pop(0)
        self.git_url_var.set(url)
        self._git_download(on_done=self._git_download_next)

    @staticmethod
    def _extract_git_urls(text: str) -> list[str]:
        """텍스트에서 모든 GitHub URL을 추출. https:// 없어도 인식"""
        import re
        matches = re.findall(r'(?:https?://)?github\.com/[\w.\-]+/[\w.\-]+', text)
        urls = []
        for m in matches:
            url = m if m.startswith("http") else "https://" + m
            # 중복 제거
            if url not in urls:
                urls.append(url)
        return urls

    @staticmethod
    def _is_valid_git_url(text: str) -> bool:
        """GitHub URL 유효성 검증"""
        import re
        return bool(re.match(r'^(?:https?://)?github\.com/[\w.\-]+/[\w.\-]+/?$', text))

    def _git_download(self, on_done=None):
        """GitHub 저장소 다운로드 (gitclone.py 호출)"""
        url = self.git_url_var.get().strip()
        if not url:
            if on_done:
                self.root.after(0, on_done)
            return
        if not self._is_valid_git_url(url):
            self.log("[GitHub] 잘못된 URL 형식입니다")
            if on_done:
                self.root.after(0, on_done)
            return
        if not url.startswith("http"):
            url = "https://" + url
            self.git_url_var.set(url)
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
                self.root.after(0, lambda: self._git_download_done(success, output, on_done))
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self._git_download_done(False, "타임아웃 (120초 초과)", on_done))
            except Exception as e:
                self.root.after(0, lambda: self._git_download_done(False, str(e), on_done))

        threading.Thread(target=_worker, daemon=True).start()

    def _git_download_done(self, success, output, on_done=None):
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
        # 큐에 다음 항목이 있으면 계속 진행
        if on_done:
            on_done()
        else:
            self.git_url_entry.focus_set()

    def _run_task(self, skip_activation=False):
        """선택한 자동실행 작업을 즉시 테스트 실행 (폴더면 열기, 실행중이면 활성화)"""
        task = self._get_selected_task()
        if not task:
            return
        if os.path.isdir(task["executable"]):
            self._open_folder_task(task)
            return
        if not skip_activation:
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
            self._refresh_routine_list()

        is_closed = self._is_closed_day(now)

        # 자동 부팅 체크 (PC별 skip_holiday 개별 판단)
        self._check_auto_boot(current_hm, is_closed)

        # 자동실행 체크
        self._check_auto_tasks(current_hm, today_str, is_closed)

        # 듀얼 모니터 감지 (3초마다)
        self._check_monitor_change()

        # 프로세스 감지 (3초마다)
        self._check_process_profiles()

        # 일정 경과시간 갱신 (1분마다)
        if now.second == 0:
            self._refresh_routine_elapsed()

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
        ctypes.windll.kernel32.GetTickCount64.restype = ctypes.c_uint64
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

    # ─── 듀얼 모니터 감지 + 프로세스 감지 ─────────────────
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
            self.log("[모니터] 듀얼 모니터 감지 → 프로파일 기반 창 이동")
            threading.Thread(target=self._on_monitor_change, daemon=True).start()

    def _check_process_profiles(self):
        """프로세스 감지 트리거 (3초 주기)"""
        self._process_check_counter += 1
        if self._process_check_counter < 3:
            return
        self._process_check_counter = 0

        profiles = db_fetch_profiles()
        for profile in profiles:
            if not profile.get("enabled", 1):
                continue
            trigger = profile.get("trigger_mode", "monitor_change")
            if trigger not in ("process_start", "both"):
                continue

            exe_name = profile["exe_name"].lower()
            profile_id = profile["id"]

            # 해당 프로세스의 현재 PID 집합 획득
            current_pids = set()
            try:
                result = subprocess.run(
                    ["tasklist", "/fi", f"imagename eq {exe_name}", "/fo", "csv", "/nh"],
                    capture_output=True, text=True, timeout=5,
                    creationflags=0x08000000,
                )
                for line in result.stdout.splitlines():
                    parts = line.strip().strip('"').split('","')
                    if len(parts) >= 2:
                        try:
                            current_pids.add(int(parts[1]))
                        except ValueError:
                            pass
            except Exception:
                continue

            if not current_pids:
                # 프로세스 종료 시 PID 기록 초기화
                self._profile_moved_pids.pop(profile_id, None)
                continue

            moved_pids = self._profile_moved_pids.setdefault(profile_id, set())
            # 종료된 PID 정리
            moved_pids &= current_pids
            # 새 PID 발견
            new_pids = current_pids - moved_pids
            if new_pids:
                rules = db_fetch_rules(profile_id)
                if rules:
                    def _do_move(p=profile, r=rules):
                        time.sleep(3)  # 프로세스 초기화 대기
                        moved = self._apply_profile_rules(p, r)
                        self.log(f"[프로세스감지] {p['name']} {moved}개 창 이동 완료")
                    threading.Thread(target=_do_move, daemon=True).start()
                moved_pids |= new_pids

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

    def _on_monitor_change(self):
        """듀얼 모니터 감지 시 trigger_mode이 monitor_change/both인 프로파일 실행"""
        try:
            time.sleep(2)  # 모니터 인식 안정화 대기

            monitors = self._get_monitors_info()
            sub = None
            for m in monitors:
                if not m[4]:
                    sub = m
                    break

            profiles = db_fetch_profiles()
            if not profiles:
                self.log("[모니터] 등록된 프로파일 없음")
                return

            total_moved = 0
            for profile in profiles:
                if not profile.get("enabled", 1):
                    continue
                trigger = profile.get("trigger_mode", "monitor_change")
                if trigger not in ("monitor_change", "both"):
                    continue
                rules = db_fetch_rules(profile["id"])
                if not rules:
                    continue
                moved = self._apply_profile_rules(profile, rules, sub_monitor_info=sub)
                total_moved += moved

            self.log(f"[모니터] {total_moved}개 창 이동 완료")
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
                pystray.MenuItem("열기", lambda icon, item: self.root.after(0, self._show_window), default=True, visible=False),
                pystray.MenuItem("재실행", lambda icon, item: self.root.after(0, self._restart_app)),
                pystray.MenuItem("종료", lambda icon, item: self.root.after(0, self._force_quit)),
            )
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception:
            pass

    def _show_window(self):
        self.hidden = False
        self.root.deiconify()
        self.root.lift()

    def _restart_app(self):
        """애플리케이션 재실행"""
        self._save_window()
        # 뮤텍스 소유권 및 핸들 해제
        if hasattr(self, "_mutex") and self._mutex:
            ctypes.windll.kernel32.ReleaseMutex(self._mutex)
            ctypes.windll.kernel32.CloseHandle(self._mutex)
            self._mutex = None
        # 절대 경로 + --restart 플래그로 재실행 (레이스 컨디션 회피를 위해 새 프로세스가 대기하도록)
        script_path = os.path.abspath(sys.argv[0])
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            [sys.executable, script_path] + sys.argv[1:] + ["--restart"],
            cwd=SCRIPT_DIR,
            close_fds=True,
            creationflags=DETACHED_PROCESS,
        )
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()

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
    # 재실행 플래그: 이전 프로세스가 뮤텍스를 해제할 때까지 대기
    _is_restart = "--restart" in sys.argv
    if _is_restart:
        sys.argv.remove("--restart")

    # 중복 실행 방지 (Windows Named Mutex)
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "AutoExec_Python")
    if _is_restart and ctypes.windll.kernel32.GetLastError() == 183:
        # 재실행 시 최대 5초까지 이전 프로세스 종료 대기
        import time
        ctypes.windll.kernel32.CloseHandle(_mutex)
        for _ in range(20):
            time.sleep(0.25)
            _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "AutoExec_Python")
            if ctypes.windll.kernel32.GetLastError() != 183:
                break
            ctypes.windll.kernel32.CloseHandle(_mutex)

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
    app._mutex = _mutex
    app.run()
    ctypes.windll.kernel32.CloseHandle(_mutex)
