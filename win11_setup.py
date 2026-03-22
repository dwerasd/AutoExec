# -*- coding: utf-8 -*-
#
# Windows 11 시스템 설정 적용 모듈
# - 레지스트리 설정 적용
# - 인트라넷 영역 등록
# - 시스템 명령어 실행
#

import json
import socket
import subprocess
import winreg
from pathlib import Path
from typing import Any, Callable


# 설정 파일 경로
SCRIPT_DIR = Path(__file__).parent
REGISTRY_CONFIG = SCRIPT_DIR / "registry_config.json"
COMMANDS_CONFIG = SCRIPT_DIR / "commands_config.json"


# ===== 레지스트리 관련 =====

HKEY_MAP: dict[str, int] = {
    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
    "HKLM": winreg.HKEY_LOCAL_MACHINE,
    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
    "HKCU": winreg.HKEY_CURRENT_USER,
    "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
    "HKCR": winreg.HKEY_CLASSES_ROOT,
    "HKEY_USERS": winreg.HKEY_USERS,
    "HKU": winreg.HKEY_USERS,
    "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
    "HKCC": winreg.HKEY_CURRENT_CONFIG,
}

TYPE_MAP = {
    "REG_SZ": winreg.REG_SZ,
    "REG_EXPAND_SZ": winreg.REG_EXPAND_SZ,
    "REG_BINARY": winreg.REG_BINARY,
    "REG_DWORD": winreg.REG_DWORD,
    "REG_QWORD": winreg.REG_QWORD,
    "REG_MULTI_SZ": winreg.REG_MULTI_SZ,
}


def parse_registry_path(full_path: str) -> tuple[int, str] | tuple[None, None]:
    parts = full_path.split("\\", 1)
    if len(parts) < 2:
        return None, None
    root_name = parts[0].upper()
    sub_key = parts[1]
    if root_name not in HKEY_MAP:
        return None, None
    return HKEY_MAP[root_name], sub_key


def write_registry_value(full_path: str, value_name: str, value: Any, reg_type: int) -> tuple[bool, str]:
    root_key, sub_key = parse_registry_path(full_path)
    if root_key is None or sub_key is None:
        return False, "잘못된 레지스트리 경로"
    try:
        with winreg.CreateKey(root_key, sub_key) as key:
            winreg.SetValueEx(key, value_name, 0, reg_type, value)
        return True, ""
    except PermissionError:
        return False, "관리자 권한 필요"
    except Exception as e:
        return False, str(e)


def deserialize_value(value, reg_type):
    if reg_type == winreg.REG_BINARY:
        if isinstance(value, str):
            hex_str = value.replace(",", "").replace(" ", "")
            return bytes.fromhex(hex_str)
        return value if value else b""
    elif reg_type == winreg.REG_MULTI_SZ:
        return value if value else []
    elif reg_type in (winreg.REG_DWORD, winreg.REG_QWORD):
        return int(value) if isinstance(value, str) else value
    else:
        return value


def load_registry_items() -> list[dict]:
    if not REGISTRY_CONFIG.exists():
        return []
    with open(REGISTRY_CONFIG, "r", encoding="utf-8") as f:
        config = json.load(f)
    return config.get("registry_items", [])


def apply_registry_items(items: list[dict], log: Callable[[str], None], dry_run: bool = False) -> tuple[int, int]:
    success, fail = 0, 0
    for item in items:
        path = item["path"]
        name = item["name"]
        type_name = item["type"]
        value = item["value"]
        desc = item.get("description", name)

        reg_type = TYPE_MAP.get(type_name, winreg.REG_SZ)
        final_value = deserialize_value(value, reg_type)

        if dry_run:
            log(f"[시뮬레이션] {desc}")
            success += 1
            continue

        ok, err = write_registry_value(path, name, final_value, reg_type)
        if ok:
            log(f"[완료] {desc}")
            success += 1
        else:
            log(f"[실패] {desc}: {err}")
            fail += 1

    return success, fail


# ===== 인트라넷 영역 =====

def get_local_ip() -> str | None:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return None


def get_ip_range(ip: str) -> str:
    parts = ip.split(".")
    if len(parts) == 4:
        return f"{parts[0]}.{parts[1]}.{parts[2]}.*"
    return ip


def get_existing_intranet_ranges() -> dict[str, str]:
    ranges = {}
    base_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_READ)
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ)
                try:
                    range_value, _ = winreg.QueryValueEx(subkey, ":Range")
                    ranges[subkey_name] = range_value
                except FileNotFoundError:
                    pass
                winreg.CloseKey(subkey)
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    return ranges


def setup_intranet_zone(log: Callable[[str], None], dry_run: bool = False) -> bool:
    ip = get_local_ip()
    if not ip:
        log("[인트라넷] IP 주소를 감지할 수 없습니다.")
        return False

    ip_range = get_ip_range(ip)
    log(f"[인트라넷] 현재 IP: {ip}, 대역: {ip_range}")

    existing = get_existing_intranet_ranges()
    for name, value in existing.items():
        if value == ip_range:
            log(f"[인트라넷] 이미 등록됨: {ip_range}")
            return True

    if dry_run:
        log(f"[인트라넷] 등록 예정: {ip_range}")
        return True

    max_num = 0
    for name in existing.keys():
        if name.startswith("Range"):
            try:
                num = int(name[5:])
                max_num = max(max_num, num)
            except ValueError:
                pass
    range_name = f"Range{max_num + 1}"
    base_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    range_path = f"{base_path}\\{range_name}"

    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, range_path)
        winreg.SetValueEx(key, ":Range", 0, winreg.REG_SZ, ip_range)
        winreg.SetValueEx(key, "*", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        log(f"[인트라넷] 등록 완료: {ip_range}")
        return True
    except Exception as e:
        log(f"[인트라넷] 등록 실패: {e}")
        return False


# ===== 명령어 실행 =====

def run_powershell(command: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True, text=True, encoding="utf-8", errors="replace"
        )
        output = result.stdout + result.stderr
        return result.returncode == 0, output.strip()
    except Exception as e:
        return False, str(e)


def run_cmd(command: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            ["cmd", "/c", command],
            capture_output=True, text=True, encoding="cp949", errors="replace"
        )
        output = result.stdout + result.stderr
        return result.returncode == 0, output.strip()
    except Exception as e:
        return False, str(e)


def load_command_items() -> list[dict]:
    if not COMMANDS_CONFIG.exists():
        return []
    with open(COMMANDS_CONFIG, "r", encoding="utf-8") as f:
        config = json.load(f)
    return [c for c in config.get("commands", []) if c.get("enabled", True)]


def apply_command_items(items: list[dict], log: Callable[[str], None], dry_run: bool = False) -> tuple[int, int]:
    success, fail = 0, 0
    for i, item in enumerate(items, 1):
        command = item["command"]
        cmd_type = item["type"]
        desc = item.get("description", command[:60])

        if dry_run:
            log(f"[시뮬레이션] [{i}/{len(items)}] {desc}")
            success += 1
            continue

        log(f"[실행중] [{i}/{len(items)}] {desc}")

        if cmd_type == "powershell":
            ok, output = run_powershell(command)
        else:
            ok, output = run_cmd(command)

        if ok:
            log(f"[완료] {desc}")
            success += 1
        else:
            err_msg = output[:100] if output else "알 수 없는 오류"
            log(f"[실패] {desc}: {err_msg}")
            fail += 1

    return success, fail
