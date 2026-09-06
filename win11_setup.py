# -*- coding: utf-8 -*-
#
# Windows 11 시스템 설정 적용 모듈
# - 레지스트리 설정 적용
# - 인트라넷 영역 등록
# - 시스템 명령어 실행
#

import ctypes
import json
import os
import shutil
import socket
import subprocess
import sys
import winreg
from ctypes import wintypes
from pathlib import Path
from typing import Any, Callable


# 설정 파일 경로
SCRIPT_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
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


def save_registry_items(items: list[dict]):
    config = {"registry_items": items, "description": "백업할 레지스트리 항목 목록입니다."}
    with open(REGISTRY_CONFIG, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


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


# ===== 프로세스 생성 감사 (이벤트 4688) =====

# 감사 하위 범주 GUID (언어 무관 — 영문/한글 Windows 공통)
AUDIT_PROCESS_CREATION_GUID = "{0CCE922B-69AE-11D9-BED3-505054503030}"

# 4688 이벤트에 전체 명령줄 포함 여부 (순수 레지스트리 DWORD)
AUDIT_CMDLINE_PATH = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit"
AUDIT_CMDLINE_NAME = "ProcessCreationIncludeCmdLine_Enabled"


def set_process_creation_audit(enable: bool, log: Callable[[str], None]) -> bool:
    """프로세스 생성 감사(이벤트 4688)를 켜거나 끈다.

    두 부분으로 구성되며 한 동작에 함께 적용한다:
      1. 감사 정책 ON/OFF: ``auditpol.exe`` (LSA 감사 API). 최신 Windows는
         하위 범주 설정 강제가 기본값이라 레지스트리로 직접 제어 불가.
      2. 명령줄 포함 ON/OFF: ``ProcessCreationIncludeCmdLine_Enabled`` DWORD (레지스트리).
    둘 다 관리자 권한이 필요하다. 상태 확인 없이 지정 상태로 강제 적용한다.

    Args:
        enable: True면 켜기(success/failure 감사 + 명령줄 기록), False면 끄기.
        log: 로그 콜백 함수.

    Returns:
        정책·레지스트리 둘 다 성공하면 True.
    """
    state = "enable" if enable else "disable"  # auditpol 인자 값
    label = "켜기" if enable else "끄기"  # 로그 표시용 라벨
    log(f"[프로세스감사] {label} 시작")

    # ① 감사 정책 (auditpol — 레지스트리로 불가, 반드시 auditpol 사용)
    # 주의: GUID에 따옴표를 붙이면 cmd /c 전달 시 따옴표가 깨져 인자 인식 실패.
    #       GUID는 공백이 없으므로 따옴표 없이 전달한다.
    ok_pol, out = run_cmd(
        f"auditpol /set /subcategory:{AUDIT_PROCESS_CREATION_GUID} "
        f"/success:{state} /failure:{state}"
    )
    if ok_pol:
        log(f"[프로세스감사] 감사 정책 {label} 완료")
    else:
        log(f"[프로세스감사] 감사 정책 실패(관리자 권한 확인): {out[:100]}")

    # ② 명령줄 포함 (레지스트리 DWORD: 1=기록, 0=미기록)
    cmdline_val = 1 if enable else 0  # 명령줄 포함 플래그
    ok_reg, err = write_registry_value(
        AUDIT_CMDLINE_PATH, AUDIT_CMDLINE_NAME, cmdline_val, winreg.REG_DWORD
    )
    if ok_reg:
        log(f"[프로세스감사] 명령줄 포함 {label} 완료")
    else:
        log(f"[프로세스감사] 명령줄 포함 실패: {err}")

    success = ok_pol and ok_reg  # 전체 성공 여부
    log(f"[프로세스감사] {label} {'성공' if success else '일부 실패'}")
    return success


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


# ===== 마우스 커서 구성표(skinpack) 설치 =====

CURSOR_SCHEME_NAME = "skinpack"
CURSOR_SRC_DIR = SCRIPT_DIR / "cursors" / CURSOR_SCHEME_NAME
CURSOR_DST_DIR = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "Cursors" / CURSOR_SCHEME_NAME
CURSOR_KEY_PATH = r"HKEY_CURRENT_USER\Control Panel\Cursors"
CURSOR_SCHEMES_PATH = CURSOR_KEY_PATH + r"\Schemes"

# 레지스트리 값 이름 → 파일명. 순서 = 마우스 속성 "포인터" 탭 순서이며 Schemes 값의 콤마 구분 순서와 같다.
CURSOR_FILES: list[tuple[str, str]] = [
    ("Arrow", "Arrow.cur"),
    ("Help", "Help.cur"),
    ("AppStarting", "AppStarting.ani"),
    ("Wait", "Wait.ani"),
    ("Crosshair", "Cross.cur"),
    ("IBeam", "IBeam.cur"),
    ("NWPen", "Handwriting.cur"),
    ("No", "NO.cur"),
    ("SizeNS", "SizeNS.cur"),
    ("SizeWE", "SizeWE.cur"),
    ("SizeNWSE", "SizeNWSE.cur"),
    ("SizeNESW", "SizeNESW.cur"),
    ("SizeAll", "SizeAll.cur"),
    ("UpArrow", "UpArrow.cur"),
    ("Hand", "Hand.cur"),
]

# KB5120998(빌드 26100.9278)의 한국어 win32kbase.sys.mui 가 위 값 이름 6개를 번역해 그 이름으로
# 레지스트리를 조회한다(값이 없어 내장 흑백 커서로 폴백). 번역된 이름으로도 같은 값을 써 둔다.
# MS 가 고치면 영문 이름으로 돌아가므로 이 값들은 무시될 뿐이다.
CURSOR_KO_ALIASES: dict[str, str] = {
    "Arrow": "화살표",
    "Wait": "대기",
    "Crosshair": "십자선",
    "No": "아니요",
    "Hand": "손",
    "Help": "도움말",
}

# 검증용: 값 이름 → 시스템 커서 ID(IDC_*)
CURSOR_IDC: dict[str, int] = {
    "Arrow": 32512, "IBeam": 32513, "Wait": 32514, "Crosshair": 32515, "UpArrow": 32516,
    "SizeNWSE": 32642, "SizeNESW": 32643, "SizeWE": 32644, "SizeNS": 32645, "SizeAll": 32646,
    "No": 32648, "Hand": 32649, "AppStarting": 32650, "Help": 32651,
}

SPI_SETCURSORS = 0x0057
SPIF_SENDCHANGE = 0x0002


class _ICONINFO(ctypes.Structure):
    _fields_ = [
        ("fIcon", wintypes.BOOL),
        ("xHotspot", wintypes.DWORD),
        ("yHotspot", wintypes.DWORD),
        ("hbmMask", wintypes.HBITMAP),
        ("hbmColor", wintypes.HBITMAP),
    ]


def _system_cursor_is_color(idc: int) -> bool | None:
    """현재 시스템 커서가 컬러 비트맵을 가졌는지 반환. 흑백(폴백) 이면 False, 조회 실패면 None."""
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    user32.LoadCursorW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.LoadCursorW.restype = ctypes.c_void_p
    user32.GetIconInfo.argtypes = [ctypes.c_void_p, ctypes.POINTER(_ICONINFO)]
    user32.GetIconInfo.restype = wintypes.BOOL
    gdi32.DeleteObject.argtypes = [ctypes.c_void_p]

    hcur = user32.LoadCursorW(None, idc)  # 공유 핸들 — 해제 불필요
    if not hcur:
        return None
    info = _ICONINFO()
    if not user32.GetIconInfo(hcur, ctypes.byref(info)):
        return None
    is_color = bool(info.hbmColor)
    # GetIconInfo 는 비트맵 복사본을 돌려주므로 호출자가 지운다
    if info.hbmMask:
        gdi32.DeleteObject(info.hbmMask)
    if info.hbmColor:
        gdi32.DeleteObject(info.hbmColor)
    return is_color


def install_cursor_scheme(log: Callable[[str], None]) -> bool:
    """skinpack 커서를 시스템 폴더에 복사·구성표 등록·현재 커서로 적용하고 실제 반영을 검증한다.

    install.inf 우클릭 설치 + 마우스 속성 적용을 한 번에 하며, KB5120998 한글 이름 번역 버그까지 보정한다.
    파일 복사에 관리자 권한이 필요하다.
    """
    tag = "[커서설치]"

    # 1) 원본 확인
    missing = [f for _, f in CURSOR_FILES if not (CURSOR_SRC_DIR / f).exists()]
    if missing:
        log(f"{tag} 원본 누락 {CURSOR_SRC_DIR}: {', '.join(missing)}")
        return False

    # 2) 시스템 폴더로 복사 (%SystemRoot%\Cursors\skinpack)
    try:
        CURSOR_DST_DIR.mkdir(parents=True, exist_ok=True)
        for _, fname in CURSOR_FILES:
            shutil.copy2(CURSOR_SRC_DIR / fname, CURSOR_DST_DIR / fname)
        log(f"{tag} 파일 {len(CURSOR_FILES)}개 복사 완료 → {CURSOR_DST_DIR}")
    except PermissionError:
        log(f"{tag} 파일 복사 실패: 관리자 권한 필요")
        return False
    except OSError as e:
        log(f"{tag} 파일 복사 실패: {e}")
        return False

    # 3) 레지스트리 — 구성표 등록 + 현재 커서 15개 + 한글 이름 보정 6개
    def _reg_path(fname: str) -> str:
        return rf"%SystemRoot%\Cursors\{CURSOR_SCHEME_NAME}\{fname}"

    fail = 0
    scheme_value = ",".join(_reg_path(f) for _, f in CURSOR_FILES)
    ok, err = write_registry_value(CURSOR_SCHEMES_PATH, CURSOR_SCHEME_NAME, scheme_value, winreg.REG_EXPAND_SZ)
    if not ok:
        log(f"{tag} 구성표 등록 실패: {err}")
        fail += 1

    for name, fname in CURSOR_FILES:
        ok, err = write_registry_value(CURSOR_KEY_PATH, name, _reg_path(fname), winreg.REG_EXPAND_SZ)
        if not ok:
            log(f"{tag} 값 쓰기 실패 {name}: {err}")
            fail += 1
        alias = CURSOR_KO_ALIASES.get(name)
        if alias:
            ok, err = write_registry_value(CURSOR_KEY_PATH, alias, _reg_path(fname), winreg.REG_EXPAND_SZ)
            if not ok:
                log(f"{tag} 값 쓰기 실패 {alias}: {err}")
                fail += 1

    ok, err = write_registry_value(CURSOR_KEY_PATH, "", CURSOR_SCHEME_NAME, winreg.REG_SZ)  # (기본값) = 현재 구성표 이름
    if not ok:
        log(f"{tag} 기본값 쓰기 실패: {err}")
        fail += 1
    ok, err = write_registry_value(CURSOR_KEY_PATH, "Scheme Source", 1, winreg.REG_DWORD)  # 1 = 사용자 구성표
    if not ok:
        log(f"{tag} Scheme Source 쓰기 실패: {err}")
        fail += 1

    if fail:
        log(f"{tag} 레지스트리 {fail}건 실패 — 적용 중단")
        return False
    log(f"{tag} 레지스트리 등록 완료 (구성표 '{CURSOR_SCHEME_NAME}', 값 {len(CURSOR_FILES)}개 + 한글 보정 {len(CURSOR_KO_ALIASES)}개)")

    # 4) 즉시 반영 — 마우스 속성 "적용"과 같은 신호. SPIF_UPDATEINIFILE 은 이 환경에서 FALSE 를 돌려주므로 쓰지 않는다.
    user32 = ctypes.windll.user32
    user32.SystemParametersInfoW.argtypes = [wintypes.UINT, wintypes.UINT, ctypes.c_void_p, wintypes.UINT]
    user32.SystemParametersInfoW.restype = wintypes.BOOL
    if not user32.SystemParametersInfoW(SPI_SETCURSORS, 0, None, SPIF_SENDCHANGE):
        log(f"{tag} 커서 다시 읽기 신호 실패 (재로그인하면 반영됨)")

    # 5) 검증 — 흑백 폴백이 남은 슬롯이 있으면 실패로 본다
    mono = [name for name, idc in CURSOR_IDC.items() if _system_cursor_is_color(idc) is False]
    if mono:
        log(f"{tag} 적용 실패 — 흑백 폴백 남음: {', '.join(mono)}")
        return False
    log(f"{tag} 완료 — 커서 {len(CURSOR_IDC)}개 전부 컬러 적용 확인")
    return True
