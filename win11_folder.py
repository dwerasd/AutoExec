# -*- coding: utf-8 -*-
#
# 폴더 백업 및 복구 모듈
# - 증분 백업 (동일 파일 건너뛰기)
# - 서비스 중지/시작 지원
# - 메타데이터 기반 복구
#

import json
import os
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Callable


CONFIG_FILE = Path(__file__).parent / "folder_config.json"

DEFAULT_EXCLUDE_FILES = [
    "desktop.ini",
    "Thumbs.db",
    ".DS_Store",
]


def expand_path(path):
    return os.path.normpath(os.path.expandvars(os.path.expanduser(path)))


def normalize_path_item(item):
    if isinstance(item, str):
        return {"path": item, "service": None, "exclude": [], "destination": None}
    return {
        "path": item.get("path", ""),
        "service": item.get("service"),
        "exclude": item.get("exclude", []),
        "destination": item.get("destination")
    }


# ===== 서비스 관리 =====

def get_service_status(service_name: str) -> str | None:
    try:
        result = subprocess.run(
            ["sc", "query", service_name],
            capture_output=True, text=True, encoding="cp949"
        )
        if "RUNNING" in result.stdout:
            return "running"
        elif "STOPPED" in result.stdout:
            return "stopped"
        return "unknown"
    except Exception:
        return None


def stop_service(service_name: str, log: Callable[[str], None], timeout: int = 30) -> bool:
    log(f"[서비스] {service_name} 중지 중...")
    status = get_service_status(service_name)
    if status == "stopped":
        log(f"[서비스] {service_name} 이미 중지됨")
        return True
    if status is None:
        log(f"[서비스] {service_name} 서비스를 찾을 수 없음")
        return False

    try:
        subprocess.run(["sc", "stop", service_name], capture_output=True, encoding="cp949")
    except Exception as e:
        log(f"[서비스] 중지 명령 실패: {e}")
        return False

    for i in range(timeout):
        time.sleep(1)
        status = get_service_status(service_name)
        if status == "stopped":
            log(f"[서비스] {service_name} 중지 완료")
            return True

    log(f"[서비스] {service_name} 중지 타임아웃")
    return False


def start_service(service_name: str, log: Callable[[str], None], timeout: int = 30) -> bool:
    log(f"[서비스] {service_name} 시작 중...")
    status = get_service_status(service_name)
    if status == "running":
        log(f"[서비스] {service_name} 이미 실행 중")
        return True
    if status is None:
        log(f"[서비스] {service_name} 서비스를 찾을 수 없음")
        return False

    try:
        subprocess.run(["sc", "start", service_name], capture_output=True, encoding="cp949")
    except Exception as e:
        log(f"[서비스] 시작 명령 실패: {e}")
        return False

    for i in range(timeout):
        time.sleep(1)
        status = get_service_status(service_name)
        if status == "running":
            log(f"[서비스] {service_name} 시작 완료")
            return True

    log(f"[서비스] {service_name} 시작 타임아웃")
    return False


# ===== 설정 관리 =====

def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"backup_paths": [], "description": "백업할 폴더 경로 목록입니다.", "last_backup_destination": None}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def get_last_backup_destination() -> str | None:
    config = load_config()
    return config.get("last_backup_destination")


def save_last_backup_destination(destination):
    config = load_config()
    config["last_backup_destination"] = os.path.normpath(destination)
    save_config(config)


# ===== 복사 유틸 =====

def make_ignore_func(exclude_list=None):
    all_excludes = set(DEFAULT_EXCLUDE_FILES)
    if exclude_list:
        all_excludes.update(exclude_list)

    def ignore_func(directory, files):
        return [f for f in files if f in all_excludes]
    return ignore_func


def smart_copy2(src, dst):
    if os.path.exists(dst):
        src_stat = os.stat(src)
        dst_stat = os.stat(dst)
        if src_stat.st_size == dst_stat.st_size and int(src_stat.st_mtime) == int(dst_stat.st_mtime):
            return dst
    return shutil.copy2(src, dst)


# ===== 백업 =====

def backup(destination: str, log: Callable[[str], None]) -> bool:
    config = load_config()
    paths = config.get("backup_paths", [])

    if not paths:
        log("[백업] 백업할 경로가 등록되어 있지 않습니다.")
        return False

    backup_folder = os.path.normpath(destination)
    save_last_backup_destination(destination)

    try:
        os.makedirs(backup_folder, exist_ok=True)
    except Exception as e:
        log(f"[백업] 폴더 생성 실패: {e}")
        return False

    metadata = {"backup_date": datetime.now().isoformat(), "paths": []}

    # 기존 메타데이터 로드 (증분 백업용)
    meta_file = os.path.join(backup_folder, "backup_metadata.json")
    existing_backup_map = {}
    if os.path.exists(meta_file):
        try:
            with open(meta_file, "r", encoding="utf-8") as f:
                existing_meta = json.load(f)
                for item in existing_meta.get("paths", []):
                    existing_backup_map[item["source_expanded"]] = item["backup"]
        except Exception:
            pass

    log(f"[백업] 시작: {backup_folder}")

    success_count = 0
    fail_count = 0
    services_to_restart = []

    for item in paths:
        path_info = normalize_path_item(item)
        source_path = path_info["path"]
        service_name = path_info["service"]
        exclude_list = path_info["exclude"]
        custom_destination = path_info["destination"]

        expanded_path = expand_path(source_path)

        # 서비스 중지
        if service_name:
            original_status = get_service_status(service_name)
            if original_status == "running":
                if not stop_service(service_name, log):
                    log(f"[실패] 서비스 중지 실패로 건너뜀: {expanded_path}")
                    fail_count += 1
                    continue
                services_to_restart.append(service_name)

        if not os.path.exists(expanded_path):
            log(f"[건너뜀] 경로 없음: {expanded_path}")
            fail_count += 1
            continue

        # 백업 폴더명 결정
        if custom_destination:
            folder_name = custom_destination
        elif expanded_path in existing_backup_map:
            folder_name = existing_backup_map[expanded_path]
        else:
            base_name = os.path.basename(expanded_path)
            parent_name = os.path.basename(os.path.dirname(expanded_path))
            folder_name = f"{parent_name}_{base_name}"

        dest_path = os.path.join(backup_folder, folder_name)

        try:
            is_incremental = os.path.exists(dest_path)
            tag = "증분" if is_incremental else "백업"
            log(f"[{tag}] {expanded_path} -> {folder_name}")

            if os.path.isfile(expanded_path):
                smart_copy2(expanded_path, dest_path)
            else:
                shutil.copytree(
                    expanded_path, dest_path,
                    ignore=make_ignore_func(exclude_list),
                    dirs_exist_ok=True,
                    copy_function=smart_copy2
                )

            meta_item = {
                "source": source_path,
                "source_expanded": expanded_path,
                "backup": folder_name,
                "type": "file" if os.path.isfile(expanded_path) else "directory"
            }
            if service_name:
                meta_item["service"] = service_name
            if exclude_list:
                meta_item["exclude"] = exclude_list
            metadata["paths"].append(meta_item)

            log(f"[완료] {expanded_path}")
            success_count += 1
        except Exception as e:
            log(f"[실패] {expanded_path}: {e}")
            fail_count += 1

    # 서비스 재시작
    for svc in services_to_restart:
        start_service(svc, log)

    # 메타데이터 저장
    with open(os.path.join(backup_folder, "backup_metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=4)

    log(f"[백업] 완료: 성공 {success_count}개, 실패 {fail_count}개")
    return True


# ===== 복구 =====

def find_backup_root(path):
    path = os.path.normpath(path)
    if os.path.exists(os.path.join(path, "backup_metadata.json")):
        return path, None
    parent = os.path.dirname(path)
    folder_name = os.path.basename(path)
    if os.path.exists(os.path.join(parent, "backup_metadata.json")):
        return parent, folder_name
    return path, None


def restore(input_path: str, log: Callable[[str], None], target: str | None = None) -> bool:
    backup_path, auto_target = find_backup_root(input_path)

    if auto_target and not target:
        target = auto_target
        log(f"[자동 감지] '{target}' 폴더만 복구합니다.")

    if not os.path.exists(backup_path):
        log(f"[복구] 백업 폴더가 존재하지 않습니다: {backup_path}")
        return False

    metadata_file = os.path.join(backup_path, "backup_metadata.json")
    if not os.path.exists(metadata_file):
        log(f"[복구] 메타데이터 파일이 없습니다: {metadata_file}")
        return False

    with open(metadata_file, "r", encoding="utf-8") as f:
        metadata = json.load(f)

    all_items = metadata["paths"]
    if target:
        restore_items = [item for item in all_items if item["backup"] == target]
        if not restore_items:
            log(f"[복구] '{target}' 항목을 찾을 수 없습니다.")
            return False
    else:
        restore_items = all_items

    log(f"[복구] 시작: {len(restore_items)}개 항목 (백업일: {metadata.get('backup_date', '알 수 없음')})")

    # 서비스별 그룹화
    service_groups: dict[str | None, list] = {}
    for item in restore_items:
        svc = item.get("service")
        if svc not in service_groups:
            service_groups[svc] = []
        service_groups[svc].append(item)

    success_count = 0
    fail_count = 0

    for service_name, items in service_groups.items():
        service_was_running = False

        if service_name:
            status = get_service_status(service_name)
            if status == "running":
                service_was_running = True
                if not stop_service(service_name, log):
                    log(f"[경고] 서비스 {service_name} 중지 실패, 건너뜁니다.")
                    fail_count += len(items)
                    continue

        for item in items:
            source = item["source"]
            source_expanded = expand_path(source)
            backup_relative = item["backup"]
            item_type = item.get("type", "directory")

            backup_item_path = os.path.join(backup_path, backup_relative)

            if not os.path.exists(backup_item_path):
                log(f"[건너뜀] 백업 항목 없음: {backup_item_path}")
                fail_count += 1
                continue

            try:
                log(f"[복구중] {source_expanded}")

                if os.path.exists(source_expanded):
                    if os.path.isfile(source_expanded):
                        os.remove(source_expanded)
                    else:
                        shutil.rmtree(source_expanded)

                if item_type == "file":
                    os.makedirs(os.path.dirname(source_expanded), exist_ok=True)
                    shutil.copy2(backup_item_path, source_expanded)
                else:
                    shutil.copytree(backup_item_path, source_expanded)

                log(f"[완료] {source_expanded}")
                success_count += 1
            except Exception as e:
                log(f"[실패] {source_expanded}: {e}")
                fail_count += 1

        if service_name and service_was_running:
            start_service(service_name, log)

    log(f"[복구] 완료: 성공 {success_count}개, 실패 {fail_count}개")
    return True
