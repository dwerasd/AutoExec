#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# GitHub 저장소 클론 스크립트
# 타인의 GitHub 저장소를 클론하고 구독 목록에 등록
#
# 사용법:
#     python gitclone.py owner/repo                    # 클론 + 구독 등록
#     python gitclone.py owner/repo --path "E:\dev"    # 경로 지정
#     python gitclone.py owner/repo --reset            # 삭제 후 재클론
#
# 구독 관리는 gitsync.py 사용:
#     python gitsync.py                # 모든 구독 저장소 업데이트
#     python gitsync.py --list         # 구독 목록 확인
#

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).parent
ENV_FILE = SCRIPT_DIR / ".env"
DATA_DIR = SCRIPT_DIR / "data"
REPOS_FILE = DATA_DIR / "repos.json"


def load_config() -> dict:
    """.env 파일에서 설정 로드"""
    config = {
        "GITHUB_USER": "",
        "GITHUB_TOKEN": "",
        "CLONE_BASE_PATH": ""
    }
    
    if not ENV_FILE.exists():
        print(f"오류: 설정 파일이 없습니다: {ENV_FILE}")
        sys.exit(1)
    
    with open(ENV_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and "=" in line and not line.startswith("#"):
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                if key in config:
                    config[key] = value
    
    return config


def load_repos() -> dict:
    """repos.json 파일 로드"""
    if not REPOS_FILE.exists():
        return {"subscriptions": []}
    
    try:
        with open(REPOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"subscriptions": []}


def save_repos(data: dict):
    """repos.json 파일 저장"""
    # data 폴더가 없으면 생성
    if not DATA_DIR.exists():
        DATA_DIR.mkdir(parents=True)
    with open(REPOS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def find_subscription(repos_data: dict, owner: str, repo_name: str) -> dict | None:
    """구독 목록에서 저장소 찾기"""
    repo_full = f"{owner}/{repo_name}"
    for sub in repos_data.get("subscriptions", []):
        if sub.get("repo") == repo_full:
            return sub
    return None


def add_subscription(owner: str, repo_name: str, local_path: str, branch: str = "main"):
    """구독 목록에 저장소 추가"""
    repos_data = load_repos()
    repo_full = f"{owner}/{repo_name}"
    
    # 이미 존재하는지 확인
    existing = find_subscription(repos_data, owner, repo_name)
    if existing:
        # 경로 업데이트
        existing["local_path"] = local_path
        existing["branch"] = branch
    else:
        # 새로 추가
        repos_data["subscriptions"].append({
            "repo": repo_full,
            "owner": owner,
            "name": repo_name,
            "local_path": local_path,
            "branch": branch,
            "added": datetime.now().strftime("%Y-%m-%d"),
            "last_commit": "",
            "auto_update": True
        })
    
    save_repos(repos_data)


def update_last_commit(owner: str, repo_name: str, commit_sha: str):
    """마지막 커밋 SHA 업데이트"""
    repos_data = load_repos()
    sub = find_subscription(repos_data, owner, repo_name)
    if sub:
        sub["last_commit"] = commit_sha
        save_repos(repos_data)


def parse_repo_input(repo_input: str) -> tuple[str, str]:
    """
    다양한 형식의 입력을 owner, repo로 파싱
    
    지원 형식:
    - owner/repo
    - https://github.com/owner/repo
    - https://github.com/owner/repo.git
    - https://github.com/owner/repo/
    - https://github.com/owner/repo/tree/main/...
    - https://github.com/owner/repo?tab=readme-ov-file
    - git@github.com:owner/repo.git
    """
    repo_input = repo_input.strip()
    
    # 쿼리 파라미터 제거 (?tab=..., ?branch=... 등)
    repo_input = re.sub(r'\?.*$', '', repo_input)
    
    # URL 형식: https://github.com/owner/repo (다양한 후속 경로 허용)
    https_match = re.match(r'https?://github\.com/([^/]+)/([^/]+?)(?:\.git)?(?:/.*)?$', repo_input)
    if https_match:
        owner = https_match.group(1)
        repo = https_match.group(2)
        # .git 제거 (혹시 남아있으면)
        repo = repo.removesuffix('.git')
        return owner, repo
    
    # SSH 형식: git@github.com:owner/repo.git
    ssh_match = re.match(r'git@github\.com:([^/]+)/([^/]+?)(?:\.git)?$', repo_input)
    if ssh_match:
        return ssh_match.group(1), ssh_match.group(2)
    
    # 단순 형식: owner/repo
    simple_match = re.match(r'^([^/]+)/([^/]+)$', repo_input)
    if simple_match:
        return simple_match.group(1), simple_match.group(2)
    
    print(f"오류: 올바른 저장소 형식이 아닙니다: {repo_input}")
    print("지원 형식:")
    print("  - owner/repo")
    print("  - https://github.com/owner/repo")
    print("  - git@github.com:owner/repo.git")
    sys.exit(1)


def run_git(args: list[str], cwd: str | None = None) -> tuple[bool, str]:
    """Git 명령 실행"""
    try:
        result = subprocess.run(
            ["git"] + args,
            cwd=cwd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        output = (result.stdout + result.stderr).strip()
        return result.returncode == 0, output
    except Exception as e:
        return False, str(e)


def get_local_commit(repo_path: str) -> str | None:
    """로컬 저장소의 현재 HEAD 커밋 SHA"""
    success, output = run_git(["rev-parse", "HEAD"], repo_path)
    return output if success else None


def get_unique_path(base_path: str, repo_name: str, owner: str) -> str:
    """
    중복되지 않는 클론 경로 생성
    
    새 구조: base_path/owner/repo_name
    """
    # 기본 경로: base_path/owner/repo_name
    target_path = os.path.join(base_path, owner, repo_name)
    if not os.path.exists(target_path):
        return target_path
    
    # 기존 폴더가 같은 owner의 저장소인지 확인
    existing_origin = get_remote_origin(target_path)
    if existing_origin and f"/{owner}/{repo_name}" in existing_origin.lower():
        # 같은 저장소면 그대로 사용
        return target_path
    
    # 중복 시 숫자 추가: base_path/owner/repo_name_2, repo_name_3, ...
    counter = 2
    while True:
        target_path = os.path.join(base_path, owner, f"{repo_name}_{counter}")
        if not os.path.exists(target_path):
            return target_path
        counter += 1
        if counter > 100:
            print("오류: 너무 많은 중복 폴더가 존재합니다.")
            sys.exit(1)


def get_remote_origin(repo_path: str) -> str | None:
    """기존 저장소의 origin URL 확인"""
    if not os.path.exists(os.path.join(repo_path, ".git")):
        return None
    
    success, output = run_git(["remote", "get-url", "origin"], repo_path)
    return output if success else None


def get_default_branch(repo_path: str) -> str:
    """기본 브랜치 이름 확인"""
    success, output = run_git(["symbolic-ref", "refs/remotes/origin/HEAD", "--short"], repo_path)
    if success:
        return output.replace("origin/", "")
    return "main"


def clone_repository(repo_input: str, base_path: str | None = None, reset: bool = False) -> bool:
    """저장소 클론 + 구독 등록"""
    
    # 설정 로드
    config = load_config()
    token = config.get("GITHUB_TOKEN", "")
    my_user = config.get("GITHUB_USER", "")
    
    # 저장소 정보 파싱
    owner, repo_name = parse_repo_input(repo_input)
    
    # 내 저장소인지 확인
    if my_user and owner.lower() == my_user.lower():
        print(f"경고: '{owner}/{repo_name}'은 본인의 저장소입니다.")
        print("  본인 저장소는 구독 대상이 아닙니다.")
        print("  계속하시겠습니까? (y/N): ", end="")
        response = input().strip().lower()
        if response != 'y':
            print("취소되었습니다.")
            return False
    
    # 기본 경로 결정
    if base_path:
        clone_base = os.path.abspath(base_path)
    elif config.get("CLONE_BASE_PATH"):
        clone_base = config["CLONE_BASE_PATH"]
    else:
        # 기본값: 스크립트 폴더의 data/
        clone_base = str(SCRIPT_DIR / "data")
    
    # 기본 경로가 없으면 생성
    if not os.path.exists(clone_base):
        os.makedirs(clone_base)
    
    # 중복 방지 경로 결정
    target_path = get_unique_path(clone_base, repo_name, owner)
    
    print(f"\n{'='*60}")
    print(f" GitHub 저장소 클론")
    print(f"{'='*60}")
    print(f"  저장소: {owner}/{repo_name}")
    print(f"  대상 경로: {target_path}")
    print(f"  토큰 사용: {'예' if token else '아니오 (공개 저장소만 가능)'}")
    print()
    
    # 기존 폴더 처리
    if os.path.exists(target_path):
        if reset:
            print(f"[1/4] 기존 폴더 삭제 중...")
            try:
                shutil.rmtree(target_path)
                print("  삭제 완료")
            except Exception as e:
                print(f"  오류: 삭제 실패 - {e}")
                return False
        else:
            print(f"오류: 대상 경로가 이미 존재합니다: {target_path}")
            print("  --reset 옵션으로 삭제 후 재클론 가능")
            return False
    else:
        print(f"[1/4] 경로 확인...")
        print("  새 폴더 생성 예정")
    
    # 상위 디렉토리 생성
    parent_dir = os.path.dirname(target_path)
    if not os.path.exists(parent_dir):
        os.makedirs(parent_dir)
    
    # Clone URL 구성
    print(f"[2/4] 저장소 클론 중...")
    if token:
        clone_url = f"https://{token}@github.com/{owner}/{repo_name}.git"
    else:
        clone_url = f"https://github.com/{owner}/{repo_name}.git"
    
    success, output = run_git(["clone", clone_url, target_path])
    
    if not success:
        print(f"  오류: 클론 실패")
        print(f"  {output}")
        print()
        print("가능한 원인:")
        print("  1. 저장소가 존재하지 않음")
        print("  2. private 저장소인데 토큰이 없거나 권한 부족")
        print("  3. 네트워크 연결 문제")
        return False
    
    print("  클론 완료")
    
    # 토큰이 포함된 URL을 일반 URL로 변경 (보안)
    print(f"[3/4] URL 정리...")
    if token:
        clean_url = f"https://github.com/{owner}/{repo_name}.git"
        run_git(["remote", "set-url", "origin", clean_url], target_path)
        print("  원격 URL 정리 완료 (토큰 제거)")
    else:
        print("  완료")
    
    # 구독 등록
    print(f"[4/4] 구독 등록...")
    branch = get_default_branch(target_path)
    add_subscription(owner, repo_name, target_path, branch)
    
    # 현재 커밋 SHA 저장
    current_commit = get_local_commit(target_path)
    if current_commit:
        update_last_commit(owner, repo_name, current_commit)
    
    print(f"  구독 등록 완료 (브랜치: {branch})")
    
    # 결과 출력
    print(f"\n{'='*60}")
    print(f" 완료!")
    print(f"{'='*60}")
    print(f"  저장소: https://github.com/{owner}/{repo_name}")
    print(f"  로컬 경로: {target_path}")
    print(f"  브랜치: {branch}")
    print()
    print("다음 명령으로 업데이트 확인:")
    print(f"  python gitsync.py")
    print()
    
    # 클론된 내용 표시
    print("클론된 내용:")
    try:
        items = os.listdir(target_path)
        if items:
            for item in sorted(items)[:15]:
                item_path = os.path.join(target_path, item)
                if os.path.isdir(item_path):
                    print(f"  📁 {item}/")
                else:
                    print(f"  📄 {item}")
            if len(items) > 15:
                print(f"  ... 외 {len(items) - 15}개")
        else:
            print("  (빈 저장소)")
    except Exception:
        print("  (내용 확인 실패)")
    
    print()
    return True


def main():
    parser = argparse.ArgumentParser(
        description="GitHub 저장소 클론 + 구독 등록",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python gitclone.py microsoft/vscode              # 클론 + 구독 등록
  python gitclone.py microsoft/vscode --path "E:\\dev"  # 경로 지정
  python gitclone.py microsoft/vscode --reset      # 삭제 후 재클론

구독 관리 (gitsync.py 사용):
  python gitsync.py                # 모든 구독 저장소 업데이트
  python gitsync.py --list         # 구독 목록 확인
  python gitsync.py --remove owner/repo  # 구독 해제

설정:
  .env 파일에 CLONE_BASE_PATH를 설정하면 기본 클론 경로로 사용됩니다.
  예: CLONE_BASE_PATH=E:\\GitHub\\clones
        """
    )
    
    parser.add_argument(
        "repo",
        help="저장소 (owner/repo 또는 GitHub URL)"
    )
    
    parser.add_argument(
        "--path", "-p",
        help="클론할 기본 경로 (미지정시 CLONE_BASE_PATH 또는 현재 디렉토리)"
    )
    
    parser.add_argument(
        "--reset",
        action="store_true",
        help="기존 폴더가 있으면 삭제 후 재클론"
    )
    
    args = parser.parse_args()
    
    success = clone_repository(args.repo, args.path, args.reset)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
