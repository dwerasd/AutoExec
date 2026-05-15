# AutoExec - Project Protocol

## 1. 프로젝트 개요
- **AutoExec.pyw** — PC 관리 / WOL 부팅 / 자동실행 스케줄러
- Python 3.14, tkinter + SQLite3 + JSON(로컬 UI 설정)
- Windows 전용 데스크탑 애플리케이션 (.pyw → pythonw.exe)

## 2. 프로젝트 구조
```
AutoExec.pyw        # 메인 애플리케이션 (단일 파일)
win11_setup.py      # Windows 11 초기 설정 유틸리티
win11_folder.py     # Windows 11 폴더 설정 유틸리티
gitclone.py         # GitHub 레포지토리 클론 헬퍼
gitsync.py          # Git 동기화 헬퍼
AutoExec.db         # SQLite3 데이터베이스 (pcs, tasks, move_targets, closed_days)
AutoExec.json       # 로컬 UI 설정 (창 위치, 크기 등)
.env                # 텔레그램 봇 토큰, 채팅 ID
data/               # Git 다운로드 저장 폴더
```

## 3. Communication
- **Language**: 모든 주석, 로그 메시지는 **한국어**.
- 서론/결론 생략. 기술적 사실과 코드 패치만 압축 전달.

## 4. Code Standards
- **Encoding**: UTF-8, 상단에 `# -*- coding: utf-8 -*-` 명시.
- **Naming**: 함수/변수 `snake_case`, 클래스 `PascalCase`, 상수 `SCREAMING_SNAKE_CASE`.
- **Exception**: bare `except:` 지양. 가능하면 구체적 예외 처리.
- **Import**: 외부 라이브러리 import 시 `# pip install [package]` 주석 기재 (표준 라이브러리 제외).
- **Win32 API**: 프로세스/창 관련 작업은 PowerShell/wmic 외부 호출 대신 ctypes Win32 API 직접 호출 우선 (성능).

## 5. 코드 검증 (필수)
코드 수정 후 반드시 아래 검증을 수행한다. **이 규칙은 협상 불가.**

### 5.1 구문 검사
```bash
python -c "import py_compile; py_compile.compile('AutoExec.pyw', doraise=True)"
```
- 모든 코드 변경 후 반드시 실행.
- 오류 발생 시 즉시 수정.

### 5.2 Pylance 경고 검토
- 사용자가 Pylance 경고를 전달하면 즉시 검토 및 수정.
- 주요 검토 대상:
  - `reportOptionalOperand`: `None` 가능 값에 대한 연산자 사용 → null 체크 추가
  - `reportArgumentType`: 타입 불일치 → 명시적 변환 또는 `type: ignore[assignment]`
  - `reportGeneralClassIssues`: 클래스 정의 문제
- `type: ignore` 사용 시 반드시 구체적 에러 코드 명시 (예: `type: ignore[assignment]`)

## 6. 데이터베이스
- SQLite3, `AutoExec.db`
- 테이블: `pcs`, `tasks`, `move_targets`, `closed_days`
- 스키마 마이그레이션: `db_init()` 내 `ALTER TABLE ... ADD COLUMN` + `try/except` 패턴
- 새 컬럼 추가 시 기존 DB 호환성을 위해 반드시 마이그레이션 코드 포함

## 7. Environment (Windows 11)
- Python 3.14 (`C:\Program Files\Python314\`)
- 가상환경 미사용 (시스템 Python 직접 사용)
- wmic 미설치 — 프로세스 조회 시 Win32 API 또는 PowerShell 사용
- 실행: `pythonw.exe AutoExec.pyw` (콘솔 창 없이 실행)

## 8. WOL 환경 의존성 (중요)
실행 PC에 다수의 가상 NIC가 공존한다: **VMware VMnet1/VMnet8**, **vEthernet (WSL Hyper-V firewall)**, 실제 LAN(Realtek 2.5GbE).
이로 인한 WOL 동작 제약을 코드가 직접 회피해야 한다. 동일 영역 수정 시 다음 규칙 유지.

### 8.1 매직 패킷 송신 (`send_wol`)
- **반드시** `target_ip`와 같은 /24 서브넷의 로컬 NIC IP에 socket을 `bind()` 후 송신. 가상 NIC가 limited broadcast를 가로채는 환경 우회.
- 송신 타깃 3종 + 포트 2종 동시 송출 (총 6회):
  - `255.255.255.255` (limited broadcast)
  - 서브넷 directed broadcast (예: `192.168.0.255`)
  - 유니캐스트 `target_ip` (NIC sleep proxy 응답 대비)
  - 포트: 9, 7
- `OSError`는 사유 문자열로 변환해 반환. 호출부에서 로그 기록 의무.

### 8.2 핑 응답 검증 (`ping_host`)
- Windows `ping`은 ICMP "Destination host unreachable" 응답 시에도 **returncode 0**을 반환한다(거짓 양성).
- 따라서 `returncode == 0` **그리고** stdout에 `TTL=` 토큰 존재할 때만 성공으로 판정한다.
- 이 검증을 풀어버리면 꺼진 PC가 1초 만에 "부팅 완료"로 잘못 표시되어 WOL 재시도 로직이 무력화된다.

### 8.3 알려진 시스템 작업 이력
- VMware 게스트 연결 끊김 회피 목적으로 Hyper-V/VBS/Credential Guard 비활성화(`bcdedit`, `reg`, `Disable-WindowsOptionalFeature`)가 부분 적용된 상태. vEthernet(WSL) NIC는 잔존 → broadcast 라우팅에 영향.
- WOL 회귀 발생 시 가장 먼저 의심: ① vEthernet NIC 상태, ② Windows 미리보기 누적 업데이트.
