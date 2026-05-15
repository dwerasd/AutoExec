# AutoExec

Windows PC 관리, WOL 원격 부팅, 자동실행 스케줄러, Windows 시스템 설정, 폴더 백업/복구를 하나로 통합한 데스크톱 애플리케이션입니다.

## 주요 기능

### PC 관리 & WOL 부팅
- PC 목록 관리 (이름, IP, MAC 주소)
- Wake-on-LAN 매직 패킷 전송으로 원격 부팅
- 지정 시간에 자동 WOL 부팅 (시작/종료 시간 설정)
- Ping 응답 확인
- 부팅 실패 시 자동 재시도 (2회) 및 텔레그램 알림

### 자동실행 스케줄러
- 실행 파일(.exe), Python 스크립트(.py/.pyw), 배치 파일(.bat) 지원
- **폴더 바로가기**: 자주 사용하는 폴더를 등록하고 더블클릭으로 열기
- Python 가상환경 경로 지정 가능
- 다양한 실행 모드:
  - **매일 1회**: 지정 시간에 하루 1회 실행
  - **부팅시 1회**: Windows 부팅 후 10분 이내 앱 시작 시 1회 자동 실행
  - **N분/N시간 간격**: 시작~종료 시간 내에서 반복 실행
  - **요일 지정**: 선택한 요일에만 실행
  - **매월 지정일**: 매월 특정 날짜에 실행
- 중복 실행 방지
- 실행 중인 작업 강제 중지
- 휴장일 제외 옵션 (기본값: 해제)

#### 폴더 등록 방법
1. 자동실행 목록에서 **추가** 클릭
2. 작업 이름 입력 후 실행 파일 옆 **폴더** 버튼으로 폴더 선택
3. 실행 모드를 **부팅시 1회**로 설정하면 앱 시작 시 자동으로 폴더 열기
4. 탐색기 창 위치를 지정하려면:
   - 해당 폴더를 미리 탐색기로 열어 원하는 위치/크기에 배치
   - 다이얼로그에서 **위치 캡처** 버튼 클릭 → X, Y, W, H 자동 입력
5. 위치 지정 없이(모두 0) 등록하면 기본 위치에서 열림
6. 등록된 폴더는 리스트에서 더블클릭으로도 열 수 있음

### 휴장일 관리
- 공휴일/휴장일 데이터베이스 내장
- PC별/작업별 휴장일 제외 옵션
- 토/일요일 자동 인식

### 듀얼 모니터 창 이동
- 모니터가 1개→2개로 변경되면 등록된 프로세스의 창을 자동 이동
- 프로세스별 이동 위치 설정:
  - **서브 모니터**: 서브 모니터로 이동 후 최대화
  - **현재 위치 저장**: 현재 창 위치를 캡처하여 저장, 이후 저장된 위치로 복원
  - **좌표 직접 입력**: X, Y 좌표 지정 (W,H=0이면 현재 크기 유지)
- 프로세스별 자동 이동 사용/미사용 설정
- 최소화된 창은 이동하지 않음
- 수동 이동 버튼으로 즉시 이동 가능

### Windows 시스템 설정 (관리 메뉴)
- **레지스트리 관리**: 레지스트리 항목 추가/편집/삭제, 더블클릭으로 개별 적용, 전체 적용
- **인트라넷 등록**: 로컬 IP 대역을 인트라넷 영역에 자동 등록

### 폴더 백업 & 복구 (관리 메뉴)
- **백업 경로 관리**: 백업 대상 폴더 추가/삭제
- **폴더 백업**: 등록된 경로를 증분 백업 (동일 파일 건너뛰기)
- **폴더 복구**: 메타데이터 기반 원본 경로로 복구
- 서비스 중지/시작 지원 (DB 등 잠긴 파일 백업 시)
- 제외 파일/폴더 지정 가능

### 기타
- GitHub 저장소 URL 붙여넣기로 빠른 다운로드
- 시스템 트레이 최소화 (닫기 버튼 = 트레이로 이동)
- 최상위 윈도우 고정 옵션 (설정 메뉴)
- 윈도우 위치/크기 자동 저장
- 중복 실행 방지 (Windows Named Mutex)

## 요구 사항

- **Python** 3.10+
- **OS**: Windows

### Python 패키지

```bash
pip install -r requirements.txt
```

## 파일 구조

```
AutoExec/
├── AutoExec.pyw          # 메인 애플리케이션
├── win11_setup.py        # 레지스트리 적용 + 인트라넷 등록 모듈
├── win11_folder.py       # 폴더 백업/복구 모듈
├── registry_config.json  # 레지스트리 설정 항목 (공개)
├── commands_config.json  # 시스템 명령어 목록 (공개)
├── folder_config.json    # 백업 경로 목록 (Git 미추적, 개인 경로 포함)
├── AutoExec.db           # SQLite3 데이터베이스 (pcs, tasks, move_targets, closed_days)
├── AutoExec.json         # 로컬 UI 설정 (윈도우 위치/크기, Git 미추적)
├── .env                  # 텔레그램 봇 + GitHub 토큰 설정 (Git 미추적)
├── requirements.txt      # Python 의존 패키지
├── gitclone.py           # GitHub 저장소 클론 + 구독 등록
├── gitsync.py            # 구독 저장소 자동 동기화
├── data/                 # 구독 저장소 데이터 (Git 미추적)
│   └── repos.json        # 구독 목록
└── README.md
```

## 설정

### 텔레그램 알림 (선택 사항)

WOL 부팅 실패 시 텔레그램으로 알림을 받을 수 있습니다.
사용하려면 프로젝트 루트에 `.env` 파일을 생성하고 아래 내용을 작성합니다.

```env
TELEGRAM_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_chat_id

GITHUB_USER=your_github_username
GITHUB_TOKEN=your_github_token

# gitclone.py 기본 클론 경로 (예: E:\GitHub\clones)
CLONE_BASE_PATH=
```

#### 텔레그램 봇 토큰 발급 방법

1. 텔레그램에서 [@BotFather](https://t.me/BotFather)를 검색하여 대화를 시작합니다.
2. `/newbot` 명령을 입력하고 봇 이름과 사용자명을 설정합니다.
3. 발급된 토큰을 `TELEGRAM_TOKEN`에 입력합니다.

#### Chat ID 확인 방법

1. 생성한 봇에게 아무 메시지를 보냅니다.
2. 브라우저에서 `https://api.telegram.org/bot<토큰>/getUpdates`에 접속합니다.
3. 응답 JSON의 `result[0].message.chat.id` 값을 `TELEGRAM_CHAT_ID`에 입력합니다.

> `.env` 파일이 없거나 값이 비어 있으면 텔레그램 알림 없이 정상 동작합니다.

## 실행

```bash
pythonw AutoExec.pyw
```

또는 Windows 시작프로그램 폴더에 바로가기를 등록하여 자동 시작할 수 있습니다.
앱 내 "시작폴더" 버튼으로 해당 폴더를 열 수 있습니다.

## GitHub 저장소 클론 & 동기화

### 저장소 클론 (gitclone.py)
```bash
python gitclone.py owner/repo                    # 클론 + 구독 등록
python gitclone.py https://github.com/owner/repo # URL도 가능
python gitclone.py owner/repo --path "E:\dev"    # 경로 지정
python gitclone.py owner/repo --reset            # 삭제 후 재클론
```

GUI에서 GitHub URL을 붙여넣으면 자동으로 `gitclone.py`를 호출합니다.

### 구독 저장소 동기화 (gitsync.py)
```bash
python gitsync.py                          # 모든 구독 저장소 업데이트
python gitsync.py --list                   # 구독 목록 확인
python gitsync.py --remove owner/repo      # 구독 해제
```

AutoExec 자동실행에 `gitsync.py`를 등록하면 구독 저장소를 주기적으로 업데이트할 수 있습니다.

## 라이선스

이 프로젝트는 MIT License에 따라 배포됩니다.

---

*이 프로젝트는 Claude Opus 4.6을 활용하여 작성되었습니다.*
