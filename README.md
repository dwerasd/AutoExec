# AutoExec

Windows PC 관리, WOL 원격 부팅, 자동실행 스케줄러를 하나로 통합한 데스크톱 애플리케이션입니다.

## 주요 기능

### PC 관리 & WOL 부팅
- PC 목록 관리 (이름, IP, MAC 주소)
- Wake-on-LAN 매직 패킷 전송으로 원격 부팅
- 지정 시간에 자동 WOL 부팅 (시작/종료 시간 설정)
- Ping 응답 확인
- 부팅 실패 시 자동 재시도 (2회) 및 텔레그램 알림

### 자동실행 스케줄러
- 실행 파일(.exe), Python 스크립트(.py/.pyw), 배치 파일(.bat) 지원
- Python 가상환경 경로 지정 가능
- 다양한 실행 모드:
  - **매일 1회**: 지정 시간에 하루 1회 실행
  - **N분/N시간 간격**: 시작~종료 시간 내에서 반복 실행
  - **요일 지정**: 선택한 요일에만 실행
  - **매월 지정일**: 매월 특정 날짜에 실행
- 중복 실행 방지
- 실행 중인 작업 강제 중지

### 휴장일 관리
- 공휴일/휴장일 데이터베이스 내장
- PC별/작업별 휴장일 제외 옵션
- 토/일요일 자동 인식

### 기타
- GitHub 저장소 URL 붙여넣기로 빠른 다운로드
- 시스템 트레이 최소화 (닫기 버튼 = 트레이로 이동)
- 최상위 윈도우 고정 옵션
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
├── AutoExec.db           # SQLite3 데이터베이스 (pcs, tasks, closed_days)
├── AutoExec.json         # 로컬 UI 설정 (윈도우 위치/크기)
├── .env                  # 텔레그램 봇 설정
├── requirements.txt      # Python 의존 패키지
├── migrate_to_sqlite.py  # MariaDB → SQLite 마이그레이션 스크립트 (1회용)
└── README.md
```

## 설정

### 텔레그램 알림 (선택 사항)

WOL 부팅 실패 시 텔레그램으로 알림을 받을 수 있습니다.
사용하려면 프로젝트 루트에 `.env` 파일을 생성하고 아래 내용을 작성합니다.

```env
TELEGRAM_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_chat_id
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

## 스크린샷

### 메인 화면 구성

| 영역 | 설명 |
|------|------|
| 상단 | 최상위 고정 체크박스, GitHub URL 입력 |
| 자동실행 | 등록된 작업 목록 (추가/편집/삭제/실행/중지/순서변경) |
| 로그 | 실행 기록 및 상태 메시지 |
| PC 관리 | PC 목록 (편집/부팅/핑/추가/삭제/순서변경) |

## 마이그레이션 (MariaDB → SQLite)

기존 MariaDB 환경에서 전환하는 경우:

```bash
python migrate_to_sqlite.py
```

이 스크립트는 MariaDB의 `pcs`, `tasks` 테이블과 `closed_days.db` 파일의 데이터를
`AutoExec.db` 하나로 통합합니다.

## License

MIT
