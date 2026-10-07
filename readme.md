# 시프티 (직원 시프트 관리 도구)

엑셀 스타일의 직원 시프트 관리 기본틀입니다.  
정적 배포(GitHub Pages / Netlify) + Supabase 연동을 전제로 구성되었습니다.

## 목적

- 팀/역할 기반 수정 권한 제어
- 근무 코드 입력(관리자 코드 관리)
- 한국 공휴일 표시
- 엑셀 Export, OCC(%) Excel Import

## 권한 정책

- `외식사업 2팀(team2)`: 2팀 수정 가능, 3팀 읽기 전용
- `외식사업 3팀(team3)`: 3팀 수정 가능, 2팀 읽기 전용
- `관리자(admin)`: 전체 수정 가능 + 근무 코드 생성 가능

## 폴더 구조

- `index.html`: 로그인 페이지 (진입점)
- `app.html`: 메인 화면 (세션 없으면 index.html로 리다이렉트)
- `js/login.js`: 로그인 페이지 로직
- `styles/main.css`: 엑셀 스타일 UI
- `js/main.js`: 앱 부트스트랩/이벤트 조합
- `js/auth.js`: 로그인 게이트 + 역할 상태 관리
- `js/accounts.js`: 로그인 계정(SHA-256 해시) 목록
- `js/permissions.js`: 역할별 수정 권한
- `js/scheduleGrid.js`: 시프트 테이블 렌더링
- `js/shiftCodes.js`: 근무 코드 검증/추가
- `js/holidayService.js`: 한국 공휴일 처리
- `js/occImport.js`: OCC 엑셀 Import
- `js/excelService.js`: 엑셀 Export
- `js/supabaseClient.js`: Supabase I/O 골격
- `netlify.toml`: Netlify 설정

## 로컬 실행

정적 서버로 실행하면 됩니다.

```bash
python3 -m http.server 5500
```

브라우저에서 `http://localhost:5500` 접속

주의: `index.html`을 `file://`로 직접 열면 모듈 스크립트가 동일 출처 정책(CORS)으로 차단됩니다.

## Supabase 연결 방법

`index.html`의 모듈 스크립트 실행 전에 전역 변수를 선언하면 연결됩니다.

```html
<script>
  window.SUPABASE_URL = "https://YOUR_PROJECT.supabase.co";
  window.SUPABASE_ANON_KEY = "YOUR_ANON_KEY";
</script>
```

정규화 테이블 및 RLS 정책 포함 전체 가이드는 `SUPABASE_SETUP.md`를 참고하세요.

## OCC Import 포맷

- 파일명에 `M000005326` 포함 필요
- 첫 번째 시트 사용
- `17행 C열부터`: 객실 수(1일 시작)
- `18행 C열부터`: OCC(%) (1일 시작)
- 기준월의 일수만큼 읽음 (예: 28일이면 AD열까지)

## 포지션 목록 (팀별)

- 외식사업 2팀: Assistant Manager / Team Leader / Senior Staff / Junior Staff / Internship / Part-Time / Manager
- 외식사업 3팀: Executive Chef / Chef de Cuisine / Sous Chef / Chef de Partie / Cook

`js/config.js`의 `TEAM2_POSITIONS`, `TEAM3_POSITIONS`에서 수정합니다.
기존 직원의 포지션이 새 목록에 없으면 드롭다운에 그대로 남겨두니 편집 시 새 포지션으로 바꿔 주세요.

## 기본 근무 코드 규칙

- `WKL`: 휴무
- `ON`: 오피스_공통 (`09:00 - 18:00`)
- `PR`: 오피스_파크로쉬 (`09:00 - 18:00`)
- `OAK`: 오피스_오크밸리 (`09:00 - 18:00`)
- `AL`: 연차
- `A~J`: 시간대 코드 (`A=05:00-14:00`, `...`, `J=14:00-23:00`)
- `A3` 같은 연장근무 표기 허용 (`A~J` + 숫자 접미사)
- `BAR (I)` 같은 BAR 파생 표기 허용

## 로그인 계정

`index.html`(로그인 페이지)에서 인증 후 `app.html`로 이동합니다. `app.html`은 세션이 없으면 본문을 그리기 전에 로그인 페이지로 되돌립니다. 계정은 `js/accounts.js`에 SHA-256 해시로 저장되며,
파일 상단 주석의 명령으로 해시를 생성해 추가/변경합니다.

- `admin` 등급: 모든 역할(관리자 포함) 선택 가능
- `guest` 등급: 외식사업 2팀/3팀 역할만 선택 가능

로그인 상태는 브라우저 탭 세션에 12시간 유지됩니다(`SESSION_TTL_MS`).

> 참고: 정적 호스팅이라 이 로그인은 클라이언트 측 차단입니다. 무단 접근을 막는 1차 장치로는 충분하지만,
> 데이터 자체를 보호하려면 Supabase RLS 정책(`supabase_setup.md`)이 함께 적용되어 있어야 합니다.

## GitHub Pages 배포

1. 이 폴더 내용을 레포 루트에 push
2. 레포 Settings → Pages → Source: `Deploy from a branch`, Branch: `main` / `/ (root)`
3. 1~2분 후 `https://<계정>.github.io/<레포>/` 에서 접속

파일명 대소문자가 import 경로와 정확히 일치해야 합니다(GitHub Pages는 대소문자를 구분).

## Netlify 배포

1. 레포 연결 후 Build command 비움(정적 파일)
2. Publish directory를 루트(`.`)로 설정
3. 환경변수 필요 시 Netlify UI에 등록
