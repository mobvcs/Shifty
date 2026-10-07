# Supabase 설정 가이드

이 문서는 `시프티` 프로젝트를 Supabase에 연결하고, 현재 코드(`js/supabaseClient.js`)가 사용하는 정규화 테이블을 준비하는 절차를 안내합니다.

## 목적

- 프론트엔드(정적 사이트)에서 Supabase를 통해 데이터 조회/저장
- 직원/시프트/이벤트/OCC/셀색상/연차설정/월마감/포지션색상/OCC메타/사용자 지정 휴무일 데이터를 정규화 테이블로 관리

## 사용 테이블

- `shift_codes`: 근무 코드
- `employees`: 근무자 기본 정보
- `shift_entries`: 일자별 시프트
- `shift_events`: 이벤트(항목별 배경색 포함)
- `occ_daily`: 객실 수, OCC(%)
- `occ_meta`: 월별 OCC update 표시일
- `cell_colors`: 셀 배경색
- `leave_configs`: 직원별 연차 설정
- `month_closings`: 월 마감 상태 및 다음 월 AL 이월값, 월별 순서 정보(이벤트/팀 섹션)
- `position_colors`: 포지션별 배경색
- `custom_off_days`: 사용자 지정 휴무일

## 1) Supabase 프로젝트 생성

1. Supabase에서 새 프로젝트를 생성합니다.
2. 프로젝트 대시보드에서 아래 값을 확인합니다.
   - `Project URL`
   - `anon public key`

## 2) 테이블 스키마 생성

프로젝트 루트의 `supabase-schema.sql` 파일 내용을 Supabase SQL Editor에 붙여 실행합니다.

```sql
-- 프로젝트 파일: supabase-schema.sql
-- 그대로 실행
```

실행 후 위 11개 테이블이 생성되어야 합니다.

> 참고: 월별 순서 정보는 `month_closings.carry_map(jsonb)`에 저장하므로, 별도 테이블 추가 SQL은 필요하지 않습니다.

## 3) RLS(권한) 설정

현재 앱은 브라우저에서 `anon key`로 직접 접근하므로, 테스트 단계에서는 아래 두 가지 중 하나를 선택해야 합니다.

- 방법 A: RLS 비활성화(간단, 내부 테스트용)
- 방법 B: RLS 활성화 + 허용 정책 추가(권장)

예시(방법 B, 빠른 시작용 전면 허용 정책):

```sql
alter table shift_codes enable row level security;
alter table employees enable row level security;
alter table shift_entries enable row level security;
alter table shift_events enable row level security;
alter table occ_daily enable row level security;
alter table occ_meta enable row level security;
alter table cell_colors enable row level security;
alter table leave_configs enable row level security;
alter table month_closings enable row level security;
alter table position_colors enable row level security;
alter table custom_off_days enable row level security;

create policy "anon read shift_codes" on shift_codes for select to anon using (true);
create policy "anon write shift_codes" on shift_codes for all to anon using (true) with check (true);

create policy "anon read employees" on employees for select to anon using (true);
create policy "anon write employees" on employees for all to anon using (true) with check (true);

create policy "anon read shift_entries" on shift_entries for select to anon using (true);
create policy "anon write shift_entries" on shift_entries for all to anon using (true) with check (true);

create policy "anon read shift_events" on shift_events for select to anon using (true);
create policy "anon write shift_events" on shift_events for all to anon using (true) with check (true);

create policy "anon read occ_daily" on occ_daily for select to anon using (true);
create policy "anon write occ_daily" on occ_daily for all to anon using (true) with check (true);

create policy "anon read occ_meta" on occ_meta for select to anon using (true);
create policy "anon write occ_meta" on occ_meta for all to anon using (true) with check (true);

create policy "anon read cell_colors" on cell_colors for select to anon using (true);
create policy "anon write cell_colors" on cell_colors for all to anon using (true) with check (true);

create policy "anon read leave_configs" on leave_configs for select to anon using (true);
create policy "anon write leave_configs" on leave_configs for all to anon using (true) with check (true);

create policy "anon read month_closings" on month_closings for select to anon using (true);
create policy "anon write month_closings" on month_closings for all to anon using (true) with check (true);

create policy "anon read position_colors" on position_colors for select to anon using (true);
create policy "anon write position_colors" on position_colors for all to anon using (true) with check (true);

create policy "anon read custom_off_days" on custom_off_days for select to anon using (true);
create policy "anon write custom_off_days" on custom_off_days for all to anon using (true) with check (true);
```

> 보안 주의: 위 정책은 빠른 연동 확인용입니다. 운영 환경에서는 팀/역할/사용자 기준으로 정책을 좁혀야 합니다.

## 4) 프론트엔드 연결

`index.html`에서 모듈 스크립트(`./js/main.js`)보다 먼저 전역 변수를 설정해야 합니다.

```html
<script>
  window.SUPABASE_URL = "https://YOUR_PROJECT.supabase.co";
  window.SUPABASE_ANON_KEY = "YOUR_ANON_KEY";
</script>
```

`js/config.js`의 `getSupabaseConfig()`가 위 값을 읽어 클라이언트를 생성합니다.

## 5) 동작 확인 체크리스트

1. 페이지 진입 시 콘솔에 `근무 코드 동기화 완료` 로그 확인
2. 셀 편집 후 새로고침 시 값 유지 확인 (`shift_entries`)
3. 근무자 추가/수정/삭제 후 유지 확인 (`employees`)
4. 이벤트 추가/수정/삭제 + 항목 배경색 유지 확인 (`shift_events`)
5. OCC Import 후 유지 확인 (`occ_daily`)
6. 사용자 지정 셀 색상 유지 확인 (`cell_colors`)
7. 잔여 연차/가산 옵션 유지 확인 (`leave_configs`)
8. 월 마감/마감취소 후 상태 유지 확인 (`month_closings`)
9. 이벤트 칩 드래그 순서(월별) 저장/복원 확인 (`month_closings.carry_map.__event_order__`)
10. 팀 섹션 드래그 순서(월별) 저장/복원 확인 (`month_closings.carry_map.__team_section_order__`)
11. 포지션별 배경색 저장/복원 확인 (`position_colors`)
12. OCC update 표시일 저장/복원 확인 (`occ_meta`)
13. 관리자 지정 휴무일 저장/복원 확인 (`custom_off_days`)
14. 근무자 퇴사일 저장/복원 확인 (`employees.resignation_date`)

## 6) `month_closings.carry_map` 저장 키

`month_closings`는 `month_key` 기준으로 월 단위 상태를 저장합니다.  
이때 `carry_map(jsonb)` 내부에 아래 키를 함께 저장합니다.

- `__carry__`: 다음 월 AL 이월값 맵
- `__rollback__`: 마감 취소 시 롤백값 맵
- `__employees_snapshot__`: 마감 월 근무자 스냅샷
- `__schedules_snapshot__`: 마감 월 스케줄 스냅샷
- `__custom_colors_snapshot__`: 마감 월 사용자 색상 스냅샷
- `__events_snapshot__`: 마감 월 이벤트 스냅샷
- `__event_order__`: 이벤트 항목 드래그 순서(월별)
- `__team_section_order__`: 팀 섹션 드래그 순서(월별)

## 7) 자주 발생하는 문제

- **`Supabase SDK가 로드되지 않았습니다.`**
  - `index.html`의 Supabase CDN script가 누락되었는지 확인
- **데이터 저장이 안 됨**
  - RLS 정책 또는 테이블 생성 여부 확인
- **파일명 규칙 오류(OCC Import)**
  - 업로드 파일명에 `M000005326` 포함 필요

