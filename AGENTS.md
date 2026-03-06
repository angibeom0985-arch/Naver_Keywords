# Codex 작업 규칙 (auto_keyword)

## Git 자동 커밋
- 완료된 모든 코딩 작업의 마지막에 `git add -A` 후 커밋을 만들고 `origin`에 푸시한다.
- 사용자가 푸시하지 말라고 명시하지 않는 한 커밋/푸시를 생략하지 않는다.
- 구현한 수정 내용을 요약하는 명확한 커밋 메시지를 사용한다.

## 사용자 명령 우선/충돌 처리
- 사용자 명령과 작업 규칙이 충돌하면 즉시 거부하지 말고 사용자에게 확인 요청을 먼저 한다.
- 사용자가 진행을 명시하면 사용자 명령을 우선 적용한다.
- 파괴적 변경(대량 삭제/배포 파일 제거 등)은 확인 후 수행한다.

## 머신 ID 보호 (현재 프로젝트 기준)
- 머신 ID는 `1인1PC` 및 사용 기간 검증의 핵심이다.
- 머신 ID 형식은 `Gold-Keyword-` + 대문자 32자리 hex만 사용한다.
- `MID-*`, `Gold Keyword-*`, `Gold-Keywords-*`, `Gold Keyword-MID-*` 등 레거시/오타 형식은 호환 조회를 허용하지 않는다.
- 머신 ID 관련 코드(`get_machine_id`, 라이선스 매칭, 저장/조회 경로, 포맷 검증)를 수정할 때는 작업 전에 반드시 사용자에게 확인 요청을 먼저 한다.
- 머신 ID 저장/조회 경로는 현재 구현 기준을 유지한다.
  - 레지스트리: `HKCU\\Software\\AutoNaverKeyword\\MachineId` (단일 저장소)
- 머신 ID 관련 로직 변경 시 라이선스 검증(`check_license_from_sheet`) 호환성을 함께 유지한다.
  - 기존 유효 구매자 통과
  - 만료 구매자 실패
  - 미등록 머신 실패

## 라이선스 시트/만료일 정책 (재발 방지)
- 라이선스 시트 URL은 현재 운영 문서 기준을 사용한다.
  - `https://docs.google.com/spreadsheets/d/1Ooj_BnlIiYC7a1de-csqmXxnGjTwXU5d2q5smi_Os30/export?format=csv&gid=0`
- `만료일`은 반드시 `YYYY-MM-DD` 형식의 실제 날짜여야 하며, 공백/문자열/`nan`은 미등록으로 처리한다.
- 만료 판정은 유예 없이 적용한다(오늘 날짜가 만료일을 지나면 즉시 만료 차단).
- 만료 시에는 `ExpiredDialog`(카카오톡 문의 버튼 포함)를 표시한다.
- 미등록 테스트용 행(`미등록`)을 운영할 때는 만료일을 반드시 과거 날짜로 유지한다.

## Apps Script/웹훅 보안 정책
- EXE 코드 및 `setting/api_keys.json`에 Apps Script 웹앱 URL/토큰(`usage_webhook_*`)을 포함하지 않는다.
- 머신 ID 자동 등록/자동 append를 유발하는 Apps Script 로직은 운영 배포본에서 사용하지 않는다.
- 사용량 추적이 필요하면 라이선스 시트와 분리된 별도 수집 경로로 운영하고, 토큰은 주기적으로 교체한다.

## 머신 ID 파일 정책
- `machine_id.txt` 잔여 파일이 dist/패키지 산출물에 포함되지 않도록 배포 전 점검한다.
- 불필요한 머신 ID 임시 파일/중복 파일은 배포 전 제거한다.
- 머신 ID는 레지스트리만 사용하며, `machine_id.txt` 및 `.auto_naver_machine_id.txt`는 생성하지 않는다.

## 프로젝트 구조/운영
- Python 단일 소스는 루트 `Auto_Naver_Gold_Keyword.py`를 기준으로 관리한다.
- `setting` 폴더는 실행/배포 보조 파일(`.exe`, `api_keys.json`, `auto_naver.ico`)만 사용한다.
- `assets` 폴더는 생성/사용하지 않는다(요청 시에도 먼저 확인).
