# ADE-21 JSON 갱신 수정 설계 및 선행 조건

**상태:** 사전 분석 완료, 구현 미착수. 2026-09-18 확인 기준 PR #1이 미병합이므로 사용자 지시에 따라 정식 브랜치 생성·코드 변경·PR 제출을 보류한다.

**Goal:** SQLite 의존성을 제거하고 Excel → 공통 records JSON → 통계·예측 → FastAPI → 웹 흐름에서 실패한 갱신이 기존 데이터에 영향을 주지 않게 한다.

**Architecture:** 새 Excel을 검증하고 날짜·게이트 단위로 기존 records와 병합한다. 운영 파일과 같은 파일시스템의 임시 JSON을 대상으로 rebuild와 API 소비 경로를 검증한 뒤 운영 records.json만 원자적으로 교체한다.

**Tech Stack:** Python, openpyxl, JSON, FastAPI, pytest. 이 문서는 구현 코드가 아닌 병합 후 작업 설계다.

## 확인한 근거

- [ADE-21 본문과 최신 댓글](https://linear.app/ade0033/issue/ADE-21): PR #1 병합 후 최신 develop에서 `feature/ADE-6-12-json-refresh` 생성. 진행 기록은 ADE-21에만 작성.
- [공통 데이터/API 규격 v1.1](https://linear.app/ade0033/document/공통-데이터api-규격-v11-d7ea1e1bb9f4): 2026-09-18T02:18:57.043Z 수정본 확인. 11절의 v1 표기보다 12절의 v1.1 결정과 수정된 필드 표를 따른다.
- [프로젝트 운영 규칙](https://linear.app/ade0033/document/프로젝트-운영-규칙-git-discord-linear-47ad5b916f61), 공용 저장소 `docs/git-workflow.md`, README, PR 템플릿 확인.
- 두 저장소 및 상위 C:/, C:/study에서 AGENTS.md를 찾지 못했다. 저장소 내부 숨김·무시 파일도 이름 검색했다.
- [PR #1](https://github.com/kangseok0427/library-congestion-service/pull/1): open, merged=false, head=`8a6b1a4727593d596a43e365133eb2f84007aa93`.
- 원격 develop을 GitHub 연결 도구로 조회: `ee9ea709f79aeac72b1a7e93c2cd0f5a2ec08592`. 로컬 origin/develop과 같으며 초기 7개 파일뿐이다. API·rebuild·테스트가 아직 없다.
- [PR #2](https://github.com/kangseok0427/library-congestion-service/pull/2): closed, merged=false. 상태와 종료 사유만 확인했으며 재사용·재개하지 않았다.
- 기존 ETL: main, HEAD=`d8d26c2efac6cd7f1873d62686c940c12af1acbc`. 공용 저장소: main, HEAD=`ee9ea70`.
- 시작 시 두 저장소 모두 추적/미추적 변경 파일이 없었다. checkout, reset, stash, merge, commit, push를 하지 않았다.
- Git 셸의 GitHub 접속은 네트워크 제한으로 실패했다. 원격 상태는 GitHub 연결 도구로 확인했다. 전역 Git 설정은 변경하지 않았다.

## 기존 코드 분석

### 재사용할 부분

`library_etl/pipeline.py`의 `_date`, `_count`, `preprocess`에서 확장자·시트·필수 컬럼 검사, 정확한 합계/평균 행 제외, 날짜/게이트 정규화, 비정상 수치·수식·중복 거부, 부분 수집 표시, 08~23시 변환을 재사용한다. source_file에는 경로가 아닌 파일명만 저장한다.

완료 데이터가 부분 데이터로 덮이는 것을 막는 규칙과 날짜·게이트별 기존 이력 보존 규칙도 유지한다. 통로ID가 다르더라도 date+gate+hour가 같으면 합산하지 않는다.

### 제거하거나 변경할 부분

| 위치 | 확인한 현재 동작 | 수정 설계 |
| --- | --- | --- |
| ETL pipeline.py | sqlite3, records/state 테이블, DB 트랜잭션, read_current(database) | JSON 읽기·병합·검증·파일 교체로 전환 |
| ETL __main__.py | --database, DB export, DB 반영 후 JSON 내보내기 | --output이 운영 JSON 경로인 refresh 명령으로 변경, 성공 전 운영 파일 쓰기 금지 |
| ETL preprocess | OUT_11 정수 유지, TOTAL_HOURLY_MISMATCH 경고 | OUT_11은 null, 경고명을 HOURLY_TOTAL_MISMATCH로 통일 |
| PR #1 backend/domain.py | 정확히 12필드만 허용, OUT 정수 강제, OUT 단순 sum | 필수 12개 유지 + 선택 품질 필드 검증·보존, null 전파 |
| PR #1 backend/service.py | hourly_total_out 단순 sum | 미확정 OUT 포함 시 완전한 합계인 것처럼 0/부분합 반환 금지 |
| PR #1 backend/adapters.py | 별도 Excel 파서, TOTAL_MISMATCH 경고 | T02와 동작 통일, 우회 경로에서도 v1.1 준수 |
| PR #1 scripts/rebuild.py | rebuild(records, destination), 검증·today 계산 후 파일 교체 | T08에서 임시 destination으로 호출; 운영 경로를 직접 전달하지 않음 |
| PR #1 backend/app.py | FileProvider, LIBRARY_RECORDS, mtime_ns+size 캐시 | 파일 교체 후 재시작 없이 최신 snapshot을 읽는지 검증; 같은 크기 갱신도 검증 |

LIBRARY_DB는 현재 독립 ETL 및 PR #1 분석 범위에서는 발견되지 않았다. 종료된 PR의 DB 연동 코드를 이식하지 않으며, 병합 후 전체 소스/설정/문서에서 잔존 여부를 다시 검색한다.

FileProvider는 새 서비스 생성 성공 후 캐시를 바꾸지만 잘못된 파일을 외부에서 덮어쓰면 이전 캐시를 반환하지 않고 오류를 전파한다. 따라서 이번 작업은 실패한 T08 입력이 운영 경로에 도달하지 않도록 해야 한다. 외부 임의 파일 훼손 복구와 T08 rollback은 구분한다.

## v1.1 처리 결정과 LEAD 연결 조건

- 필수 필드는 date, day_of_week, gate, gate_name, passage_id, hour, in_count, out_count, total_in, total_out, is_partial, source_file이다.
- is_closed_day(boolean/null), is_low_volume(boolean), quality_note(string/null)는 선택이며 없어도 v1 입력을 읽는다. 유지되는 기존 records의 선택 필드를 삭제하지 않는다.
- 11시 OUT은 원본 숫자를 서비스 값으로 사용하지 않고 null로 변환한다. 일반 시간의 잘못된 OUT을 자동으로 null/0으로 바꾸지 않는다. OUT_11 결함 예외를 다른 필드 검증 완화에 사용하지 않는다.
- IN 합계 불일치는 원본 일일 총량·시간대 값을 그대로 두고 HOURLY_TOTAL_MISMATCH를 기록한다. OUT의 원본 합계 진단을 남기더라도 복제된 OUT_11을 포함한 원시 진단임을 표시한다. null 제외 부분합을 완전한 OUT 합계로 비교하지 않는다.
- records 내 quality_note와 반환 진단에 경고를 보존하는 방향으로 설계한다. 보고서 파일이 필요하면 커밋 전 준비하며, 별도 보고서 쓰기 실패를 운영 JSON 교체 후 갱신 실패로 보고하지 않도록 한다.
- v1.1은 정문·후문 행 복제를 GATE_ROW_DUPLICATED 오류로 지정한다. 같은 키 중복과 교차 게이트 복제는 별개다. 원본 측정 벡터 비교 기준은 T01/LEAD의 병합된 구현을 우선한다. 단순히 IN 합계가 같다는 이유로 복제 오류를 만들지 않는다.
- 공식 월요일 휴관, 휴관/부분 데이터 학습 제외, baseline 표본 2개 미만 null, 운영시간 제한은 v1.1의 LEAD 담당 영역이다. 현재 PR #1에 반영되지 않은 부분이 있으므로 병합 후 LEAD 변경과 충돌 없이 연결한다.
- 공통 집계 표에는 out_count 정수 표기가 남아 있다. null을 0으로 바꿀 수 없으므로 OUT 집계의 null 전파 및 hourly_total_out 의미를 LEAD 연동 시 명시해야 한다.

## 병합·게시 설계

공통 규격은 병합 또는 기간 교체를 허용하지만 정확한 범위를 고정하지 않았다. 기존 T08은 **입력에 있는 (date, gate)의 16시간 그룹만 교체**한다. PR #1 architecture 문서도 기간 병합은 T08 책임이고 rebuild CLI는 전체 snapshot 교체라고 설명한다. 이 기존 규칙을 기본값으로 유지한다.

1. 같은 운영 경로에 대한 동시 refresh를 직렬화한다. 잠금은 기존 JSON 읽기부터 마지막 교체까지 유지해 두 작업이 서로의 이력을 잃지 않게 한다. 다중 프로세스 CLI도 고려한다.
2. 새 Excel 전체를 검사하고 records로 변환한다. 빈 입력, 중복, 잘못된 필드가 하나라도 있으면 중단한다.
3. 운영 JSON이 있으면 검증해 읽는다. 파일이 있으나 손상된 경우 새 파일만으로 덮어쓰지 않고 실패한다. 최초 생성 시에는 빈 이력으로 시작한다.
4. 새 (date, gate) 집합에 속한 이전 records만 제거하고 새 16시간 그룹을 넣는다. 입력에서 빠진 다른 날짜·게이트는 그대로 둔다. min~max 날짜 범위를 통째로 삭제하지 않는다.
5. 완료→partial 하향 갱신을 거부하고 병합 결과의 키 유일성·일일 총량 일관성·선택 필드를 검증한다. 정렬을 고정해 반복 실행 시 논리적으로 동일한 결과가 되게 한다.
6. 운영 JSON과 같은 디렉터리 아래 임시 작업 영역을 만들고 임시 records.json을 기록한다. `rebuild(merged_records, staged_destination)`만 호출한다. rebuild 예외가 나면 운영 파일에 접근하지 않는다.
7. staged JSON을 다시 읽고 소비자 LibraryService의 통계·예측 직렬화까지 확인한다. 현재 rebuild는 today만 호출하므로 stats의 null 합계 오류를 별도로 검출한다. 테스트에서는 고정 날짜/시간을 사용한다.
8. 모든 필수 계산과 기록 준비가 끝난 후 staged 파일을 운영 records.json으로 단 한 번 os.replace한다. 임시 파일은 같은 파일시스템을 사용하고 flush/fsync한다. 기존 파일을 먼저 지우지 않는다.
9. 교체 실패 시 기존 파일 바이트·수정시각·API snapshot을 유지한다. 교체 후 정리 실패는 이미 반영된 갱신을 실패/rollback으로 오인하게 하지 않는다. 다중 JSON 파일을 차례대로 바꾸는 방식은 사용하지 않는다.
10. FileProvider가 다음 요청에서 새 snapshot을 읽는지 확인한다. stat/읽기 경합과 같은 크기·빠른 연속 교체의 캐시 충돌은 회귀 테스트로 확인하고 필요한 경우 파일 식별자/내용 버전 감지를 보강한다.

대안으로 rebuild에 운영 경로를 바로 넘길 수 있지만, 이후 통계·보고서 단계가 실패하면 이미 게시된 데이터를 되돌리기 어렵다. records·통계·예측을 별도 파일로 각각 교체하는 방식도 부분 성공 위험이 있어 채택하지 않는다. 현재 API는 records로 서비스를 재구성하므로 단일 records 게시가 현재 구조에 맞는다.

## 구현 순서 — PR #1 병합 후 실행

- [ ] 팀장이 PR #1을 병합했는지 다시 조회하고 최신 develop SHA와 v1.1 지원 상태를 확인한다.
- [ ] 작업 디렉터리 변경사항을 다시 확인한 뒤 최신 develop에서 `feature/ADE-6-12-json-refresh`를 새로 생성한다. 기존 PR #2 브랜치/커밋을 재사용하지 않는다.
- [ ] `tests/test_etl_preprocess.py`에 합성 XLSX로 null·경고·선택 필드·중복·부분 데이터 검증을 추가하고 실패를 먼저 확인한다.
- [ ] `library_etl/pipeline.py`, `library_etl/__init__.py`, `backend/adapters.py`, 필요한 `backend/domain.py`/`backend/service.py`를 v1.1 경계에 맞춘다. LEAD가 이미 수정했다면 중복 구현하지 않는다.
- [ ] `tests/test_json_refresh.py`에 아래 실패 주입·반복·이력 보존 테스트를 추가한 뒤 `library_etl/refresh.py`에 JSON 갱신을 구현한다. DB read_current와 테이블 코드는 제거한다.
- [ ] `tests/test_refresh_api.py`에서 FileProvider와 FastAPI TestClient를 갱신 전후 동일 인스턴스로 유지하여 응답 변경/유지를 확인한다.
- [ ] `library_etl/__main__.py`, README, architecture 문서, requirements 및 CI를 실제 구현과 맞추고 SQLite/LIBRARY_DB/DB export 명령을 제거한다.
- [ ] 전체 자동 테스트, 최신 develop 충돌 검사, 커밋 목록과 민감 파일 포함 여부 검토 후 develop 대상 새 PR을 생성한다.
- [ ] PR 제목에 [ADE-21], 본문에 ADE-6/ADE-12/ADE-21 링크와 정확한 테스트 결과를 넣는다. CI 완료를 확인하고 실패를 수정한다. 진행·검증·PR 링크는 ADE-21에만 기록한다. 팀장이 병합하며 검토가 남으면 Done으로 바꾸지 않는다.

## 필수 테스트와 판정 기준

| 검증 | 합성 입력/실패 주입 | 통과 조건 | 이번 실행 |
| --- | --- | --- | --- |
| 정상 Excel 갱신 | 과거 고정 날짜의 IN·일일 총량 변경 | JSON 내용과 동일 client의 stats/forecast 값 변경 | 미실행: 구현 전 |
| 잘못된 Excel | 누락 컬럼·수식·음수·중복 | 오류 + 기존 JSON 바이트/mtime/API 응답 동일 | 미실행 |
| rebuild 실패 | 임시 destination 기록 전/후 예외 | 운영 JSON과 API 값 동일 | 미실행 |
| 파일 교체 실패 | 최종 운영 경로 os.replace에 PermissionError 주입 | 기존 JSON 바이트/mtime/API 유지, 기존 파일 선삭제 없음 | 미실행 |
| 동일 Excel 반복 | 같은 파일 두 번 + 다른 과거 그룹 | 키/행수/내용 중복 없음, 다른 이력 보존 | 미실행 |
| 품질 규칙 | OUT_11 복제값·일일 총량 불일치 | null 유지, 0 대체 없음, 경고 유지, 원본 총량 불변 | 미실행 |
| 회귀 | 병합된 공용 전체 테스트 + 기존 ETL 보존 규칙 | 의도한 규격 변경 외 회귀 없음 | 기존 ETL만 13 통과/1 skip |

추가로 같은 크기 연속 갱신, 완료→partial 거부, 선택 필드 없는 v1 입력, 선택 필드가 있는 v1.1 입력, 동시 갱신 이력 유실 방지, 손상된 기존 JSON 거부를 검증한다. prediction 응답의 reference_time처럼 시계에 따라 바뀌는 값은 시간을 고정한 뒤 비교한다.

## 실행 기록과 향후 실행 예시

실제로 실행한 기존 ETL 기준선 명령(PowerShell, ETL 저장소에서):

```powershell
& 'C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -B -m unittest discover -s tests -v
```

결과: 14개 발견, 13개 통과, 1개 skip, 종료코드 0. skip은 LIBRARY_SAMPLE_XLSX 미지정에 따른 실제 파일 테스트다. 두 작업 저장소에서 숨김/무시 파일을 포함해 xlsx/xls를 검색했으나 발견하지 못했다. 다른 개인 폴더 전체를 검색한 것은 아니다. 처음 PATH의 python은 실행 권한 오류로 시작하지 못했고 번들 Python으로 다시 실행했다.

공용 develop에는 테스트가 없으므로 공용 회귀/새 JSON 통합/CI는 실행하지 않았다. PR #1 본문의 26 passed는 작성자의 과거 기록이며 이번 실행 결과로 주장하지 않는다. 새 브랜치가 없어 최신 develop 대상 충돌 검사는 아직 해당 없음이며 PR 직전에 재검사한다.

다음은 **구현 후 제공할 CLI 형태의 설계 예시이며 현재 실행 가능한 명령이 아니다**:

```powershell
python -m library_etl refresh 'C:\data\new.xlsx' --output 'data/processed/records.json' --partial-date 2026-09-18
$env:LIBRARY_RECORDS = (Resolve-Path 'data/processed/records.json').Path
python -m uvicorn backend.app:app
python -m pytest -q
```

원본 Excel·운영 JSON·실데이터 보고서·개인정보는 커밋하지 않는다. CI의 Excel은 openpyxl로 합성해 임시 디렉터리에 생성한다. 실제 파일 검증이 가능해지면 별도 로컬 출력으로 수행하고 결과 요약만 남긴다.

## 현재 필요한 선행 조건

1. 팀장의 PR #1 → develop 병합. 현재 가장 직접적인 작업 차단 조건이다.
2. 병합본의 v1.1 소비자 지원(null OUT, 선택 필드, null 집계) 확인 및 LEAD 담당 변경과 조율. 현재 PR #1 head만으로는 v1.1 T02 출력을 받을 수 없다.
3. 정식 구현 후 실행 가능한 Python 의존성 환경과 push/PR/CI 접근 확인. 현재 읽기 및 Linear 댓글 도구는 사용 가능하지만 Git 셸 네트워크는 제한되어 있다.

이 문서는 로컬 미커밋 산출물로 남긴다. 구현 완료나 ADE-21 Done을 의미하지 않는다.
