# T02 데이터 전처리 · T08 Excel 갱신

박수용 담당 작업. Python 3.11 이상. Linear 공통 데이터/API 규격 v1(2026-09-14)을 기준으로 구현한 독립 모듈입니다.
T02=ADE-6, T08=ADE-12. 팀 저장소와 아직 연결하지 않았습니다.

## 현재 구현 범위

- T02: 실제 XLSX 읽기, 요약행 제외, 08~23시 세로 변환, 공통 12개 필드 출력, 오류/중복 검증.
- T08: XLSX 경로로 수동 갱신, SQLite 저장, 날짜·게이트 단위 교체, 기존 기간 유지, 반복 실행 중복 방지.
- 재계산 함수를 연결하는 `rebuild` 인자와 결과를 함께 읽는 `read_current` 제공.
- 통계(T03), 혼잡 판정(T04), 예측(LEAD), HTTP API(T06), 업로드 화면은 각 담당자 코드와 연결해야 합니다. 현재 해당 기능의 완성품이 아닙니다.
- API 연결 검증 전이므로 T08 전체 완료 또는 Linear Done으로 판단하면 안 됩니다.

## Windows에서 실행

ZIP 압축을 푼 뒤 `requirements.txt`가 보이는 `library_etl` 폴더를 VS Code로 엽니다.
VS Code 메뉴의 터미널 → 새 터미널에서 다음 명령을 실행합니다. PowerShell 가상환경 활성화 없이 실행할 수 있습니다.

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

아래 `C:\data\방문자수-20260910_114341.xlsx`는 실제 파일 위치로 바꿉니다.

```powershell
.\.venv\Scripts\python.exe -m library_etl preprocess "C:\data\방문자수-20260910_114341.xlsx" --partial-date 2026-09-10 --output output_first
.\.venv\Scripts\python.exe -m library_etl refresh "C:\data\방문자수-20260910_114341.xlsx" --partial-date 2026-09-10 --database data/library.sqlite3 --output output_refresh
.\.venv\Scripts\python.exe -m library_etl export data/library.sqlite3 --output output_export
```

출력 폴더에는 `records.json`, `report.json`이 생성됩니다. 기존 폴더 덮어쓰기를 막으므로 재실행 시 새 출력 폴더명을 사용합니다.
검증 실패 시 종료코드 2, `errors.json`에 행/필드별 오류가 기록됩니다. 파일 I/O 실패는 종료코드 3입니다.
새 파일을 갱신할 때 원본 파일명은 무엇이든 사용할 수 있으며, 파일 내용의 컬럼 구조를 검사합니다.
여러 시트가 있으면 `--sheet Data`처럼 시트를 명시합니다. 현재 `.xlsx`만 지원합니다.

## 출력 규격

`records.json`은 아래 필드만 가진 객체의 배열입니다. 숫자는 JSON 정수, `is_partial`은 JSON boolean입니다.

| 필드 | 타입 | 의미 |
| --- | --- | --- |
| date | string | YYYY-MM-DD |
| day_of_week | string | Mon/Tue/Wed/Thu/Fri/Sat/Sun |
| gate | string | front/back |
| gate_name | string | 자료실.정문/자료실.후문 |
| passage_id | string | 원본 통로ID |
| hour | integer | 8~23 |
| in_count | integer | 해당 시간대 IN |
| out_count | integer | 해당 시간대 OUT |
| total_in | integer | 원본 날짜·게이트 전체_IN |
| total_out | integer | 원본 날짜·게이트 전체_OUT |
| is_partial | boolean | 부분 수집 날짜 여부 |
| source_file | string | 원본 파일명 |

`report.json`과 SQLite 내부 상태는 진단/저장용입니다. 공통 API 응답으로 그대로 반환하는 규격이 아닙니다.
월은 T03에서 date로부터 계산합니다. 표준 12개 필드에 임의로 month 같은 필드를 추가하지 않았습니다.

## 실제 파일에서 확인한 사항

검증 파일: 방문자수-20260910_114341.xlsx, 시트 Data.

- 헤더 1행, 합계·평균 2행, 정상 데이터 1,836행.
- 2022-10-07~2026-09-10 중 기록이 존재하는 918일. 연속된 매일 데이터라는 뜻이 아닙니다.
- 정문 918행, 후문 918행. 시간대 변환 후 29,376행.
- 원본 필수값 누락 0, 날짜·게이트 중복 0.
- 전체_IN과 08~23시 합계 불일치 1,815행, OUT 불일치 1,811행.
- 총 3,626개의 불일치 경고를 기록하고 원본값은 변경하지 않습니다. 센서 집계 구간 차이 등 원인은 T01/도서관 확인이 필요합니다.

**total_in/total_out은 시간대마다 반복되는 날짜·게이트 전체값입니다. 16개 시간대를 합산하면 16배 중복 집계됩니다.**
T03은 날짜·게이트별 전체값을 한 번만 사용하거나, 시간대 수치 합계를 별도 지표로 계산해야 합니다.
시간대 합계가 일별 전체와 같다고 가정하면 안 됩니다. 누락된 날짜를 휴관일/0명으로 추정하지 않습니다.

## 부분 데이터와 오류 처리

- 이 샘플은 2026-09-10 오전 추출 자료이므로 실행 예제에서 `--partial-date 2026-09-10`을 명시합니다.
- 수집 중인 날짜의 뒤 시간대 0은 원본대로 보존하되 해당 날짜 전체를 partial로 표시합니다. 미래 시간대의 확정 0명으로 해석하면 안 됩니다.
- 파일명이나 실행 날짜만으로 수집 완료 여부를 확정하지 않습니다.
- 부분 날짜를 명시하지 않으면 최신 날짜를 보수적으로 partial 처리하고 `PARTIAL_DATE_UNCONFIRMED` 경고를 남깁니다. 이는 확인 전 기본 처리이며 실제 완료 여부의 판정이 아닙니다.
- 모든 날짜가 완료되었다고 확인한 경우에만 `--all-complete`를 사용합니다.
- 같은 날짜·게이트의 중복은 다른 통로ID여도 중복 오류로 처리합니다. 규격의 키가 date+gate+hour이기 때문입니다.
- 빈 값, 음수, 소수, 비수치, 수식, 알 수 없는 게이트, 잘못된 날짜가 있으면 해당 파일 전체를 반영하지 않습니다. 오류 목록을 보존합니다.
- 원본 합계/평균 행은 정확한 `전체/그룹/합계` 또는 `전체/그룹/평균` 조합만 제외합니다.
- 완료된 날짜를 부분 데이터로 덮어쓰는 갱신은 차단합니다.
- 수집시각이 공통 필드에 없으므로 같은 완결성의 신구 파일을 자동 판별하지 못합니다. 운영자는 최신 파일인지 확인해야 합니다.

## 팀 코드 연결

T02 함수는 `(list[dict], report)`를 반환합니다. `list[dict]`가 공통 전처리 출력입니다.

```python
from library_etl import preprocess, refresh, read_current

records, report = preprocess(
    "new.xlsx",
    partial_dates=["2026-09-10"],
    source_name="방문자수-20260910_114341.xlsx",
)

# T03/LEAD 함수가 준비되면 팀 통합 코드에서 연결합니다.
# 함수 이름/모듈 경로는 아직 합의되지 않았으므로 아래는 연결 구조 예시입니다.
def rebuild(records):
    stats = team_statistics(records)  # 팀 T03 함수
    predictions = team_prediction(stats)  # 팀 LEAD 함수
    return {"stats": stats, "predictions": predictions}

state = refresh("new.xlsx", "data/library.sqlite3",
                partial_dates=["2026-09-10"], rebuild=rebuild)
records, state = read_current("data/library.sqlite3")
```

위 team_statistics/team_prediction은 실제 코드에 포함된 함수가 아닙니다. 팀 함수로 교체해야 합니다.
`rebuild`는 JSON으로 저장 가능한 값을 반환하고 외부 DB/API/파일을 직접 수정하지 않는 순수 계산 함수로 작성합니다.
SQLite 트랜잭션은 외부 부수효과까지 되돌릴 수 없습니다. 장시간 계산은 쓰기 잠금을 길게 유지하므로 MVP에서는 가벼운 계산을 연결합니다.

- 재계산 성공 시 원본 전처리 데이터와 파생 결과를 하나의 트랜잭션으로 반영합니다.
- 재계산 실패 시 데이터/파생 결과 모두 이전 상태를 유지합니다.
- 재계산 미연결 시 `integration_status=pending_downstream`, `derived=null`입니다. 이전 예측을 최신 결과처럼 반환하지 않습니다.
- API 담당자는 `read_current`로 같은 시점의 데이터와 파생 결과를 읽고 기존 공통 JSON 구조로 변환합니다. 캐시가 있다면 갱신 처리도 연결해야 합니다.
- `DataError.as_dict()`는 공통 `{error: {code, message}}` 구조입니다. 행별 오류는 별도 `issues` 속성에 있습니다.
- 통합 검증: 새 파일 갱신 → 통계 재계산 → 예측 재계산 → 실제 API 값 변경 확인. 마지막 API 확인은 팀 서버가 준비된 후 수행해야 합니다.

## Git 인수인계

팀 저장소의 기존 디렉터리 구조를 확인한 뒤 `library_etl/` 패키지와 테스트를 넣습니다.
팀 requirements가 있으면 openpyxl 요구사항을 병합하고 기존 파일을 통째로 덮어쓰지 않습니다.
프로젝트별 공통 규격 원문은 `docs/common_contract_v1.md`에 보관했습니다.

권장 순서: develop 기준 기능 브랜치 생성 → 코드/테스트 추가 → feature에서 develop으로 PR → 팀장 검토.
원본 XLSX, SQLite DB, 실행 결과, 가상환경은 Git에 커밋하지 않습니다.
예시 브랜치: `feature/ADE-6-ADE-12-excel-pipeline`.

## 검증 기록

Python 환경에서 14개 unittest 통과. 실제 파일 테스트 포함.
같은 파일 두 번 갱신 후 29,376행 유지 및 내용 동일 확인.
결측/음수/소수/수식/중복/잘못된 날짜·게이트 차단, 후속 계산 실패 시 롤백 확인.
Windows 명령은 제공했으나 실제 Windows 환경 실행과 팀 API 통합은 아직 검증하지 않았습니다.
원본 파일은 ZIP에 포함하지 않습니다. 실제 파일 테스트를 실행하려면 다음처럼 지정합니다.

```powershell
$env:LIBRARY_SAMPLE_XLSX = "C:\data\방문자수-20260910_114341.xlsx"
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```
