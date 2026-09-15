원문 출처: https://linear.app/ade0033/document/공통-데이터api-규격-v1-2026-09-14-d7ea1e1bb9f4
최종 수정: 2026-09-14T01:20:11.411Z

# 목적

이 문서는 전처리, 통계, 백엔드/API, 프론트엔드, Excel 갱신 담당자가 서로 같은 형식으로 데이터를 주고받기 위한 공통 규격입니다.

내부 구현 방식이나 라이브러리는 각 담당자가 자유롭게 선택해도 되지만, 다른 작업으로 넘기는 입력/출력 형식은 이 문서에 맞춥니다.

# 1. 원본 Excel 기준

현재 확인한 원본 파일 기준 주요 컬럼은 다음과 같습니다.

* 수집일자
* 게이트
* 통로ID
* 시작IN / 종료IN
* 시작OUT / 종료OUT
* 전체\_IN / 전체\_OUT
* IN_08 / OUT_08 \~ IN_23 / OUT_23

현재 확인된 게이트 예시는 `자료실.정문`, `자료실.후문`입니다.

주의: 현재 데이터에서 단순히 IN-OUT 누적으로 '현재 체류인원'을 확정하지 않습니다. 센서 누적값/리셋 방식과 오차를 완전히 확인하기 전까지 사용자 화면에는 `현재 이용자 수`를 표시하지 않습니다.

# 2. 전처리 표준 출력 규격

원본의 가로형 시간대 컬럼을 아래 세로형 구조로 변환합니다.

| 필드명 | 타입 | 예시 | 설명 |
| -- | -- | -- | -- |
| date | string(YYYY-MM-DD) | 2026-09-10 | 수집일자 |
| day_of_week | string | Thu | 요일 |
| gate | string | front / back | 정문/후문 표준값 |
| gate_name | string | 자료실.정문 | 원본 게이트명 |
| passage_id | string | 10.10.2.37A | 통로ID |
| hour | integer | 11 | 시간대, 8\~23 |
| in_count | integer | 13 | 해당 시간대 IN |
| out_count | integer | 18 | 해당 시간대 OUT |
| total_in | integer | 75 | 해당 날짜·게이트 전체 IN |
| total_out | integer | 58 | 해당 날짜·게이트 전체 OUT |
| is_partial | boolean | true | 당일 수집 중인 부분 데이터 여부 |
| source_file | string | 방문자수-20260910_114341.xlsx | 원본 파일명 |

## 전처리 규칙

1. `수집일자`는 반드시 `YYYY-MM-DD` 형식으로 통일합니다.
2. `자료실.정문`은 `front`, `자료실.후문`은 `back`으로 표준화합니다.
3. 시간은 `08~23`을 숫자 `8~23`으로 저장합니다.
4. IN/OUT 값은 정수형으로 변환합니다.
5. 빈 셀, 숫자로 변환할 수 없는 값, 음수 값은 오류 목록에 기록합니다.
6. 같은 `date + gate + hour` 조합이 중복되면 자동 합산하지 말고 중복으로 표시합니다.
7. 최신 날짜가 당일 진행 중 데이터이면 `is_partial=true`로 표시합니다.
8. 결측값을 임의로 0으로 바꾸지 않습니다. 원본이 실제 0인지 누락인지 구분이 필요합니다.

# 3. 팀 내부 공통 집계 규격

정문/후문을 합쳐 도서관 전체 시간대 이용량을 만들 때 아래 필드를 사용합니다.

| 필드명 | 타입 | 설명 |
| -- | -- | -- |
| date | string | 기준 날짜 |
| hour | integer | 시간대 |
| in_count | integer | 정문+후문 시간대 IN 합계 |
| out_count | integer | 정문+후문 시간대 OUT 합계 |
| visit_count | integer | 기본 방문량 지표. 우선 `in_count` 사용 |
| baseline_avg | number/null | 비교 기간의 같은 요일·시간 평균 방문량 |
| difference_rate | number/null | 평소 대비 증감률(%) |
| congestion_score | number/null | 0\~100으로 정규화한 혼잡 점수 |
| congestion_level | string/null | quiet / normal / busy |
| is_partial | boolean | 부분 데이터 여부 |

## 방문량 기준

MVP의 기본 방문량은 `IN`을 사용합니다.

* `visit_count = front.in_count + back.in_count`
* OUT은 검증 및 참고용으로 보존합니다.
* IN/OUT 차이를 현재 체류인원으로 해석하지 않습니다.

# 4. 기본 통계 담당 출력

최소한 다음 결과를 만들 수 있어야 합니다.

* 날짜별 총 입장량
* 시간대별 입장량
* 요일별 평균 입장량
* 월별 평균 입장량
* 같은 요일·같은 시간대 평균
* 평소 대비 증감률

## 평소 비교 기본값

초기 MVP에서는 가능하면 `이전 4주 같은 요일의 같은 시간대 평균`을 baseline으로 사용합니다.

예: 2026-09-14 월요일 14시를 표시할 경우, 직전 4개 월요일의 14시 IN 합계를 평균냅니다.

데이터가 부족하면 사용 가능한 기간만 쓰고 `baseline_avg=null` 처리할 수 있습니다.

# 5. 혼잡도 규격

혼잡도 단계는 3단계로 통일합니다.

* `quiet` = 여유
* `normal` = 보통
* `busy` = 혼잡

초기 기준은 절대 인원 기준이 아니라 해당 시간대의 과거 분포 기준으로 정합니다.

권장 초기안:

* 하위 35% 이하 → quiet
* 35% 초과 \~ 70% 이하 → normal
* 70% 초과 → busy

이 값은 T04 혼잡도 기준 설계 담당자가 실제 분포를 본 뒤 조정할 수 있습니다. 조정하더라도 API의 `quiet / normal / busy` 값 이름은 바꾸지 않습니다.

# 6. API 공통 규격

기본 API 경로는 아래처럼 통일합니다.

## GET /api/v1/congestion/today

오늘의 예상 혼잡도와 시간대별 정보를 반환합니다.

```json
{
  "date": "2026-09-14",
  "data_status": "forecast",
  "reference_time": "2026-09-14T10:00:00+09:00",
  "congestion": {
    "level": "normal",
    "label": "보통",
    "score": 58
  },
  "recommendation": {
    "best_start_hour": 10,
    "best_end_hour": 12,
    "message": "오전 10시~12시 방문을 추천합니다."
  },
  "hourly": [
    {
      "hour": 8,
      "expected_visitors": 24,
      "baseline_avg": 31.5,
      "difference_rate": -23.8,
      "level": "quiet"
    },
    {
      "hour": 9,
      "expected_visitors": 38,
      "baseline_avg": 35.2,
      "difference_rate": 8.0,
      "level": "normal"
    }
  ],
  "updated_at": "2026-09-14T10:00:00+09:00"
}
```

## GET /api/v1/stats?date=YYYY-MM-DD

선택한 날짜의 실제/집계 데이터를 반환합니다.

```json
{
  "date": "2026-09-10",
  "data_status": "partial",
  "total_in": 122,
  "total_out": 114,
  "hourly": [
    {
      "hour": 8,
      "in_count": 10,
      "out_count": 6
    }
  ]
}
```

## GET /api/v1/meta

프론트에서 사용할 서비스 기준정보를 반환합니다.

```json
{
  "library_name": "용산꿈나무도서관",
  "available_hours": [8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23],
  "levels": {
    "quiet": "여유",
    "normal": "보통",
    "busy": "혼잡"
  }
}
```

# 7. API 공통 오류 응답

모든 API 오류는 가능하면 같은 형태를 사용합니다.

```json
{
  "error": {
    "code": "DATA_NOT_FOUND",
    "message": "해당 날짜의 데이터를 찾을 수 없습니다."
  }
}
```

초기 공통 오류 코드:

* `DATA_NOT_FOUND`
* `INVALID_DATE`
* `EXCEL_FORMAT_ERROR`
* `PROCESSING_ERROR`

# 8. 프론트엔드 필수 표시값

MVP 웹페이지에는 최소 아래 항목을 표시합니다.

1. 기준 날짜
2. 오늘의 예상 혼잡 단계(여유/보통/혼잡)
3. 시간대별 예상 방문량 그래프
4. 추천 방문 시간
5. 평소 같은 요일·시간대 대비 정보(가능한 경우)
6. 데이터 기준/최근 갱신 시각

## 표시 금지/주의

* 데이터 검증 전 `현재 이용자 82명` 같은 실시간 체류인원 표현 금지
* 예측값은 반드시 `예상`, `과거 이용 패턴 기준` 등의 표현 사용
* 부분 데이터는 `현재까지 집계된 데이터`라고 표시

# 9. Excel 갱신 규격

새 Excel 파일이 들어오면 다음 순서로 처리합니다.

1. 파일 확장자/시트/필수 컬럼 검사
2. 전처리 규격으로 변환
3. 중복 날짜·게이트 데이터 확인
4. 기존 데이터에 병합 또는 해당 기간 교체
5. 통계 재계산
6. 예측 재계산
7. API가 최신 결과를 반환하는지 확인
8. 처리 결과 및 오류 기록

파일명은 고정하지 않습니다. 파일 내용의 컬럼 구조를 기준으로 검사합니다.

# 10. 담당자 간 인수인계 기준

* T01 원본 분석 → 원본 구조와 이상치 기준을 T02에 전달
* T02 전처리 → 위 표준 전처리 데이터 구조로 T03/T08에 전달
* T03 통계 → 공통 집계 규격 결과를 T06과 LEAD에 전달
* T04 혼잡 기준 → `quiet/normal/busy` 판정 함수와 기준값 전달
* T06 API → 이 문서의 JSON 구조를 유지해 T05에 전달
* T05 웹 → API가 아직 없어도 동일한 JSON 구조의 더미 데이터를 사용해 먼저 개발
* T08 Excel 갱신 → 갱신 후 동일 전처리/통계 파이프라인을 다시 실행
* LEAD 통합 → 전체 단계의 입력/출력이 이 문서를 지키는지 최종 확인

# 11. 규격 변경 원칙

필드명이나 JSON 구조를 바꿔야 하면 담당자끼리 임의 변경하지 않습니다.

변경 이유를 Discord에 공유하고 팀장 확인 후 이 문서를 먼저 수정한 뒤 코드에 반영합니다.

버전은 현재 `v1`입니다.
