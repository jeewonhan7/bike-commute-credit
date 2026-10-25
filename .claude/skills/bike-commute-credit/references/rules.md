# 정산 규칙과 검증 코드

수치는 기본값이며 [config.json](config.json)에서 바꾼다.

## 규칙
| 항목 | 내용 | 설정 |
|---|---|---|
| 단가 | 라이딩 1회 2,000원 (왕복 = 2회) | `rate_per_ride` |
| 연 한도 | 1인 400,000원. 도달하면 끝. 분기 중 한도를 넘으면 남은 금액까지만 지급, 이후 분기는 0원 | `annual_cap` |
| 월/분기 한도 | 없음 | — |
| 정산·지급 | 1-3월 → 4/25, 4-6월 → 7/25, 7-9월 → 10/25, 10-12월 → 익년 1/25 | `payout_day` |
| 입력 마감 | 분기 마지막 달의 익월 7일. 이월 없음, 마감 후 입력 제외 | `deadline_day` |
| 근무일 | OFFICE 월-금, 공휴일 휴무 / STORE 화-일, 공휴일 근무 | `kinds.*.workdays`, `off_on_holidays` |
| 지급 방식 | 직원 청구 없음. 담당자가 엑셀을 보고 크레딧을 지급 | — |
| 스크리닝 | 랜덤 샘플링으로 엑셀 vs 인증 게시글 대조. 상이하면 해당 월 전체 무효 | — |

## 엑셀 구조
- `Summary`: 11행 헤더, 12행부터 직원 마스터, D10에 연도. 월별 칸은 각 월 시트의 Total을 가져오는 수식
- `MM_OFFICE`, `MM_STORE`: E열부터 1~31일, 마지막 열 Total. 칸 값은 그날 라이딩 횟수(1 또는 2). 시트 접미사는 `sheet_suffix`로 바꾼다
- 정산은 Summary 합계가 아니라 **월별 시트의 일자별 원본 값**을 다시 읽어 계산한다(비근무일 입력이 합계에 섞여 있기 때문)
- 샘플은 `py tests/make_sample.py sample/sample.xlsx`로 만든다

## 검증 코드
| 코드 | 심각도 | 처리 | 의미 |
|---|---|---|---|
| NON_WORKDAY | ERROR | 제외 | 근무일이 아닌 날의 입력 |
| INVALID_VALUE | ERROR | 제외 | 1·2가 아닌 값, 숫자가 아닌 값 |
| FUTURE_DATE | ERROR | 제외 | 정산 기준일보다 미래 날짜 |
| BEFORE_JOIN / AFTER_LEAVE | ERROR | 제외 | 입사일 이전 / 퇴사일 이후 (`--roster` 필요) |
| LATE_ENTRY | ERROR | 제외 | 마감일 사본(`--snapshot`)에 없던 입력 |
| DUP_IN_MONTH | ERROR | 제외 | 같은 달 시트에 같은 사번이 두 번 나옴(두 번째 행 제외) |
| NOT_IN_SUMMARY | ERROR / WARN | 제외 / 확인 | Summary에 없는 사번. 라이딩이 있으면 ERROR, 없으면 퇴사자 정리 대상 WARN |
| MASTER_DUP | WARN | 1명으로 계산 | Summary 사번 중복 행 |
| NAME_MISMATCH / UNIT_MISMATCH | WARN | 확인 | Summary와 월별 시트의 이름·부서 불일치 |
| KIND_CONFLICT | WARN | 확인 | 한 사람이 여러 구분의 시트에 나옴 |
| TEXT_NUMBER | WARN | 숫자로 인정 | 숫자가 텍스트로 저장됨 |
| QUARTER_OPEN | WARN | 마감 후 재실행 | 기준일이 입력 마감일 이전 |
| LATE_UNVERIFIABLE | INFO | 참고 | 마감 후 입력 여부를 확인할 수 없음(`--snapshot` 없음) |
| MISSING_SHEET / SHEET_LAYOUT | WARN / ERROR | 확인 | 월별 시트 누락, 일자 헤더 이상 |
| SCREENING_REVIEW | WARN | 결과 확정 | 스크리닝이 REVIEW 상태 |

## 가정(규칙에 명시되지 않아 정한 것)
- 스크리닝 "상이"의 기준: 엑셀에 입력한 **모든 라이딩 날짜**에 그 직원의 인증 게시글이 있어야 PASS. 사진 수가 라이딩 횟수와 같아야 한다고 볼 때는 `--strict-photos`.
- 한도에 도달한 분기의 초과분은 지급하지 않고 비고에 미지급액을 표시한다.
- 중복 사번은 첫 행만 인정한다.
