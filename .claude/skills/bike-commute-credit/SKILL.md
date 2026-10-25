---
name: bike-commute-credit
description: 자전거 출퇴근(통근) 크레딧 분기 정산 자동화. 라이딩 횟수 엑셀(Summary + 월별 시트)의 규칙 검증, 입사자/퇴사자 명단 반영, 인증 게시글(예: 네이버 밴드) 랜덤 스크리닝, 분기 정산 엑셀과 검증 오류 목록 생성. "통근 크레딧 정산", "라이딩 횟수 검증", "인증샷 대조", "입퇴사자 반영" 요청에 사용.
---

# 자전거 통근 크레딧 정산

분기마다 라이딩 크레딧을 정산한다. 담당자(HR)가 엑셀을 보고 크레딧을 지급하며, 직원이 청구하지 않는 제도를 가정한다.
규칙 상세는 [references/rules.md](references/rules.md), 인증 게시글 대조 절차는 [references/band-verification.md](references/band-verification.md)를 읽는다.
**규칙의 수치(단가, 연 한도, 근무요일, 마감일, 지급일)는 [references/config.json](references/config.json)** 에 있다. 요약 문장을 믿지 말고 이 파일을 기준으로 답한다.

## 규칙 (기본값 기준)
- 1회 2,000원(왕복 2회). 1인 연 한도 400,000원 — 도달하면 그 분기는 남은 금액까지만 지급하고 이후는 0원.
- 분기 정산·지급: 분기 마지막 달의 익월 25일(Q4는 익년 1월). 입력 마감은 익월 7일.
- 근무일만 인정: OFFICE 월-금 + 공휴일 휴무, STORE 화-일(공휴일 근무).
- 랜덤 스크리닝에서 엑셀과 인증 게시글이 상이하면 **해당 월 전체 무효**. 마감 이후 입력은 제외.

## 사용 원칙
- **실제 데이터(엑셀, 명단, 인증 게시글, 결과 파일)는 개인정보다.** 저장소에 커밋하거나 외부로 보내지 않는다. 결과는 `output/`(gitignore)에 둔다.
- 원본 엑셀은 수정하지 않는다. 변경이 필요하면 사본을 만든다.
- 돈이 오가는 최종 판단(스크리닝 FAIL 확정, 한도 예외, 오류 행 처리)은 담당자가 한다. FAIL/REVIEW는 사용자에게 보여 주고 확정받은 뒤 정산에 반영한다.
- 정산 전에 기준일(`--as-of`)이 입력 마감일 이후인지 확인한다. 마감 전이면 `QUARTER_OPEN` 경고가 뜬다.
- 시트 이름이 `MM_OFFICE`/`MM_STORE`가 아니면 config의 `sheet_suffix`를 바꾼다(`--config`로 다른 설정 파일 지정 가능).

## 워크플로 (분기 정산)
스크립트는 `.claude/skills/bike-commute-credit/scripts/cc.py`. 의존성은 `openpyxl` 하나. (Windows는 `py`, 그 외는 `python3`)

1. **입·퇴사자 반영** — 명단 CSV(`staff_no,name,unit,type,join_date,leave_date`)로 비교한다. `type`은 config의 구분 이름(OFFICE/STORE) 또는 시트 접미사.
   ```
   python cc.py roster --workbook <엑셀> --roster <명단.csv>            # 비교만(dry-run)
   python cc.py roster --workbook <엑셀> --roster <명단.csv> --apply    # 신규 입사자 행을 추가한 사본 생성
   ```
   결과 `output/roster_diff_<연도>.xlsx`를 사용자에게 보여 주고 확인받은 뒤 `--apply`. 퇴사자는 행을 지우지 않는다(과거 크레딧 보존). 이후 검증에서 퇴사일 이후 입력은 자동 제외된다.
2. **규칙 검증** — 오류 목록만 먼저 본다.
   ```
   python cc.py validate --workbook <엑셀> --quarter 3 --as-of 2026-10-25 [--roster 명단.csv] [--snapshot 마감일사본.xlsx]
   ```
3. **랜덤 스크리닝 대상 추출** — 시드를 기록해 재현 가능하다.
   ```
   python cc.py sample --workbook <엑셀> --quarter 3 --rate 0.1 [--seed N] [--all-months]
   ```
4. **인증 게시글 대조** — [references/band-verification.md](references/band-verification.md)의 절차(사용자가 로그인한 브라우저를 읽기 전용으로 확인)로 대상자의 게시글(작성자, 날짜)을 `posts.csv`로 정리하고 대조한다.
   ```
   python cc.py screen-check --workbook <엑셀> --quarter 3 --plan output/screening_plan_2026Q3.csv --posts posts.csv
   ```
   결과 `screening_result_*.csv`의 FAIL/REVIEW를 사용자와 함께 확인해 확정한다.
5. **정산** — 확정된 스크리닝 결과를 반영한다.
   ```
   python cc.py settle --workbook <엑셀> --quarter 3 --as-of 2026-10-25 --screening output/screening_result_2026Q3.csv [--prior-paid 이전지급액.csv] [--snapshot 마감일사본.xlsx]
   ```
   산출물: `output/settlement_<연도>Q<분기>.xlsx`(정산, 부서별 요약, 기준), `output/validation_errors_<연도>Q<분기>.xlsx`(검증 오류 목록).
6. 결과를 요약해 사용자에게 보고한다: 총 지급액, 지급 인원, 한도 도달자, 무효 월, ERROR/WARN 건수와 주요 내용. 지급 전 담당자 확인이 필요한 항목을 분명히 말한다.

## 알아둘 한계
- 엑셀에는 입력 시각이 없어 마감 후 입력은 `--snapshot`(마감일에 저장해 둔 사본)이 있어야 걸러낸다. 없으면 `LATE_UNVERIFIABLE` 안내만 한다.
- 공휴일은 `references/holidays_kr.json`에 있는 날만 반영한다(대한민국 예시). 연도가 바뀌면 이 파일을 갱신하고 휴무 기준을 확인한다.
- 인증 게시판은 로그인이 필요해 자동 수집하지 않는다. 사진 속 내용이 실제 라이딩인지는 사람이 최종 판단한다.
- `roster --apply`는 openpyxl로 저장하므로 수식 계산값이 비어 있다가 Excel에서 열면 다시 계산된다. 저장 후 Excel에서 열어 확인한다.
- 같은 규칙이 `web/index.html`에도 있다. 규칙 로직을 바꾸면 `scripts/cclib.py`와 웹 도구를 함께 고친다.
