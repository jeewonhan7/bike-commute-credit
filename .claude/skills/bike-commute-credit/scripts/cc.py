#!/usr/bin/env python3
"""자전거 통근 크레딧 정산 CLI.

  validate      규칙 검증 → 검증 오류 목록 엑셀
  settle        검증 + 분기 정산 → 정산 엑셀 + 검증 오류 목록 엑셀
  roster        HR 명단과 비교(기본 dry-run), --apply 로 신규 입사자 추가 사본 생성
  sample        밴드 대조용 랜덤 스크리닝 대상 추출
  screen-check  추출 대상과 밴드 게시글 CSV 대조 → 스크리닝 결과 CSV
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import cclib  # noqa: E402
import report  # noqa: E402
import roster as roster_mod  # noqa: E402
import screening  # noqa: E402

DEFAULT_HOLIDAYS = HERE.parent / "references" / "holidays_kr.json"


def add_common(p, quarter=True):
    p.add_argument("--workbook", required=True, help="정산 엑셀 파일")
    p.add_argument("--year", type=int, help="연도(기본: Summary D10)")
    p.add_argument("--as-of", type=date.fromisoformat, default=date.today(), help="정산 기준일 YYYY-MM-DD(기본: 오늘)")
    p.add_argument("--roster", help="HR 명단 CSV(입·퇴사일 검증용)")
    p.add_argument("--holidays", default=str(DEFAULT_HOLIDAYS), help="공휴일 JSON")
    p.add_argument("--config", help="규칙 설정 JSON(기본: references/config.json)")
    p.add_argument("--snapshot", help="입력 마감일에 저장해 둔 엑셀 사본(마감 후 입력 제외용)")
    p.add_argument("--out-dir", default="output", help="결과 저장 폴더(기본: ./output)")
    if quarter:
        p.add_argument("--quarter", type=int, choices=[1, 2, 3, 4], required=True)


def prepare(a, quarter=None):
    book = cclib.load_book(a.workbook, a.year)
    holidays = cclib.load_holidays(a.holidays, book.year)
    roster = cclib.load_roster(a.roster) if a.roster else None
    snap = None
    if a.snapshot:
        sb = cclib.load_book(a.snapshot, book.year)
        snap = cclib.snapshot_index(sb)
    cclib.validate(book, holidays, a.as_of, roster, snap)
    if quarter:
        cclib.quarter_notes(book, quarter, a.as_of, snap is not None)
    return book


def cmd_validate(a):
    book = prepare(a, a.quarter)
    months = cclib.QUARTER_MONTHS[a.quarter]
    out = Path(a.out_dir) / f"validation_errors_{book.year}Q{a.quarter}.xlsx"
    n = report.write_errors(out, book, months)
    _print_counts(book, months)
    print(f"검증 오류 목록: {out} ({n}건)")


def cmd_settle(a):
    book = prepare(a, a.quarter)
    months = cclib.QUARTER_MONTHS[a.quarter]
    screen = cclib.load_screening(a.screening) if a.screening else {}
    for (no, m), v in screen.items():
        if v["result"] == "REVIEW" and cclib.quarter_of(m) == a.quarter:
            p = book.master.get(no)
            book.add("WARN", "SCREENING_REVIEW", f"{m}월 스크리닝이 '확인필요' 상태입니다. 결과 확정 전에는 무효 처리하지 않았습니다. {v['note']}",
                     "결과 확정", staff_no=no, name=p.name if p else "", unit=p.unit if p else "", month=m)
    prior = cclib.load_prior_paid(a.prior_paid) if a.prior_paid else None
    rows = cclib.settle(book, a.quarter, screen, prior)
    voids = sum(1 for v in screen.values() if v["result"] == "FAIL")
    out_dir = Path(a.out_dir)
    s_path = out_dir / f"settlement_{book.year}Q{a.quarter}.xlsx"
    e_path = out_dir / f"validation_errors_{book.year}Q{a.quarter}.xlsx"
    report.write_settlement(s_path, book, a.quarter, rows, a.as_of, voids)
    n = report.write_errors(e_path, book, months)
    _print_counts(book, months)
    total = sum(r["payout"] for r in rows)
    paid = sum(1 for r in rows if r["payout"] > 0)
    print(f"{book.year}년 {a.quarter}분기 총 지급액 {total:,}원 / 지급 대상 {paid}명 / 지급일 {cclib.payout_date(book.year, a.quarter)}")
    print(f"정산 엑셀: {s_path}\n검증 오류 목록: {e_path} ({n}건)")


def _print_counts(book, months):
    c = {}
    for i in book.issues:
        if i.month is None or i.month in months:
            c[i.severity] = c.get(i.severity, 0) + 1
    print("검증 결과: " + ", ".join(f"{k} {c.get(k, 0)}건" for k in ("ERROR", "WARN", "INFO")))


def cmd_roster(a):
    book = cclib.load_book(a.workbook, a.year)
    roster = cclib.load_roster(a.roster)
    d = roster_mod.diff(book, roster)
    out_dir = Path(a.out_dir)
    diff_path = out_dir / f"roster_diff_{book.year}.xlsx"
    roster_mod.write_diff(diff_path, d)
    print(f"신규 입사자 {len(d['new_hires'])}명, 퇴사자 {len(d['leavers'])}명, "
          f"이름 변경 {len(d['name_changes'])}건, 부서 변경 {len(d['unit_changes'])}건, "
          f"명단에 없는 직원 {len(d['not_in_roster'])}명")
    print(f"비교 결과: {diff_path}")
    if a.apply:
        out = out_dir / f"{Path(a.workbook).stem}_roster_updated.xlsx"
        for line in roster_mod.apply_new_hires(book, d["new_hires"], out):
            print(" -", line)
        print(f"사본 저장(원본은 수정하지 않음): {out}")
        print("주의: 수식 계산값은 Excel에서 열 때 다시 계산됩니다. 퇴사자는 행을 지우지 않고 --roster 검증(퇴사일 이후 입력 제외)으로 처리합니다.")


def cmd_sample(a):
    book = prepare(a, None)
    plan = screening.sample(book, a.quarter, a.rate, a.seed, a.all_months)
    out = Path(a.out_dir) / f"screening_plan_{book.year}Q{a.quarter}.csv"
    screening.write_csv(out, screening.PLAN_FIELDS, plan)
    print(f"스크리닝 대상 {len(plan)}건(seed={a.seed}, 비율={a.rate:.0%}) → {out}")


def cmd_screen_check(a):
    book = prepare(a, None)
    res = screening.check(book, a.plan, a.posts, a.strict_photos)
    out = Path(a.out_dir) / f"screening_result_{book.year}Q{a.quarter}.csv"
    screening.write_csv(out, screening.RESULT_FIELDS, res)
    c = {k: sum(1 for r in res if r["result"] == k) for k in ("PASS", "FAIL", "REVIEW")}
    print(f"대조 결과 PASS {c['PASS']}, FAIL {c['FAIL']}, REVIEW {c['REVIEW']} → {out}")
    print("FAIL/REVIEW는 사람이 게시글을 직접 확인한 뒤 결과를 확정하세요. FAIL이 확정되면 해당 월은 전체 무효입니다.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("validate")
    add_common(p)
    p.set_defaults(fn=cmd_validate)

    p = sub.add_parser("settle")
    add_common(p)
    p.add_argument("--screening", help="스크리닝 결과 CSV(staff_no,month,result)")
    p.add_argument("--prior-paid", help="이전 분기까지 실제 지급액 CSV(staff_no,paid). 없으면 시트로 재계산")
    p.set_defaults(fn=cmd_settle)

    p = sub.add_parser("roster")
    p.add_argument("--workbook", required=True)
    p.add_argument("--year", type=int)
    p.add_argument("--config", help="규칙 설정 JSON(기본: references/config.json)")
    p.add_argument("--roster", required=True, help="명단 CSV(staff_no,name,unit,type,join_date,leave_date)")
    p.add_argument("--out-dir", default="output")
    p.add_argument("--apply", action="store_true", help="신규 입사자를 추가한 사본을 저장")
    p.set_defaults(fn=cmd_roster)

    p = sub.add_parser("sample")
    add_common(p)
    p.add_argument("--rate", type=float, default=0.1, help="구분별 추출 비율(기본 0.1, 최소 1명)")
    p.add_argument("--seed", type=int, default=None, help="난수 시드(기록용, 기본: 기준일 YYYYMMDD)")
    p.add_argument("--all-months", action="store_true", help="추출된 직원의 분기 내 모든 월을 대조")
    p.set_defaults(fn=cmd_sample)

    p = sub.add_parser("screen-check")
    add_common(p)
    p.add_argument("--plan", required=True)
    p.add_argument("--posts", required=True, help="밴드 게시글 CSV(author,date[,photos])")
    p.add_argument("--strict-photos", action="store_true", help="입력한 라이딩 횟수만큼 사진 수도 요구")
    p.set_defaults(fn=cmd_screen_check)

    a = ap.parse_args()
    if a.config:
        cclib.set_config(a.config)
    if a.cmd == "sample" and a.seed is None:
        a.seed = int(a.as_of.strftime("%Y%m%d"))
    a.fn(a)


if __name__ == "__main__":
    main()
