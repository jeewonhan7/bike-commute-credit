"""네이버 밴드 인증샷 랜덤 스크리닝: 대상 추출(sample)과 밴드 게시글 대조(check)."""
from __future__ import annotations

import csv
import math
import random
from collections import defaultdict
from datetime import date
from pathlib import Path

from cclib import QUARTER_MONTHS, Book, clean, parse_date

PLAN_FIELDS = ["staff_no", "name", "unit", "type", "month", "entered_days", "total_rides", "seed"]
RESULT_FIELDS = ["staff_no", "name", "month", "result", "missing_dates", "note"]


def _entered(book: Book) -> dict:
    """(staff_no, month) -> [(date, rides)] (유효한 입력만)"""
    out = defaultdict(list)
    for e in book.entries:
        if e.valid:
            out[(e.staff_no, e.date.month)].append((e.date, e.rides))
    for v in out.values():
        v.sort()
    return out


def sample(book: Book, quarter: int, rate: float, seed: int, all_months: bool = False) -> list[dict]:
    """분기 내 라이딩이 있는 직원 중 구분(config의 kinds)별로 비율만큼(최소 1명) 무작위 추출한다."""
    months = QUARTER_MONTHS[quarter]
    entered = _entered(book)
    kind_of = {r["staff_no"]: r["kind"] for r in book.rows}
    pool = defaultdict(list)
    for no in sorted({no for (no, m) in entered if m in months}):
        if no in book.master:
            pool[kind_of.get(no, "")].append(no)
    rng = random.Random(seed)
    plan = []
    for kind in sorted(pool):
        people = pool[kind]
        picked = rng.sample(people, min(len(people), max(1, math.ceil(len(people) * rate))))
        for no in sorted(picked):
            p = book.master[no]
            ms = [m for m in months if entered.get((no, m))]
            for m in (ms if all_months else [rng.choice(ms)]):
                days = entered[(no, m)]
                plan.append({
                    "staff_no": no, "name": p.name, "unit": p.unit, "type": kind, "month": m,
                    "entered_days": ", ".join(f"{d.month}/{d.day}({r})" for d, r in days),
                    "total_rides": sum(r for _, r in days), "seed": seed,
                })
    return plan


def write_csv(path, fields, rows) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def _compact(s: str) -> str:
    return "".join(clean(s).split()).lower()


def check(book: Book, plan_path, posts_path, strict_photos: bool = False) -> list[dict]:
    """계획의 (직원, 월)마다 엑셀에 입력한 모든 날짜에 그 사람의 밴드 게시글이 있는지 대조한다.

    posts CSV: author,date[,photos]   (author = 밴드 표시 이름, date = YYYY-MM-DD)
    """
    posts = defaultdict(lambda: defaultdict(int))  # author -> date -> photos
    with open(posts_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            d = parse_date(row.get("date", ""))
            if d:
                n = int(float(row.get("photos") or 1))
                posts[_compact(row.get("author", ""))][d] += n
    entered = _entered(book)
    results = []
    with open(plan_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            no, month, name = row["staff_no"], int(row["month"]), row["name"]
            key = _compact(name)
            authors = [a for a in posts if key and (key in a or a in key)]
            note = ""
            if not authors:
                note = "밴드 게시글에서 이 직원의 작성자를 찾지 못함(이름 표기 확인 필요)"
            elif len(authors) > 1:
                note = f"작성자 후보가 여러 명({len(authors)}): 이름이 겹치는지 확인 필요"
            days = defaultdict(int)
            for a in authors:
                for d, n in posts[a].items():
                    days[d] += n
            missing = []
            for d, rides in entered.get((no, month), []):
                if days.get(d, 0) < (rides if strict_photos else 1):
                    missing.append(d.isoformat())
            results.append({
                "staff_no": no, "name": name, "month": month,
                "result": "FAIL" if missing and authors else ("REVIEW" if not authors else "PASS"),
                "missing_dates": " ".join(missing), "note": note,
            })
    return results
