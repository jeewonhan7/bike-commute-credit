import csv
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / ".claude" / "skills" / "bike-commute-credit" / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(Path(__file__).parent))

import cclib  # noqa: E402
import json  # noqa: E402
import make_sample  # noqa: E402
import roster as roster_mod  # noqa: E402
import screening  # noqa: E402

HOLIDAYS = SCRIPTS.parent / "references" / "holidays_kr.json"
AS_OF = date(2026, 10, 25)


class BikeCommuteCreditTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.xlsx = make_sample.build(Path(cls.tmp.name) / "sample.xlsx")

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def prepared(self, roster=None):
        book = cclib.load_book(self.xlsx)
        cclib.validate(book, cclib.load_holidays(HOLIDAYS, book.year), AS_OF, roster)
        return book

    def codes(self, book, staff=None):
        return sorted(i.code for i in book.issues if staff is None or i.staff_no == staff)

    def test_dates(self):
        self.assertEqual(cclib.payout_date(2026, 1), date(2026, 4, 25))
        self.assertEqual(cclib.payout_date(2026, 4), date(2027, 1, 25))
        self.assertEqual(cclib.deadline(2026, 3), date(2026, 10, 7))
        self.assertEqual(cclib.deadline(2026, 4), date(2027, 1, 7))

    def test_workday_rules(self):
        book = self.prepared()
        self.assertIn("NON_WORKDAY", self.codes(book, "900001"))  # 토요일
        self.assertIn("NON_WORKDAY", self.codes(book, "900101"))  # 매장직 월요일
        self.assertIn("NON_WORKDAY", self.codes(book, "900002"))  # 사무직 대체공휴일
        self.assertIn("INVALID_VALUE", self.codes(book, "900002"))

    def test_master_and_foreign_rows(self):
        book = self.prepared()
        self.assertIn("MASTER_DUP", self.codes(book))
        self.assertIn("NOT_IN_SUMMARY", self.codes(book, "900999"))
        self.assertEqual(len(book.master), 6)

    def test_settle_q3(self):
        book = self.prepared()
        rows = {r["staff_no"]: r for r in cclib.settle(book, 3)}
        self.assertEqual(rows["900001"]["payout"], 4 * 2000)  # 7/4 제외
        self.assertEqual(rows["900101"]["payout"], 3 * 2000)  # 7/6 제외
        self.assertEqual(rows["900002"]["payout"], 1 * 2000)  # 8/17, 값 3 제외
        self.assertEqual(rows["900003"]["payout"], 4 * 2000)
        self.assertNotIn("900999", rows)

    def test_screening_fail_voids_month(self):
        book = self.prepared()
        screen = {("900003", 9): {"result": "FAIL", "note": ""}}
        rows = {r["staff_no"]: r for r in cclib.settle(book, 3, screen)}
        self.assertEqual(rows["900003"]["payout"], 0)
        self.assertIn("무효", rows["900003"]["note"])

    def test_annual_cap(self):
        book = self.prepared()
        rows = {r["staff_no"]: r for r in cclib.settle(book, 3, prior_paid={"900003": 395_000})}
        r = rows["900003"]
        self.assertEqual(r["gross"], 8000)
        self.assertEqual(r["payout"], 5000)
        self.assertEqual(r["cum_after"], cclib.ANNUAL_CAP)
        self.assertIn("한도", r["note"])

    def test_snapshot_blocks_late_entries(self):
        book = cclib.load_book(self.xlsx)
        snap = {k: v for k, v in cclib.snapshot_index(book).items() if k[0] != "900003"}
        cclib.validate(book, cclib.load_holidays(HOLIDAYS, 2026), AS_OF, None, snap)
        rows = {r["staff_no"]: r for r in cclib.settle(book, 3)}
        self.assertEqual(rows["900003"]["payout"], 0)
        self.assertIn("LATE_ENTRY", self.codes(book, "900003"))

    def test_roster_employment_range(self):
        roster = {"900002": cclib.RosterRow("900002", "가상사무2", "HR", "OFFICE", None, date(2026, 8, 17))}
        book = self.prepared(roster)
        self.assertIn("AFTER_LEAVE", self.codes(book, "900002"))

    def test_roster_diff_and_apply(self):
        roster = cclib.load_roster(self._write_roster())
        book = cclib.load_book(self.xlsx)
        d = roster_mod.diff(book, roster)
        self.assertEqual([r.staff_no for r in d["new_hires"]], ["900005"])
        self.assertEqual([r.staff_no for r in d["leavers"]], ["900002"])
        out = Path(self.tmp.name) / "updated.xlsx"
        roster_mod.apply_new_hires(book, d["new_hires"], out)
        new = cclib.load_book(out)
        self.assertIn("900005", new.master)
        added = [r for r in new.rows if r["staff_no"] == "900005"]
        self.assertEqual({r["month"] for r in added}, set(range(8, 13)))

    def test_screening_sample_and_check(self):
        book = self.prepared()
        plan = screening.sample(book, 3, 0.5, seed=1)
        self.assertEqual(plan, screening.sample(book, 3, 0.5, seed=1))  # 시드 재현
        plan_path = Path(self.tmp.name) / "plan.csv"
        screening.write_csv(plan_path, screening.PLAN_FIELDS,
                            [p for p in plan if p["staff_no"] == "900001"] or
                            [{"staff_no": "900001", "name": "가상사무1", "unit": "Finance", "type": "OFFICE",
                              "month": 7, "entered_days": "", "total_rides": 0, "seed": 1}])
        posts = Path(self.tmp.name) / "posts.csv"
        with open(posts, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["author", "date", "photos"])
            w.writerow(["OFFICE 가상사무1", "2026-07-01", 1])  # 7/2 인증샷 없음
        res = screening.check(book, plan_path, posts)
        self.assertEqual(res[0]["result"], "FAIL")
        self.assertEqual(res[0]["missing_dates"], "2026-07-02")

    def test_config_changes_rules(self):
        """규칙은 config.json에서 바뀐다: 단가 1,000원, 매장직도 월요일 근무."""
        cfg = json.loads(cclib.DEFAULT_CONFIG.read_text(encoding="utf-8"))
        cfg["rate_per_ride"] = 1000
        cfg["kinds"]["STORE"]["workdays"] = [0, 1, 2, 3, 4, 5, 6]
        path = Path(self.tmp.name) / "config_custom.json"
        path.write_text(json.dumps(cfg), encoding="utf-8")
        cclib.set_config(path)
        try:
            book = self.prepared()
            rows = {r["staff_no"]: r for r in cclib.settle(book, 3)}
            self.assertEqual(rows["900101"]["payout"], 5 * 1000)  # 7/6(월)도 인정, 단가 1,000원
            self.assertNotIn("NON_WORKDAY", self.codes(book, "900101"))
        finally:
            cclib.set_config()

    def _write_roster(self):
        p = Path(self.tmp.name) / "roster.csv"
        p.write_text(
            "staff_no,name,unit,type,join_date,leave_date\n"
            "900001,가상사무1,Finance,OFFICE,2020-01-01,\n"
            "900002,가상사무2,HR,OFFICE,2020-01-01,2026-08-17\n"
            "900005,가상신입,Finance,OFFICE,2026-08-10,\n",
            encoding="utf-8-sig")
        return p


if __name__ == "__main__":
    unittest.main()
