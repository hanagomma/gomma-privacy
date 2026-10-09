import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "samples"))

import auto_journal as aj  # noqa: E402
import make_samples  # noqa: E402

S = HERE / "samples"


class SampleRun(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        make_samples.main()
        cls.out = Path(tempfile.mkdtemp()) / "out.xlsx"
        cls.j = aj.run(S / "통장_202609.xlsx", S / "거래처원장.xlsx",
                       [S / "매출세금계산서.xlsx"], [S / "매입세금계산서.xlsx"],
                       [S / "급여대장_202608.xlsx", S / "급여대장_202609.xlsx"],
                       [S / "사업소득_202609.xlsx"], cls.out, withholding=[S / "예수금원장.xlsx"])
        cls.by_desc = {}
        for m, e in zip(cls.j.match_log, cls.j.entries):
            cls.by_desc[m["통장표시"]] = (m, e)

    def entry(self, key):
        for desc, (m, e) in self.by_desc.items():
            if key in desc:
                return m, e
        self.fail(f"{key} 없음")

    def lines(self, e):
        return {(l.side, l.account): l.amount for l in e.lines}

    def test_all_balanced(self):
        self.assertTrue(all(e.balanced() for e in self.j.entries))
        self.assertTrue(self.out.exists())

    def test_truncated_company_name(self):
        m, e = self.entry("한빛상")
        self.assertEqual(m["거래처"], "(주)한빛상사")
        self.assertEqual(m["거래처코드"], "00101")
        self.assertEqual(self.lines(e)[("대변", "외상매출금")], 1_100_000)
        self.assertFalse(e.review)

    def test_person_name_resolves_to_company(self):
        m, e = self.entry("박푸른")
        self.assertEqual(m["인식방법"], "대표자명일치")
        self.assertEqual(m["거래처"], "푸른유통")
        self.assertFalse(e.review)

    def test_new_partner_from_invoice_ceo(self):
        m, _ = self.entry("한새롬")
        self.assertEqual(m["거래처"], "(주)새로운고객")
        self.assertEqual(m["원장등록"], "미등록(신규)")
        self.assertIn("(주)새로운고객", [p.name for p in self.j.new_partners.values()])

    def test_multi_invoice_sum(self):
        m, e = self.entry("동해기계")
        self.assertIn("2건 합계", e.basis)

    def test_purchase(self):
        _, e = self.entry("서울자재")
        self.assertEqual(self.lines(e)[("차변", "외상매입금")], 880_000)

    def test_purchase_with_transfer_fee(self):
        # 세금계산서 1,100,000 + 송금수수료 500 = 출금 1,100,500
        m, e = self.entry("대한부품")
        ln = self.lines(e)
        self.assertEqual(m["거래처"], "(주)대한부품")
        self.assertEqual(ln[("차변", "외상매입금")], 1_100_000)
        self.assertEqual(ln[("차변", "지급수수료")], 500)
        self.assertEqual(ln[("대변", "보통예금")], 1_100_500)
        self.assertFalse(e.review)
        self.assertTrue(e.balanced())

    def test_payroll(self):
        # 급여는 미지급비용을 털어낸다: 차변 미지급비용 / 대변 보통예금
        _, e = self.entry("홍길동")
        self.assertIn("2026-09", e.lines[0].memo)
        self.assertEqual(self.lines(e), {("차변", "미지급비용"): 2_636_100, ("대변", "보통예금"): 2_636_100})
        self.assertFalse(e.review)

    def entries_for(self, name):
        return [e for m, e in zip(self.j.match_log, self.j.entries) if name in m["통장표시"]]

    def test_payroll_with_transfer_fee(self):
        # 실지급액 2,219,120 + 송금수수료 500 = 출금 2,219,620
        e = next(e for e in self.entries_for("김영희") if e.lines[0].amount == 2_219_120)
        self.assertEqual(self.lines(e), {("차변", "미지급비용"): 2_219_120, ("차변", "지급수수료"): 500,
                                         ("대변", "보통예금"): 2_219_620})

    def test_dual_payroll_and_business_person_flagged(self):
        self.assertEqual(self.j.dual_names, ["김영희"])
        es = self.entries_for("김영희")
        self.assertEqual(len(es), 2)                          # 급여 지급 + 사업소득 지급
        for e in es:
            self.assertIn("급여·사업소득 양쪽", e.review)        # 노란색 + 확인필요 시트
        biz = next(e for e in es if e.lines[0].amount == 290_100)
        self.assertEqual(self.lines(biz), {("차변", "미지급비용"): 290_100, ("대변", "보통예금"): 290_100})
        _, e = self.entry("이강사")                            # 사업소득만 있는 사람은 표시 안 함
        self.assertFalse(e.review)

    def test_business_income(self):
        # 차인지급액 967,000 + 송금수수료 500 = 출금 967,500
        _, e = self.entry("이강사")
        self.assertEqual(self.lines(e), {("차변", "미지급비용"): 967_000, ("차변", "지급수수료"): 500,
                                         ("대변", "보통예금"): 967_500})
        self.assertFalse(e.review)

    def test_withholding_tax_matches_prior_month(self):
        _, e = self.entry("원천세")
        self.assertFalse(e.review, e.review)
        _, e = self.entry("지방소득세")
        self.assertFalse(e.review, e.review)

    def test_insurance_split(self):
        _, e = self.entry("국민연금")
        self.assertEqual(self.lines(e)[("차변", "예수금")], 247_500)
        self.assertEqual(self.lines(e)[("차변", "세금과공과")], 247_500)
        _, e = self.entry("건강보험")
        self.assertEqual(self.lines(e)[("차변", "예수금")], 220_210)   # 예수금원장 8월 건강+장기요양
        self.assertEqual(self.lines(e)[("차변", "보험료")], 220_210)
        self.assertIn("예수금원장 2026-08", e.basis)
        _, e = self.entry("고용보험")
        self.assertEqual(self.lines(e)[("차변", "예수금")], 49_500)
        self.assertEqual(self.lines(e)[("차변", "보험료")], 63_250)
        # 산재: 예수금 없이 전액 보험료(821)
        _, e = self.entry("산재보험")
        self.assertEqual(self.lines(e), {("차변", "보험료"): 38_500, ("대변", "보통예금"): 38_500})
        self.assertFalse(e.review)

    def test_kt_with_invoice_mismatch_and_rules(self):
        m, e = self.entry("KT")
        self.assertEqual(m["거래처"], "케이티")
        self.assertEqual(m["인식방법"], "별칭일치")
        self.assertFalse(e.review)
        self.assertIn(("차변", "외상매입금"), self.lines(e))
        _, e = self.entry("한국전력")
        self.assertIn(("차변", "수도광열비"), self.lines(e))
        _, e = self.entry("예금이자")
        self.assertIn(("대변", "이자수익"), self.lines(e))

    def test_unknown(self):
        # 거래처 연결 안 된 출금 → 외상매입금 (확인필요)
        _, e = self.entry("알수없음")
        self.assertEqual(self.lines(e), {("차변", "외상매입금"): 300_000, ("대변", "보통예금"): 300_000})
        self.assertTrue(e.review)
        # 끝자리 500원이면 거래처를 못 찾아도 송금수수료 분리
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 5, 1), "(주)모르는카", 0, 15_000_500)])
        e = j.entries[0]
        self.assertEqual([(l.side, l.account, l.amount) for l in e.lines],
                         [("차변", "외상매입금", 15_000_000), ("차변", "지급수수료", 500), ("대변", "보통예금", 15_000_500)])
        self.assertTrue(e.review)

    def test_check_card_excluded(self):
        # '체크' 표시 행은 전표에 없고 제외 목록에만
        self.assertEqual(sorted(t.amount for t in self.j.excluded), [8_700, 45_000])
        self.assertFalse(any("GS25" in m["통장표시"] or "쿠팡" in m["통장표시"] for m in self.j.match_log))
        from openpyxl import load_workbook
        wb = load_workbook(self.out)
        voucher = [r for r in wb["일반전표(통장)"].iter_rows(values_only=True)]
        self.assertFalse(any(r[8] and ("GS25" in r[8] or "쿠팡" in r[8]) for r in voucher))
        self.assertEqual(wb["제외(체크카드)"].max_row, 3)

    def test_semusarang_upload_file(self):
        import xlrd
        sh = xlrd.open_workbook(str(self.j.upload_path)).sheet_by_index(0)
        self.assertEqual(sh.cell_value(0, 0), "엑셀자료 일반전표전송")          # 양식 머리 유지
        self.assertIn("년도월일", sh.cell_value(9, 0))
        rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(10, sh.nrows)]
        rows = [r for r in rows if r[0] != ""]
        self.assertEqual(len(rows), sum(len(e.lines) for e in self.j.entries))
        # 일자별 차변(3) 합계 = 대변(4) 합계
        by_day = {}
        for r in rows:
            d = by_day.setdefault(int(r[0]), [0, 0])
            d[0 if int(r[1]) == 3 else 1] += r[8]
        self.assertTrue(all(dr == cr for dr, cr in by_day.values()))
        # 체크카드 행 없음, 송금수수료는 831
        self.assertFalse(any("GS25" in str(r[7]) or "쿠팡" in str(r[7]) for r in rows))
        fee = [r for r in rows if r[7] == "송금수수료"]
        self.assertTrue(fee and all(int(r[2]) == 831 and r[8] == 500 for r in fee))
        # 거래처 코드·사업자번호·대표자
        hb = next(r for r in rows if r[5] == "(주)한빛상사")
        self.assertEqual((hb[4], hb[15], hb[21]), ("00101", "220-81-11111", "이한빛"))
        self.assertEqual(int(rows[0][0]), 20260904)

    def test_date_range(self):
        txns = aj.load_bank(S / "통장_202609.xlsx", aj.date(2026, 9, 22), aj.date(2026, 9, 28))
        self.assertTrue(txns and all(aj.date(2026, 9, 22) <= t.date <= aj.date(2026, 9, 28) for t in txns))

    def test_tax_office_deposit_is_receivable_clear(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 9, 14), "테스트세무서", 12_345_670, 0)])
        e = j.entries[0]
        self.assertEqual([(l.side, l.account, l.amount) for l in e.lines],
                         [("차변", "보통예금", 12_345_670), ("대변", "미수금", 12_345_670)])
        self.assertFalse(e.review)

    def test_refund_pair(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.journalize_bank([
            aj.Txn(1, aj.date(2026, 3, 24), "주식회사 가람모터 타행송금", 0, 1_000_500, party="주식회사 가람모터"),
            aj.Txn(2, aj.date(2026, 4, 6), "(주)다라핸즈 쇼룸 당행송금", 0, 1_000_000, party="(주)다라핸즈 쇼룸"),
            aj.Txn(3, aj.date(2026, 4, 7), "주식회사가람모터 타행이체", 1_000_000, 0, party="주식회사가람모터"),
            aj.Txn(4, aj.date(2026, 4, 7), "다라핸즈 대체", 1_000_000, 0, party="다라핸즈 주식회사 다라핸즈"),
            aj.Txn(5, aj.date(2026, 4, 8), "다른회사 타행송금", 0, 1_000_500, party="다른회사"),
        ])
        got = [[(l.side, l.account, l.amount) for l in e.lines] for e in j.entries]
        self.assertEqual(got[0], [("차변", "예수금", 1_000_000), ("차변", "지급수수료", 500), ("대변", "보통예금", 1_000_500)])
        self.assertEqual(got[1], [("차변", "예수금", 1_000_000), ("대변", "보통예금", 1_000_000)])
        self.assertEqual(got[2], [("차변", "보통예금", 1_000_000), ("대변", "예수금", 1_000_000)])
        self.assertEqual(got[3], [("차변", "보통예금", 1_000_000), ("대변", "예수금", 1_000_000)])
        self.assertNotIn("예수금", [l.account for l in j.entries[4].lines])   # 반환 없는 송금은 해당 없음

    def test_review_keywords(self):
        cfg = aj.load_config()
        cfg["review_keywords"] = ["다라핸즈"]
        j = aj.Journalizer(cfg, [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 4, 6), "다라핸즈", 0, 1_000_000, party="(주)다라핸즈"),
                           aj.Txn(2, aj.date(2026, 4, 7), "다라핸즈", 1_000_000, 0, party="다라핸즈"),
                           aj.Txn(3, aj.date(2026, 4, 7), "한국전력", 0, 50_000, party="한국전력")])
        self.assertEqual([("사용자 확인 예정" in e.review) for e in j.entries], [True, True, False])

    def test_extra_rules_pattern_and_party(self):
        cfg = aj.load_config()
        cfg["extra_rules"] = [
            {"side": "입금", "keywords": ["테스트상사"], "account": "보통예금", "memo": "자기 계좌 대체"},
            {"side": "입금", "pattern": "\\d{2,3}[가-힣]\\d{3,4}", "account": "외상매출금", "partner_from_party": True},
            {"side": "입금", "pattern": "^[A-Za-z .]+$", "on": "party", "account": "외상매출금", "partner_from_party": True}]
        j = aj.Journalizer(cfg, [], [], [], [])
        j.journalize_bank([
            aj.Txn(1, aj.date(2026, 5, 1), "주식회사 테스트상사 대체", 1_000, 0, party="주식회사 테스트상사 주식회사 테스트상사"),
            aj.Txn(2, aj.date(2026, 5, 2), "12가3456 TEST BUYER 대체", 2_000, 0, party="12가3456 TEST BUYER"),
            aj.Txn(3, aj.date(2026, 5, 3), "JOHN DOE JOHN DOE 타행이체", 3_000, 0, party="JOHN DOE JOHN DOE"),
            aj.Txn(4, aj.date(2026, 5, 4), "홍길동 타행이체", 4_000, 0, party="홍길동")])
        e1, e2, e3, e4 = j.entries
        self.assertEqual([(l.side, l.account) for l in e1.lines], [("차변", "보통예금"), ("대변", "보통예금")])
        self.assertEqual((e2.lines[1].account, e2.lines[1].partner.name), ("외상매출금", "TEST BUYER"))
        self.assertEqual(e3.lines[1].partner.name, "JOHN DOE")
        self.assertFalse(e1.review or e2.review or e3.review)
        self.assertTrue(e4.review)                       # 한글 개인 이름은 규칙 대상 아님 → 확인필요

    def test_negative_debit_refund(self):
        cfg = aj.load_config()
        cfg["extra_rules"] = [{"side": "입금", "keywords": ["가나손보"], "account": "보험료", "negative_debit": True}]
        j = aj.Journalizer(cfg, [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 3, 24), "가나손보변경", 123_400, 0)])
        e = j.entries[0]
        self.assertEqual([(l.side, l.account, l.amount) for l in e.lines],
                         [("차변", "보통예금", 123_400), ("차변", "보험료", -123_400)])
        self.assertTrue(e.balanced())
        self.assertFalse(e.review)

    def test_tax_certificate(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.tax_payments = [aj.TaxPayment("2025", "법인세", aj.date(2026, 3, 31), 23_456_780),
                          aj.TaxPayment("2026", "근로소득세(갑)", aj.date(2026, 7, 13), 2_420_590),
                          aj.TaxPayment("2026", "배당소득세", aj.date(2026, 7, 13), 2_597_260),
                          aj.TaxPayment("2025", "사업소득세", aj.date(2026, 5, 13), 727_290)]
        j.journalize_bank([aj.Txn(1, aj.date(2026, 3, 31), "국고_주식회사 공과금", 0, 23_456_780),
                           aj.Txn(2, aj.date(2026, 7, 13), "국고_주식회사 공과금", 0, 2_420_590),
                           aj.Txn(3, aj.date(2026, 7, 13), "국고_주식회사 공과금", 0, 2_597_260),
                           aj.Txn(4, aj.date(2026, 5, 13), "국고_주식회사 공과금", 0, 730_180),
                           aj.Txn(5, aj.date(2026, 8, 1), "국고_주식회사 공과금", 0, 111_111)])
        acc = [(e.lines[0].account, bool(e.review)) for e in j.entries]
        self.assertEqual(acc, [("미지급세금", False), ("예수금", False), ("외상매입금", True),
                               ("예수금", False), ("외상매입금", True)])
        # 5/13: 증명서 727,290 + 차액 2,890 → 잡손실
        self.assertEqual([(l.account, l.amount) for l in j.entries[3].lines],
                         [("예수금", 727_290), ("잡손실", 2_890), ("보통예금", 730_180)])
        self.assertEqual(aj.load_config()["accounts"]["미지급세금"], "261")

    def test_local_tax_with_surcharge(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.tax_payments = [aj.TaxPayment("2025-07", "지방소득세(특별징수)", aj.date(2026, 3, 31), 72_290),
                          aj.TaxPayment("2026-06", "자동차세(자동차)", aj.date(2026, 7, 29), 266_170, note="12가3456"),
                          aj.TaxPayment("2026-06", "자동차세(자동차)", aj.date(2026, 7, 29), 27_780, note="34나5678")]
        j.journalize_bank([aj.Txn(1, aj.date(2026, 3, 31), "테스트구 지방세 공과금", 0, 72_290),
                           aj.Txn(2, aj.date(2026, 7, 29), "테스트구 지방세 공과금", 0, 28_610),
                           aj.Txn(3, aj.date(2026, 7, 29), "테스트구 지방세 공과금", 0, 274_150)])
        e1, e2, e3 = j.entries
        self.assertEqual((e1.lines[0].account, e1.review), ("예수금", ""))
        self.assertIn("34나5678", e2.lines[0].memo)            # 3% 더 많은 출금 → 가까운 납부와 연결
        self.assertIn("12가3456", e3.lines[0].memo)
        # 증명서 금액은 세목 계정(자동차세는 미지정 → 외상매입금), 차액은 잡손실 930
        self.assertEqual([(l.side, l.account, l.amount) for l in e2.lines],
                         [("차변", "외상매입금", 27_780), ("차변", "잡손실", 830), ("대변", "보통예금", 28_610)])
        self.assertEqual(aj.load_config()["accounts"]["잡손실"], "930")
        self.assertTrue(e2.balanced() and e3.balanced())

    def test_corporate_local_income_tax(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.tax_payments = [aj.TaxPayment("2025", "지방소득세(법인소득)", aj.date(2026, 4, 29), 3_456_780)]
        j.journalize_bank([aj.Txn(1, aj.date(2026, 4, 29), "테스트구 지방세", 0, 3_456_780)])
        self.assertEqual((j.entries[0].lines[0].account, j.entries[0].review), ("미지급세금", ""))

    def test_rule_partner_code(self):
        cfg = aj.load_config()
        cfg["extra_rules"] = [{"side": "입금", "keywords": ["테스트상사"], "account": "보통예금", "partner_code": "99000"}]
        j = aj.Journalizer(cfg, [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 5, 1), "주식회사 테스트상사 대체", 1_000, 0)])
        bank, other = j.entries[0].lines
        self.assertEqual((bank.side, bank.partner), ("차변", None))
        self.assertEqual((other.side, other.account, other.partner.code), ("대변", "보통예금", "99000"))

    def test_bank_partner_code(self):
        cfg = aj.load_config()
        cfg["bank_partner_code"] = "99001"
        cfg["extra_rules"] = [{"side": "입금", "keywords": ["테스트상사"], "account": "보통예금", "partner_code": "99000"}]
        j = aj.Journalizer(cfg, [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 5, 1), "주식회사 테스트상사 대체", 1_000, 0),
                           aj.Txn(2, aj.date(2026, 5, 2), "모르는곳", 0, 2_000)])
        self.assertEqual([(l.side, l.partner.code if l.partner else "") for l in j.entries[0].lines],
                         [("차변", "99001"), ("대변", "99000")])
        self.assertEqual(j.entries[1].lines[-1].partner.code, "99001")

    def test_company_match_not_partial(self):
        t = aj.party_tokens
        self.assertFalse(aj.company_match("나모터스", t("(주)가나모터스 (주)가나모터스 타행송금")))
        self.assertFalse(aj.company_match("사모터스", t("주식회사아사모터스 타행송금")))
        self.assertTrue(aj.company_match("사모터스", t("나영(사모터스) 나영(사모터스) 타행송금")))
        self.assertTrue(aj.company_match("다온카", t("최가람(다온카) 타행송금")))
        self.assertTrue(aj.company_match("별관세사무소", t("김가온(별 관세사무 타행송금")))       # 잘린 이름
        self.assertTrue(aj.company_match("주식회사 하늘모터스", t("주식회사하늘모터스 타행송금")))
        self.assertFalse(aj.company_match("투모터스(TWO모터스)", t("(주)에스투모터스 타행송금")))
        self.assertFalse(aj.company_match("포스트타워101호", t("강변 타워 관리단 당행송금")))
        self.assertTrue(aj.company_match("(주)누리 본점", t("(주)누리 (주)누리 타행송금")))
        self.assertFalse(aj.company_match("주식회사 에스누리", t("(주)누리 (주)누리 타행송금")))

    def test_same_company_name_resolved_by_ceo(self):
        L = [aj.Partner("사모터스", "00501", ceo="고길동"), aj.Partner("사모터스", "00502", ceo="나영"),
             aj.Partner("별빛카", "00601", ceo="이영희"), aj.Partner("달빛카", "00602", ceo="이영희")]
        j = aj.Journalizer(aj.load_config(), L, [], [], [])
        self.assertEqual([p.code for p, _ in j.identify("나영(사모터스) 타행송금")], ["00502"])
        self.assertEqual([p.code for p, _ in j.identify("고길동 타행송금")], ["00501"])
        self.assertEqual(len(j.identify("이영희 타행송금")), 2)        # 동명 대표 → 후보 2곳 (연결 안 함)
        L2 = [aj.Partner("한결모터스", "00701"), aj.Partner("(주)한결모터스", "00702")]
        j2 = aj.Journalizer(aj.load_config(), L2, [], [], [])
        self.assertEqual([p.code for p, _ in j2.identify("(주)한결모터스 (주)한결모터스 타행송금")], ["00702"])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 9, 3), "이영희 타행송금", 0, 1_000_500, party="이영희")])
        self.assertTrue(all(l.partner is None or l.partner.code == "" for l in j.entries[0].lines
                            if l.account != "보통예금"))

    def test_fin_partner_list(self):
        fins = aj.load_fin_partners(S / "금융거래처.xlsx")
        self.assertEqual([(f.code, f.bizno) for f in fins], [("99000", "111222333"), ("99001", "000000000000")])
        self.assertEqual(aj.bank_account_number(S / "통장_202609.xlsx"), "000000000000")
        out = Path(tempfile.mkdtemp()) / "fin.xlsx"
        j = aj.run(S / "통장_202609.xlsx", fin_partners=S / "금융거래처.xlsx", out=out)
        bank_lines = [l for e in j.entries for l in e.lines if l.account == "보통예금"]
        self.assertTrue(all(l.partner and (l.partner.code, l.partner.name) == ("99001", "테스트은행(원화)")
                            for l in bank_lines))

    def test_unknown_deposit_is_receivable(self):
        j = aj.Journalizer(aj.load_config(), [], [], [], [])
        j.journalize_bank([aj.Txn(1, aj.date(2026, 9, 1), "타행이체 모르는사람", 50_000, 0)])
        e = j.entries[0]
        self.assertEqual({(l.side, l.account): l.amount for l in e.lines},
                         {("차변", "보통예금"): 50_000, ("대변", "외상매출금"): 50_000})
        self.assertTrue(e.review)


class WithholdingLedger(unittest.TestCase):
    """예수금 원장의 직전월 예수금만 예수금으로, 나머지는 보험료(821)."""

    def journal(self, rows, txn):
        j = aj.Journalizer(aj.load_config(), [], [], [], [], rows)
        j.journalize_bank([txn])
        return j.entries[0]

    def test_health_feb_payment_uses_jan_withheld(self):
        rows = [aj.WithheldRow(aj.date(2026, 1, 25), "2026-01", "health", 500_000, "국민건강보험공단")]
        e = self.journal(rows, aj.Txn(1, aj.date(2026, 2, 10), "CMS 국민건강보험공단", 0, 1_000_000))
        got = [(l.side, aj.load_config()["accounts"].get(l.account), l.account, l.amount) for l in e.lines]
        self.assertEqual(got, [("차변", "254", "예수금", 500_000),
                               ("차변", "821", "보험료", 500_000),
                               ("대변", "103", "보통예금", 1_000_000)])
        self.assertFalse(e.review)

    def codes(self, e):
        acc = aj.load_config()["accounts"]
        return [(l.side, acc.get(l.account), l.amount) for l in e.lines]

    def test_each_insurance_account(self):
        rows = [aj.WithheldRow(None, "2026-01", t, 100_000) for t in ("health", "pension", "employment")]
        cases = {
            "국민건강보험공단": [("차변", "254", 100_000), ("차변", "821", 150_000), ("대변", "103", 250_000)],
            "국민연금공단": [("차변", "254", 100_000), ("차변", "817", 150_000), ("대변", "103", 250_000)],
            "고용보험": [("차변", "254", 100_000), ("차변", "821", 150_000), ("대변", "103", 250_000)],
            "산재보험": [("차변", "821", 250_000), ("대변", "103", 250_000)],
        }
        for desc, want in cases.items():
            with self.subTest(desc):
                e = self.journal(rows, aj.Txn(1, aj.date(2026, 2, 10), desc, 0, 250_000))
                self.assertEqual(self.codes(e), want)

    def test_comwel_only_is_flagged(self):
        rows = [aj.WithheldRow(None, "2026-01", "employment", 100_000)]
        e = self.journal(rows, aj.Txn(1, aj.date(2026, 2, 10), "근로복지공단", 0, 250_000))
        self.assertTrue(e.review)

    def test_withheld_not_deducted_twice(self):
        rows = [aj.WithheldRow(None, "2026-01", "employment", 100_000)]
        j = aj.Journalizer(aj.load_config(), [], [], [], [], rows)
        j.journalize_bank([aj.Txn(1, aj.date(2026, 2, 10), "고용보험", 0, 250_000),
                           aj.Txn(2, aj.date(2026, 2, 11), "고용보험", 0, 30_000)])
        self.assertEqual(self.codes(j.entries[1]), [("차변", "821", 30_000), ("대변", "103", 30_000)])

    def test_ledger_file_parsing(self):
        rows = aj.load_withholding(S / "예수금원장.xlsx", 2026)
        by = {}
        for r in rows:
            by[(r.typ, r.month)] = by.get((r.typ, r.month), 0) + r.amount
        self.assertEqual(by[("health", "2026-08")], 220_210)      # 건강 + 장기요양, 이월·월계·차변 제외
        self.assertEqual(by[("pension", "2026-08")], 247_500)
        self.assertEqual(by[("income_tax", "2026-08")], 115_980)
        self.assertEqual(by[("local_tax", "2026-08")], 11_590)

    def test_withheld_bigger_than_payment_is_flagged(self):
        rows = [aj.WithheldRow(None, "2026-01", "health", 1_200_000)]
        e = self.journal(rows, aj.Txn(1, aj.date(2026, 2, 10), "국민건강보험", 0, 1_000_000))
        self.assertEqual(e.lines[0].amount, 1_000_000)
        self.assertTrue(e.review)


class FeeOnlyFixedAmount(unittest.TestCase):
    """500원이 아닌 차액은 지급수수료로 처리하지 않는다."""

    def make(self, txn_amount, desc="급여 홍길동"):
        cfg = aj.load_config()
        pay = [aj.PayRow("2026-09", "홍길동", 3_000_000, 74_350, 7_430, 135_000, 120_120, 27_000, 0, 2_636_100)]
        biz = [aj.BizRow("2026-09", "이강사", 1_000_000, 30_000, 3_000, 967_000)]
        j = aj.Journalizer(cfg, [], [], pay, biz)
        t = aj.Txn(1, aj.date(2026, 9, 25), desc, 0, txn_amount)
        j.journalize_bank([t])
        return j.entries[0]

    def test_payroll_other_difference_not_fee(self):
        e = self.make(2_637_100)  # 1,000원 차이
        ln = {(l.side, l.account): l.amount for l in e.lines}
        self.assertNotIn(("차변", "지급수수료"), ln)
        self.assertEqual(ln[("차변", "미지급비용")], 2_636_100)
        self.assertEqual(ln[("차변", "가지급금")], 1_000)
        self.assertTrue(e.review)

    def test_payroll_paid_less_leaves_payable(self):
        e = self.make(2_600_000)  # 36,100원 덜 나감
        ln = {(l.side, l.account): l.amount for l in e.lines}
        self.assertEqual(ln, {("차변", "미지급비용"): 2_600_000, ("대변", "보통예금"): 2_600_000})
        self.assertIn("남은 금액 36,100", e.basis)   # 분할지급 — 남은 금액은 급여지급대조 시트

    def test_payroll_exact_has_no_fee(self):
        e = self.make(2_636_100)
        self.assertNotIn("지급수수료", [l.account for l in e.lines])

    def test_business_other_difference_not_fee(self):
        e = self.make(968_000, "타행이체 이강사")  # 1,000원 차이 → 사업소득으로 단정하지 않음
        self.assertNotIn("송금수수료", [l.memo for l in e.lines])
        self.assertTrue(e.review)


class SemusarangPayroll(unittest.TestCase):
    """세무사랑 급상여대장(3줄 블록) + 영문 이름 + 분할지급."""

    def test_parse(self):
        rows = aj.load_payroll(S / "세무사랑_급상여대장_202605.xlsx")
        self.assertEqual(len(rows), 1)
        p = rows[0]
        self.assertEqual((p.month, p.name, p.gross, p.net), ("2026-05", "알렉스 TEST ALEX", 4_200_000, 3_700_000))
        self.assertEqual((p.pension, p.health, p.income_tax, p.local_tax, p.employment), (180_000, 170_000, 136_360, 13_640, 0))
        self.assertEqual(p.deductions, 500_000)

    def test_name_match(self):
        self.assertTrue(aj.name_in("알렉스 TEST ALEX", "TESTALEX"))
        self.assertTrue(aj.name_in("알렉스 TEST ALEX", "타행 알렉스"))
        self.assertTrue(aj.name_in("KIM TESTOVICH", "KIMTESTOVI"))   # 통장에서 잘린 이름
        self.assertFalse(aj.name_in("KIM TESTOVICH", "KIM"))

    def test_split_payment_with_prior_period(self):
        pay = aj.load_payroll(S / "세무사랑_급상여대장_202605.xlsx")
        j = aj.Journalizer(aj.load_config(), [], [], pay, [])
        j.prime([aj.Txn(1, aj.date(2026, 5, 2), "TESTALEX", 0, 2_000_000)])      # 작업 시작 전 지급
        j.journalize_bank([aj.Txn(2, aj.date(2026, 5, 25), "TESTALEX", 0, 1_700_000),
                           aj.Txn(3, aj.date(2026, 5, 28), "TESTALEX", 0, 900_000)])
        first, extra = j.entries
        self.assertEqual({(l.side, l.account): l.amount for l in first.lines},
                         {("차변", "미지급비용"): 1_700_000, ("대변", "보통예금"): 1_700_000})
        self.assertFalse(first.review)
        self.assertIn("분할지급 완료", first.basis)
        # 다른 사람에게 보낸 같은 금액은 급여로 붙지 않음
        j2 = aj.Journalizer(aj.load_config(), [], [], aj.load_payroll(S / "세무사랑_급상여대장_202605.xlsx"), [])
        j2.journalize_bank([aj.Txn(9, aj.date(2026, 5, 3), "KIM OTHER", 0, 3_700_500)])
        self.assertNotIn("미지급비용", [l.account for l in j2.entries[0].lines])
        # 급여대장보다 더 나간 돈 → 확인필요
        self.assertTrue(extra.review)
        self.assertEqual(pay[0].remaining, 0)


class BusinessCumulative(unittest.TestCase):
    """사업소득: 송금 건마다 500원 수수료, 사람별 누적 털기, 철자 다른 이름 별칭."""

    def setUp(self):
        cfg = aj.load_config()
        cfg["name_aliases"] = {"TEST ANNA": ["TEST ANNNA"]}
        biz = [aj.BizRow("2026-05", "TEST ANNA", 3_000_000, 90_000, 9_000, 2_901_000),
               aj.BizRow("2026-06", "TEST ANNA", 1_000_000, 30_000, 3_000, 967_000)]
        self.j = aj.Journalizer(cfg, [], [], [], biz)
        self.biz = biz

    def acc(self, e):
        return {(l.side, l.account): l.amount for l in e.lines}

    def test_fee_each_transfer_and_cumulative(self):
        self.j.journalize_bank([
            aj.Txn(1, aj.date(2026, 5, 10), "TEST ANNNA", 0, 2_000_500),   # 별칭으로 인식
            aj.Txn(2, aj.date(2026, 5, 20), "TEST ANNNA", 0, 1_000_500),   # 5월분 초과 → 6월분으로 누적
            aj.Txn(3, aj.date(2026, 6, 20), "TEST ANNNA", 0, 868_000),     # 끝자리 000 → 수수료 없음
        ])
        e1, e2, e3 = self.j.entries
        self.assertEqual(self.acc(e1), {("차변", "미지급비용"): 2_000_000, ("차변", "지급수수료"): 500,
                                        ("대변", "보통예금"): 2_000_500})
        self.assertEqual(self.acc(e2), {("차변", "미지급비용"): 1_000_000, ("차변", "지급수수료"): 500,
                                        ("대변", "보통예금"): 1_000_500})
        self.assertEqual(self.acc(e3), {("차변", "미지급비용"): 868_000, ("대변", "보통예금"): 868_000})
        self.assertFalse(any(e.review for e in self.j.entries))
        self.assertEqual(sum(b.remaining for b in self.biz), 0)

    def test_excess_over_total_flagged(self):
        self.j.journalize_bank([aj.Txn(1, aj.date(2026, 6, 30), "TEST ANNNA", 0, 4_000_500)])
        e = self.j.entries[0]
        self.assertEqual(self.acc(e)[("차변", "미지급비용")], 3_868_000)
        self.assertEqual(self.acc(e)[("차변", "가지급금")], 132_000)
        self.assertTrue(e.review)

    def test_client_config_overlay(self):
        import json, tempfile, os
        fd, path = tempfile.mkstemp(suffix=".json")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump({"name_aliases": {"A B": ["AB2"]}, "accounts": {"보통예금": "999"}}, f)
        cfg = aj.load_config(None, path)
        self.assertEqual(cfg["name_aliases"], {"A B": ["AB2"]})
        self.assertEqual(cfg["accounts"]["보통예금"], "999")
        self.assertEqual(cfg["accounts"]["예수금"], "254")   # 나머지는 기본값 유지


class SemusarangInvoiceList(unittest.TestCase):
    def test_parse_and_code(self):
        sales = aj.load_invoices(S / "세무사랑_매입매출장.xlsx", "매출")
        buys = aj.load_invoices(S / "세무사랑_매입매출장.xlsx", "매입")
        self.assertEqual([(i.partner.name, i.partner.code, i.total) for i in sales], [("샘플모터스", "00777", 11_000_000)])
        self.assertEqual([(i.partner.name, i.partner.code, i.total) for i in buys], [("샘플부품", "00888", 1_100_000)])
        j = aj.Journalizer(aj.load_config(), [], sales + buys, [], [])
        self.assertEqual(j.new_partners, {})                 # 코드가 있으니 신규거래처 아님
        j.journalize_bank([aj.Txn(1, aj.date(2026, 5, 9), "샘플모터스", 11_000_000, 0)])
        self.assertEqual(j.entries[0].lines[1].partner.code, "00777")
        self.assertFalse(j.entries[0].review)


if __name__ == "__main__":
    unittest.main()
