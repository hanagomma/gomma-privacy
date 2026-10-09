#!/usr/bin/env python3
"""통장 거래내역 자동 분개.

거래처 통장 내역을 받아 세금계산서(매출/매입), 급여대장, 사업소득 지급내역,
예수금(원천세·4대보험) 납부를 반영해 일반전표를 만들고,
개인 이름으로 들어온 입출금은 거래처 원장의 대표자명으로 회사를 찾아 거래처로 지정한다.

사용 예:
  python auto_journal.py --bank 통장.xlsx --ledger 거래처원장.xlsx \
      --sales 매출세금계산서.xlsx --purchase 매입세금계산서.xlsx \
      --payroll 급여대장.xlsx --business 사업소득.xlsx --out 분개결과.xlsx
"""
from __future__ import annotations

import argparse
import itertools
import json
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------- 공통 유틸

CORP_TOKENS = ["농업회사법인", "영농조합법인", "주식회사", "유한회사", "합자회사", "합명회사",
               "사단법인", "재단법인", "(주)", "㈜", "(유)", "(합)", "(사)", "(재)", "주)", "(주"]


def norm(s) -> str:
    """상호 비교용 정규화: 법인 표기·공백·기호 제거, 영문 대문자화."""
    s = "" if s is None or (isinstance(s, float) and pd.isna(s)) else str(s)
    for t in CORP_TOKENS:
        s = s.replace(t, "")
    return re.sub(r"[\s()\[\]{}.\-_,·/&'\"*]", "", s).upper()


def digits(s) -> str:
    return re.sub(r"\D", "", "" if s is None else str(s))


def to_amount(v) -> int:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    s = re.sub(r"[,\s원₩]", "", str(v))
    if s in ("", "-", "nan", "None"):
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def to_date(v, default_year: int | None = None) -> date | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, pd.Timestamp):
        return v.date()
    s = str(v).strip()
    if re.fullmatch(r"\d{8}(\.0)?", s):
        return datetime.strptime(s[:8], "%Y%m%d").date()
    s = re.sub(r"[./년월]", "-", s).replace("일", "")
    m = re.fullmatch(r"(\d{1,2})-(\d{1,2})-?", s.strip())
    if m and default_year:  # 더존 원장처럼 연도 없이 '01-31' 로 나오는 경우
        return date(default_year, int(m.group(1)), int(m.group(2)))
    try:
        return pd.to_datetime(s.split()[0] if " " in s else s).date()
    except (ValueError, TypeError):
        return None


def to_month(v) -> str | None:
    """'2026-09', '2026.09', 202609, 날짜 → 'YYYY-MM'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    m = re.match(r"(\d{4})\D?(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    d = to_date(v)
    return d.strftime("%Y-%m") if d else None


def prev_month(ym: str) -> str:
    y, m = map(int, ym.split("-"))
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


def hangul_tokens(s: str) -> list[str]:
    return re.findall(r"[가-힣]{2,}", s or "")


def party_tokens(text: str) -> set[str]:
    """통장 상대방 표시를 비교용 단어로: 띄어쓰기·괄호 단위 + 괄호 안/밖 묶음 + 전체."""
    text = text or ""
    groups = re.findall(r"\(([^()]*)\)?", text) + [re.sub(r"\(.*?(\)|$)", " ", text)]
    words = [w for w in re.split(r"[\s()]+", text) if w]
    toks = {norm(w) for w in words}
    toks |= {norm("".join(words[i:i + k])) for k in (2, 3) for i in range(len(words) - k + 1)}  # '별 관세사무'
    toks |= {norm(g) for g in groups} | {norm(text)}
    return {t for t in toks if len(t) >= 2}   # 2글자는 상호와 완전히 같을 때만 일치 (company_match)


def company_match(name: str, tokens: set[str]) -> bool:
    """상호가 통장 단어와 같거나, 통장에서 잘린 앞부분(3자 이상)일 때만 일치.
    '나모터스' 가 '가나모터스' 안에 들어 있는 것처럼 일부만 겹치는 경우는 일치로 보지 않는다."""
    n = norm(name)
    core = norm(re.sub(r"\(.*?\)", "", name))
    base = re.sub(r"(본점|본사|지점)$", "", core)          # '(주)누리 본점' → '누리'
    for t in tokens:
        for k in {n, core, base}:
            if len(k) >= 2 and (k == t or (len(t) >= 3 and k.startswith(t) and len(t) >= len(k) * 0.6)):
                return True
    return False


def name_in(name: str, desc: str) -> bool:
    """사람 이름이 통장 표시에 있는지. '홍길동 HONG GILDONG' 처럼 한글·영문이 섞인 이름은
    한글 부분(토큰 일치)과 영문 부분(공백 없이, 통장에서 잘린 경우 포함)을 각각 비교한다."""
    d = norm(desc)
    if not name or not d:
        return False
    toks = {norm(t) for t in hangul_tokens(desc)}
    full = norm(name)
    if full == d or full in toks:
        return True
    kor = "".join(re.findall(r"[가-힣]+", name))
    eng = "".join(re.findall(r"[A-Za-z]+", name)).upper()
    if eng and kor and len(kor) >= 2 and kor in toks:
        return True
    if len(eng) >= 5:
        de = re.sub(r"[^A-Z]", "", d)
        return eng in de or (len(de) >= 8 and eng.startswith(de))
    return False


# ---------------------------------------------------------------- 표 읽기

def _read_raw(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "cp949", "euc-kr"):
            try:
                return pd.read_csv(path, header=None, dtype=object, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"CSV 인코딩을 알 수 없습니다: {path}")
    return pd.read_excel(path, header=None, dtype=object)


def read_table(path, aliases: dict[str, list[str]], required: list[str]) -> pd.DataFrame:
    """머리글 행을 자동으로 찾아 표준 열 이름(aliases의 키)으로 바꿔 돌려준다.

    중복 머리글(홈택스의 '상호' 두 번 등)은 '상호', '상호.1' 로 구분된다.
    """
    path = Path(path)
    raw = _read_raw(path)
    best_row, best_hits = None, 0
    for i in range(min(len(raw), 30)):
        cells = {str(c).strip().replace(" ", "") for c in raw.iloc[i].tolist() if pd.notna(c)}
        hits = sum(any(a.replace(" ", "") in cells for a in al) for al in aliases.values())
        if hits > best_hits:
            best_row, best_hits = i, hits
    if best_row is None:
        raise ValueError(f"{path.name}: 머리글 행을 찾지 못했습니다.")

    header, seen = [], defaultdict(int)
    for c in raw.iloc[best_row].tolist():
        name = str(c).strip().replace(" ", "") if pd.notna(c) else ""
        header.append(name if seen[name] == 0 else f"{name}.{seen[name]}")
        seen[name] += 1
    df = raw.iloc[best_row + 1:].copy()
    df.columns = header
    df = df.dropna(how="all")

    out = pd.DataFrame(index=df.index)
    for key, al in aliases.items():
        for a in al:
            a = a.replace(" ", "")
            if a in df.columns:
                out[key] = df[a]
                break
    missing = [k for k in required if k not in out.columns]
    if missing:
        raise ValueError(f"{path.name}: 필수 열 {missing} 을(를) 찾지 못했습니다. 머리글: {header}")
    out.attrs["raw"] = df
    return out


# ---------------------------------------------------------------- 데이터 모델

@dataclass
class Partner:
    name: str
    code: str = ""
    bizno: str = ""
    ceo: str = ""
    account: str = ""          # 거래처 원장의 기본 계정(있으면 매입 계정으로 사용)
    aliases: tuple = ()        # 통장에 찍히는 다른 이름(예: 케이티 → KT)
    in_ledger: bool = True

    @property
    def key(self) -> str:
        return self.bizno or norm(self.name)


@dataclass
class Txn:
    row: int
    date: date
    desc: str
    deposit: int
    withdrawal: int
    balance: int | None = None
    raw_text: str = ""         # 통장 행의 모든 칸 텍스트(제외 키워드 판정용)
    party: str = ""            # 상대방 이름 칸만(적요·의뢰인/수취인 등) — 거래 구분 단어 제외

    @property
    def amount(self) -> int:
        return self.deposit or self.withdrawal

    @property
    def side(self) -> str:
        return "입금" if self.deposit else "출금"


@dataclass
class Invoice:
    idx: int
    kind: str                  # 매출 / 매입
    date: date | None
    partner: Partner
    supply: int
    vat: int
    total: int
    item: str = ""
    matched_rows: list = field(default_factory=list)


@dataclass
class PayRow:
    month: str | None
    name: str
    gross: int
    income_tax: int
    local_tax: int
    pension: int
    health: int                # 건강보험 + 장기요양
    employment: int
    other: int
    net: int
    paid: bool = False
    paid_amt: int = 0          # 통장으로 이미 지급한 금액(분할지급 누적)

    @property
    def deductions(self) -> int:
        return self.income_tax + self.local_tax + self.pension + self.health + self.employment + self.other

    @property
    def remaining(self) -> int:
        return self.net - self.paid_amt


@dataclass
class BizRow:
    month: str | None
    name: str
    gross: int
    income_tax: int
    local_tax: int
    net: int
    paid: bool = False
    paid_amt: int = 0

    @property
    def remaining(self) -> int:
        return self.net - self.paid_amt


@dataclass
class WithheldRow:
    """예수금 원장 한 줄(대변 = 예수 발생)."""
    date: date | None
    month: str | None
    typ: str                   # health / pension / employment / income_tax / local_tax
    amount: int
    partner: str = ""
    memo: str = ""


@dataclass
class Line:
    side: str                  # 차변 / 대변
    account: str
    amount: int
    partner: Partner | None = None
    memo: str = ""


@dataclass
class Entry:
    date: date
    lines: list[Line]
    basis: str
    review: str = ""           # 비어 있지 않으면 확인 필요 사유
    source: str = ""           # 통장 n행 등

    def balanced(self) -> bool:
        dr = sum(l.amount for l in self.lines if l.side == "차변")
        cr = sum(l.amount for l in self.lines if l.side == "대변")
        return dr == cr


# ---------------------------------------------------------------- 로더

BANK_ALIASES = {
    "date": ["거래일자", "거래일", "거래일시", "거래일자시간", "일자", "날짜", "거래날짜", "년월일"],
    "deposit": ["입금액", "입금", "맡기신금액", "입금금액", "받은금액", "입금(원)"],
    "withdrawal": ["출금액", "출금", "찾으신금액", "출금금액", "지급금액", "출금(원)"],
    "balance": ["잔액", "거래후잔액", "잔액(원)"],
    "amount": ["거래금액", "금액"],
    "io": ["입출금구분", "입출구분", "구분", "거래구분"],
}
BANK_DESC_COLS = ["적요", "내용", "기재내용", "거래내용", "받는분/보낸분", "보낸분/받는분", "상대방",
                  "의뢰인/수취인", "입금자명", "수취인", "거래기록사항", "메모", "비고", "통장표시내용",
                  "추가메모", "구분", "거래구분", "거래특이사항"]


# 거래 종류를 나타내는 말 — 상대방 이름 비교에서 제외
GENERIC_WORDS = {"타행이체", "타행송금", "당행이체", "당행송금", "인터넷", "인터넷뱅킹", "모바일", "자동이체", "대체",
                 "이체", "송금", "입금", "출금", "체크카드", "공과금", "지로요금", "통신요금", "예금이자", "결산",
                 "급여", "CMS", "펌뱅킹", "폰뱅킹", "오픈뱅킹", "현금", "ATM", "타행", "당행", "인증서"}
def bank_account_number(path) -> str:
    """통장 파일 위쪽 '계좌번호 : 123-456789-01234' 에서 계좌번호(숫자만)."""
    raw = _read_raw(Path(path))
    for r in raw.head(15).itertuples(index=False):
        for v in r:
            m = re.search(r"계좌번호\s*[:：]\s*([\d-]{8,})", str(v))
            if m:
                return digits(m.group(1))
    return ""


FIN_ALIASES = {"code": ["코드"], "name": ["금융기관명", "거래처명"], "account": ["계좌번호"]}


def load_fin_partners(path) -> list[Partner]:
    """세무사랑 '금융기관거래처 등록 LIST': 코드·금융기관명·계좌번호 (계좌번호는 bizno 칸에 숫자만)."""
    df = read_table(path, FIN_ALIASES, ["code", "name"])
    out = []
    for _, r in df.iterrows():
        code = "" if pd.isna(r.get("code")) else str(r.get("code")).strip().removesuffix(".0")
        if not re.fullmatch(r"\d+", code):
            continue
        acct = r.get("account")
        acct = "" if acct is None or pd.isna(acct) else (str(int(acct)) if isinstance(acct, float) else str(acct))
        out.append(Partner(name=str(r.get("name")).strip(), code=code, bizno=digits(acct)))
    return out


BANK_PARTY_COLS = ["적요", "내용", "기재내용", "받는분/보낸분", "보낸분/받는분", "상대방", "의뢰인/수취인",
                   "입금자명", "수취인", "통장표시내용"]


def load_bank(path, start: date | None = None, end: date | None = None) -> list[Txn]:
    df = read_table(path, BANK_ALIASES, ["date"])
    raw = df.attrs["raw"]
    desc_cols = [c for c in raw.columns if c in BANK_DESC_COLS]
    txns = []
    for n, (i, r) in enumerate(df.iterrows(), start=1):
        d = to_date(r.get("date"))
        dep, wd = to_amount(r.get("deposit")), to_amount(r.get("withdrawal"))
        if not dep and not wd and "amount" in df:  # 금액 한 열 + 입금/출금 구분 열 형식
            amt, io = to_amount(r.get("amount")), str(r.get("io") or "")
            dep, wd = (amt, 0) if "입" in io else (0, amt) if ("출" in io or "지급" in io) else \
                ((amt, 0) if amt > 0 else (0, -amt))
        if d is None or (dep == 0 and wd == 0):
            continue
        if (start and d < start) or (end and d > end):
            continue
        desc = " ".join(str(raw.at[i, c]).strip() for c in desc_cols
                        if pd.notna(raw.at[i, c]) and str(raw.at[i, c]).strip())
        raw_text = " ".join(str(v).strip() for v in raw.loc[i].tolist() if pd.notna(v))
        party = " ".join(str(raw.at[i, c]).strip() for c in raw.columns if c in BANK_PARTY_COLS
                         and pd.notna(raw.at[i, c]) and str(raw.at[i, c]).strip())
        txns.append(Txn(n, d, desc, dep, wd, to_amount(r.get("balance")) if "balance" in df else None,
                        raw_text, party))
    return sorted(txns, key=lambda t: (t.date, t.row))


LEDGER_ALIASES = {
    "code": ["거래처코드", "코드", "거래처번호"],
    "name": ["거래처명", "상호", "상호명", "회사명", "거래처", "법인명"],
    "bizno": ["사업자번호", "사업자등록번호", "등록번호", "사업자/주민번호", "사업자(주민등록)번호"],
    "ceo": ["대표자", "대표자명", "대표", "대표자성명", "성명"],
    "account": ["계정과목", "기본계정", "계정"],
    "alias": ["통장표시명", "별칭", "약칭", "통장이름"],
}


def load_ledger(path) -> list[Partner]:
    df = read_table(path, LEDGER_ALIASES, ["name"])
    out = []
    for _, r in df.iterrows():
        name = str(r.get("name") or "").strip()
        if not name or name == "nan":
            continue
        clean = lambda k: "" if pd.isna(r.get(k)) else str(r.get(k)).strip()
        if "code" in df and not re.fullmatch(r"\d+(\.0)?", clean("code")):
            continue   # 페이지마다 반복되는 머리글 줄·담당자 줄 (세무사랑 거래처 등록 LIST)
        out.append(Partner(name=name, code=clean("code").removesuffix(".0"),
                           bizno=digits(clean("bizno")), ceo=clean("ceo"), account=clean("account"),
                           aliases=tuple(a.strip() for a in re.split(r"[,;/]", clean("alias")) if a.strip())))
    return out


INVOICE_ALIASES = {
    "date": ["작성일자", "작성일", "일자", "발행일자", "발급일자"],
    "total": ["합계금액", "합계", "총금액"],
    "supply": ["공급가액"],
    "vat": ["세액", "부가세"],
    "item": ["품목명", "품목"],
    # 홈택스 목록: 공급자 정보가 먼저, 공급받는자 정보가 뒤(.1)에 온다.
    "s_bizno": ["공급자사업자등록번호"],
    "s_name": ["상호"],
    "s_ceo": ["대표자명"],
    "b_bizno": ["공급받는자사업자등록번호"],
    "b_name": ["상호.1"],
    "b_ceo": ["대표자명.1"],
    # 간단 양식(거래처 열이 하나뿐인 경우)
    "p_name": ["거래처", "거래처명", "상대방상호"],
    "p_bizno": ["사업자번호", "거래처사업자번호"],
    "p_ceo": ["대표자", "거래처대표자"],
    "p_code": ["거래처코드", "코드"],           # 세무사랑 매입매출장: 첫 '코드' 열이 거래처코드
    "io": ["구분"],                            # 세무사랑 매입매출장: 매출/매입
}


def load_invoices(path, kind: str, start_idx: int = 0) -> list[Invoice]:
    df = read_table(path, INVOICE_ALIASES, ["date"])
    if "total" not in df and "supply" not in df:
        raise ValueError(f"{Path(path).name}: 합계금액/공급가액 열이 없습니다.")
    # 매출 → 상대방은 공급받는자, 매입 → 상대방은 공급자
    # 간단 양식(상호 열이 하나)이면 그 '상호'가 곧 상대 거래처
    hometax = "b_name" in df
    pre = ("b_" if kind == "매출" else "s_") if hometax else "s_"
    out = []
    for n, (_, r) in enumerate(df.iterrows()):
        g = lambda k: "" if k not in df or pd.isna(r.get(k)) else str(r.get(k)).strip()
        name = g(pre + "name") if hometax else (g("p_name") or g("s_name"))
        if not name:
            continue
        if g("io") in ("매출", "매입") and g("io") != kind:
            continue  # 매입매출장에 매출·매입이 섞여 있으면 해당 구분만
        supply, vat = to_amount(r.get("supply")), to_amount(r.get("vat"))
        total = to_amount(r.get("total")) or supply + vat
        if not supply:
            supply = total - vat
        p = Partner(name=name, code=g("p_code").removesuffix(".0") if not hometax else "",
                    bizno=digits(g(pre + "bizno") or g("p_bizno")),
                    ceo=g(pre + "ceo") or g("p_ceo"), in_ledger=False)
        out.append(Invoice(start_idx + n, kind, to_date(r.get("date")), p, supply, vat, total, g("item")))
    return out


PAY_ALIASES = {
    "month": ["귀속월", "귀속연월", "지급월", "급여월", "지급일", "지급일자", "귀속년월"],
    "name": ["성명", "이름", "사원명", "직원명"],
    "gross": ["지급총액", "지급합계", "총지급액", "급여총액", "과세총액", "지급액계", "급여합계", "총급여"],
    "income_tax": ["소득세", "근로소득세"],
    "local_tax": ["지방소득세", "주민세"],
    "pension": ["국민연금"],
    "health": ["건강보험"],
    "ltc": ["장기요양보험", "장기요양", "요양보험"],
    "employment": ["고용보험"],
    "ded_total": ["공제총액", "공제합계", "공제액계", "공제계"],
    "net": ["차인지급액", "실지급액", "실수령액", "차감지급액", "실지급"],
}


def load_payroll_semusarang(path) -> list[PayRow] | None:
    """세무사랑 「급상여대장」: 머리글 3줄, 사람마다 3줄 블록. 해당 양식이 아니면 None."""
    raw = _read_raw(Path(path)).fillna("")
    cells = raw.values.tolist()
    h = next((i for i, r in enumerate(cells) if any(str(v).replace(" ", "") == "사원번호" for v in r)), None)
    if h is None or h + 3 > len(cells):
        return None
    labels = {}
    for k in range(3):
        for c, v in enumerate(cells[h + k]):
            lab = str(v).replace(" ", "").replace("\n", "")
            if lab:
                labels.setdefault(lab, (k, c))
    if "차인지급액" not in labels or "성명" not in labels:
        return None
    month = None
    for r in cells[:h]:
        m = re.search(r"(\d{4})년\s*(\d{1,2})월분", " ".join(map(str, r)))
        if m:
            month = f"{m.group(1)}-{int(m.group(2)):02d}"
            break
    kc, cc = labels["사원번호"]
    out, i = [], h + 3
    while i + 2 < len(cells) + 2 and i < len(cells):
        key = str(cells[i][cc]).strip()
        if not re.fullmatch(r"\d+(\.0)?", key):     # '인원 : 1' 합계 블록 등
            i += 1
            continue
        block = cells[i:i + 3]
        g = lambda lab: to_amount(block[labels[lab][0]][labels[lab][1]]) \
            if lab in labels and labels[lab][0] < len(block) else 0
        name = str(block[labels["성명"][0]][labels["성명"][1]]).strip()
        gross, net = g("지급합계"), g("차인지급액")
        income = g("소득세") + g("연말정산소득세")
        local = g("지방소득세") + g("연말정산지방소득세")
        pension = g("국민연금")
        health = g("건강보험") + g("장기요양보험") + g("건강보험료정산") + g("장기요양보험료정산")
        emp = g("고용보험")
        other = (gross - net) - (income + local + pension + health + emp)
        out.append(PayRow(month, name, gross, income, local, pension, health, emp, other, net))
        i += 3
    return out


def load_payroll(path, default_month: str | None = None) -> list[PayRow]:
    rows = load_payroll_semusarang(path)
    if rows is not None:
        return rows
    df = read_table(path, PAY_ALIASES, ["name"])
    out = []
    for _, r in df.iterrows():
        name = str(r.get("name") or "").strip()
        if not name or name == "nan" or name in ("합계", "계", "총계", "소계"):
            continue
        a = {k: to_amount(r.get(k)) for k in PAY_ALIASES if k not in ("month", "name")}
        known = a["income_tax"] + a["local_tax"] + a["pension"] + a["health"] + a["ltc"] + a["employment"]
        gross, net = a["gross"], a["net"]
        ded = a["ded_total"] or known
        if not gross and net:
            gross = net + ded
        if not net:
            net = gross - ded
        other = max(0, (gross - net) - known)
        out.append(PayRow(to_month(r.get("month")) or default_month, name, gross, a["income_tax"],
                          a["local_tax"], a["pension"], a["health"] + a["ltc"], a["employment"], other, net))
    return out


BIZ_ALIASES = {
    "month": ["귀속월", "귀속년월", "지급월", "지급일", "지급일자", "귀속연월", "지급년월일"],
    "name": ["성명", "이름", "소득자", "소득자명", "상호"],
    "gross": ["지급액", "지급총액", "총지급액", "지급금액"],
    "income_tax": ["소득세"],
    "local_tax": ["지방소득세"],
    "net": ["차인지급액", "실지급액", "실수령액", "차감지급액"],
}


def load_business(path, default_month: str | None = None) -> list[BizRow]:
    df = read_table(path, BIZ_ALIASES, ["name"])
    out = []
    for _, r in df.iterrows():
        name = str(r.get("name") or "").strip()
        if not name or name == "nan" or name in ("합계", "계", "총계"):
            continue
        gross, it, lt, net = (to_amount(r.get(k)) for k in ("gross", "income_tax", "local_tax", "net"))
        if not gross and net:  # 실지급액만 있는 경우 역산 (3.3%)
            gross = round(net / 0.967)
        if not it and "income_tax" not in df:
            it = gross * 3 // 100 // 10 * 10          # 소득세 10원 미만 절사
        if not lt and "local_tax" not in df:
            lt = it // 10 // 10 * 10
        if not net:
            net = gross - it - lt
        out.append(BizRow(to_month(r.get("month")) or default_month, name, gross, it, lt, net))
    return out


WH_ALIASES = {
    "date": ["일자", "날짜", "전표일자", "거래일자", "월일", "월/일"],
    "month": ["귀속월", "귀속연월", "귀속년월"],
    "account": ["계정과목", "계정", "계정명"],
    "partner": ["거래처", "거래처명", "거래처명칭"],
    "memo": ["적요", "적요명", "내용"],
    "type": ["구분", "항목", "공제항목"],
    "debit": ["차변", "차변금액"],
    "credit": ["대변", "대변금액"],
    "amount": ["금액", "예수금액", "예수액"],
}
# 앞에서부터 판정 — '지방소득세'가 '소득세'보다 먼저
WH_TYPES = [
    ("local_tax", ["지방소득세", "지방세", "위택스", "구청", "시청", "군청"]),
    ("health", ["건강", "장기요양", "요양보험", "건보"]),
    ("pension", ["연금"]),
    ("employment", ["고용", "산재", "근로복지"]),
    ("income_tax", ["소득세", "원천세", "세무서", "국세"]),
]
WH_LABEL = {"health": "건강보험", "pension": "국민연금", "employment": "고용·산재보험",
            "income_tax": "원천세(소득세)", "local_tax": "지방소득세(원천)"}
SKIP_MEMO = ("전기이월", "전월이월", "이월", "월계", "누계", "합계", "소계")


def _withholding_from_pdf(path, default_year: int | None = None) -> list[WithheldRow]:
    """세무사랑 '거래처 원장' PDF(예수금). 글자 위치로 차변/대변 열을 가려 대변(예수 발생)만 읽는다."""
    from pypdf import PdfReader
    out = []
    for page in PdfReader(str(path)).pages:
        items = []
        page.extract_text(visitor_text=lambda t, cm, tm, fd, fs: items.append((tm[4], tm[5], t.strip()))
                          if t.strip() else None)
        head = {t: x for x, y, t in items}
        x_dr = next((x for t, x in head.items() if t.replace(" ", "") == "차변"), None)
        x_cr = next((x for t, x in head.items() if t.replace(" ", "") == "대변"), None)
        if x_dr is None or x_cr is None:
            continue
        mid = (x_dr + x_cr) / 2 + 20                     # 금액은 오른쪽 정렬 → 대변 열 경계
        year = default_year
        partner = ""
        for x, y, t in items:
            m = re.search(r"(\d{4})\.\d{2}\.\d{2}\s*~", t)
            if m:
                year = int(m.group(1))
            m = re.search(r"거래처명\s*:\s*(?:\[\d+\])?\s*(.+)", t)
            if m:
                partner = m.group(1).strip()
        rows = defaultdict(list)
        for x, y, t in items:
            rows[round(y)].append((x, t))
        for y in sorted(rows, reverse=True):
            cells = sorted(rows[y])
            date_txt = next((t for x, t in cells if re.fullmatch(r"\d{2}-\d{2}", t)), None)
            if not date_txt:
                continue
            memo = " ".join(t for x, t in cells if 180 <= x < x_dr - 30)
            cr = [t for x, t in cells if x >= mid and re.fullmatch(r"-?[\d,]+", t)]
            if not cr:
                continue
            d = to_date(date_txt, year)
            text = f"{partner} {memo}"
            typ = next((tp for tp, kws in WH_TYPES if any(k in text for k in kws)), None)
            if typ:
                out.append(WithheldRow(d, d.strftime("%Y-%m"), typ, to_amount(cr[0]), partner, memo))
    return out


def load_withholding(path, default_year: int | None = None) -> list[WithheldRow]:
    """예수금 거래처원장/계정별원장. 대변(예수 발생)만 귀속월별로 모은다. 차변은 지난 납부라 무시."""
    if Path(path).suffix.lower() == ".pdf":
        return _withholding_from_pdf(path, default_year)
    df = read_table(path, WH_ALIASES, [])
    if "credit" not in df and "amount" not in df:
        raise ValueError(f"{Path(path).name}: 대변(또는 금액) 열이 없습니다.")
    out, cur_date = [], None
    for _, r in df.iterrows():
        g = lambda k: "" if k not in df or pd.isna(r.get(k)) else str(r.get(k)).strip()
        if g("account") and "예수금" not in g("account"):
            continue
        text = " ".join((g("type"), g("partner"), g("memo")))
        if any(t in text.replace(" ", "") for t in SKIP_MEMO):
            continue
        d = to_date(r.get("date"), default_year) if g("date") else None
        cur_date = d or cur_date  # 원장은 같은 날짜를 첫 줄에만 적는 경우가 많다
        amt = to_amount(r.get("credit")) if "credit" in df else to_amount(r.get("amount"))
        if amt <= 0:
            continue
        typ = next((t for t, kws in WH_TYPES if any(k in text for k in kws)), None)
        if typ is None:
            continue
        month = to_month(g("month")) if g("month") else (cur_date.strftime("%Y-%m") if cur_date else None)
        out.append(WithheldRow(cur_date, month, typ, amt, g("partner"), g("memo")))
    return out


@dataclass
class TaxPayment:
    """국세청 납부내역증명 한 줄."""
    year: str
    item: str                  # 세목 (법인세, 근로소득세(갑) ...)
    date: date
    amount: int
    matched_row: int | None = None
    note: str = ""


def load_tax_payments(path) -> list[TaxPayment]:
    """홈택스 '납부내역증명' PDF(또는 같은 열의 엑셀: 귀속연도·세목·납부일·합계)."""
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        import subprocess
        try:
            text = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, text=True,
                                  check=True).stdout
        except (OSError, subprocess.CalledProcessError):
            from pypdf import PdfReader
            text = "\n".join(pg.extract_text() or "" for pg in PdfReader(str(path)).pages)
        out = []
        for m in re.finditer(r"^\s*(\d{4})\s+(\S[^\n]*?)\s+(\d{4}-\d{2}-\d{2})\s+([\d,]+)\s*$", text, re.M):
            out.append(TaxPayment(m.group(1), m.group(2).strip(), to_date(m.group(3)), to_amount(m.group(4))))
        return out
    df = read_table(path, {"year": ["귀속연도", "귀속"], "item": ["세목"], "date": ["납부일", "납부일자"],
                           "amount": ["합계", "납부금액", "금액"], "note": ["비고", "과세대상"]},
                    ["item", "date", "amount"])
    return [TaxPayment(str(r.get("year") or ""), str(r["item"]).strip(), to_date(r["date"]), to_amount(r["amount"]),
                       note="" if pd.isna(r.get("note")) else str(r.get("note")))
            for _, r in df.iterrows() if to_amount(r["amount"])]


# ---------------------------------------------------------------- 분개 엔진

class Journalizer:
    def __init__(self, config: dict, ledger: list[Partner], invoices: list[Invoice],
                 payroll: list[PayRow], business: list[BizRow], withheld: list[WithheldRow] | None = None):
        self.cfg = config
        self.withheld = withheld or []
        self.tax_payments: list[TaxPayment] = []
        self._wh_used: dict = defaultdict(int)
        self.ap_paid: dict = defaultdict(int)     # 매입 거래처별 통장 지급 누계(수수료 제외)  # (종류, 귀속월) → 이미 납부로 차감한 예수금
        # 급여대장과 사업소득 내역 양쪽에 있는 사람 — 별도 확인 대상(노란색)
        self.dual_names = sorted({p.name for p in payroll} & {b.name for b in business}, key=norm)
        self.ledger = ledger
        self.invoices = invoices
        self.payroll = payroll
        self.business = business
        self.entries: list[Entry] = []
        self.match_log: list[dict] = []
        self.new_partners: dict[str, Partner] = {}
        self._link_invoice_partners()

    # ----- 거래처 -----
    def _ledger_lookup(self, p: Partner) -> Partner | None:
        if p.code:
            for lp in self.ledger:
                if lp.code and lp.code == p.code:
                    return lp
        for lp in self.ledger:
            if p.bizno and lp.bizno and p.bizno == lp.bizno:
                return lp
        for lp in self.ledger:
            if norm(lp.name) and norm(lp.name) == norm(p.name):
                return lp
        return None

    def _link_invoice_partners(self):
        """세금계산서 상대방을 거래처 원장과 연결. 원장에 없으면 신규거래처로 모은다."""
        for inv in self.invoices:
            lp = self._ledger_lookup(inv.partner)
            if lp:
                if not lp.ceo and inv.partner.ceo:
                    lp.ceo = inv.partner.ceo
                inv.partner = lp
            elif inv.partner.code:          # 세무사랑 매입매출장에 거래처코드가 있으면 이미 등록된 거래처
                inv.partner.in_ledger = True
                same = next((lp for lp in self.ledger if lp.code == inv.partner.code), None)
                if same:
                    inv.partner = same
                else:
                    self.ledger.append(inv.partner)
            else:
                key = inv.partner.key
                inv.partner = self.new_partners.setdefault(key, inv.partner)

    def _all_partners(self) -> list[Partner]:
        return self.ledger + list(self.new_partners.values())

    def identify(self, desc: str) -> list[tuple[Partner, str]]:
        """통장 적요로 거래처 후보를 찾는다: (거래처, 방법)."""
        d = norm(desc)
        ptoks = party_tokens(desc)
        tokens = set(hangul_tokens(desc)) | {norm(t) for t in hangul_tokens(desc)}
        hits, seen = [], set()
        for p in self._all_partners():
            ok = company_match(p.name, ptoks)
            method = "상호일치" if ok else ""
            if not ok and any(norm(a) and norm(a) in d for a in p.aliases):
                method = "별칭일치"
            if not ok and p.ceo and len(norm(p.ceo)) >= 2 and norm(p.ceo) in tokens:
                method = "대표자명일치"
            if method and id(p) not in seen:
                seen.add(id(p))
                hits.append((p, method))
        # 상호 일치가 있으면 대표자명 일치보다 우선
        if any(m != "대표자명일치" for _, m in hits):
            hits = [h for h in hits if h[1] != "대표자명일치"]
            # 같은 상호가 여럿이면 통장에 같이 찍힌 대표자명으로 하나를 고른다 ('나영(사모터스)')
            if len(hits) > 1:
                by_ceo = [h for h in hits if h[0].ceo and norm(h[0].ceo) in tokens]
                words = set(re.split(r"\s+", desc or ""))
                exact = [h for h in hits if re.sub(r"\s+", "", h[0].name) in words]   # '(주)한결모터스' 단어 그대로
                if len(by_ceo) == 1:
                    hits = [(by_ceo[0][0], "상호+대표자명일치")]
                elif len(exact) == 1:
                    hits = [(exact[0][0], "상호일치(표기 동일)")]
        return hits

    def _note_partner(self, txn: Txn, p: Partner | None, method: str, extra: str = ""):
        self.match_log.append({
            "통장행": txn.row, "일자": txn.date, "구분": txn.side, "금액": txn.amount,
            "통장표시": txn.desc, "거래처": p.name if p else "", "거래처코드": p.code if p else "",
            "대표자": p.ceo if p else "", "인식방법": method,
            "원장등록": ("등록" if p.in_ledger else "미등록(신규)") if p else "", "비고": extra,
        })

    # ----- 공통 -----
    def acct(self, name: str) -> str:
        return self.cfg["accounts"].get(name, "")

    def _bank_line(self, txn: Txn, partner=None, memo="") -> Line:
        side = "차변" if txn.side == "입금" else "대변"
        if self.cfg.get("bank_partner_code"):   # 이 통장의 금융거래처코드 (고객 설정)
            partner = Partner(name=self.cfg.get("bank_partner_name", ""), code=str(self.cfg["bank_partner_code"]))
        return Line(side, self.cfg["bank_account"], txn.amount, partner, memo or txn.desc)

    def _add(self, txn: Txn, lines: list[Line], basis: str, review: str = ""):
        # 사용자가 나중에 직접 확인하겠다고 한 거래 (고객 설정 'review_keywords')
        if self.cfg.get("review_keywords") and self._has(txn.party or txn.desc, self.cfg["review_keywords"]):
            review = "; ".join(x for x in (review, "사용자 확인 예정") if x)
        e = Entry(txn.date, lines, basis, review, f"통장 {txn.row}행")
        assert e.balanced(), f"차대 불일치: {basis} {txn}"
        self.entries.append(e)

    def _has(self, desc: str, keywords) -> bool:
        d = norm(desc)
        return any(norm(k) in d for k in keywords)

    def _fee_targets(self, txn: Txn) -> list[tuple[int, int]]:
        """출금액을 대조할 금액 후보: (그대로, 0) 먼저, 다음에 (고정 송금수수료를 뺀 금액, 수수료).
        수수료는 설정된 고정 금액(기본 500원)만 인정 — 다른 차액은 수수료로 보지 않는다."""
        fee = int(self.cfg.get("transfer_fee", 0)) if txn.side == "출금" else 0
        return [(txn.amount, 0)] + ([(txn.amount - fee, fee)] if 0 < fee < txn.amount else [])

    def _fee_line(self, fee: int) -> Line:
        return Line("차변", self.cfg.get("transfer_fee_account", "지급수수료"), fee, None, "송금수수료")

    # ----- 급여·사업소득 -----
    def _name_in(self, name: str, desc: str) -> bool:
        """통장 이름 비교. 철자가 다른 경우는 config 'name_aliases' {자료 이름: [통장 이름, ...]}."""
        return name_in(name, desc) or any(name_in(a, desc) for a in self.cfg.get("name_aliases", {}).get(name, []))

    def _try_payroll(self, txn: Txn) -> bool:
        mine = [p for p in self.payroll if self._name_in(p.name, txn.desc)]
        if mine:
            return self._pay_person(txn, mine)
        # 이름 없이 일괄 이체: 적요에 '급여' 등이 있고 같은 월 미지급 실지급액 합계 일치
        # (금액만 같은 남의 송금이 급여로 붙지 않도록 키워드 필수)
        if not self._has(txn.desc, self.cfg["payroll_keywords"]):
            return False
        by_month = defaultdict(list)
        for p in self.payroll:
            if p.remaining > 0:
                by_month[p.month].append(p)
        for target, fee in self._fee_targets(txn):
            for m, rows in by_month.items():
                if sum(r.remaining for r in rows) == target:
                    for r in rows:
                        r.paid_amt, r.paid = r.net, True
                    names = ", ".join(r.name for r in rows)
                    self._note_partner(txn, None, "급여대장", names)
                    self._settle_payable(txn, target, fee, f"{m or ''} 급여 {names}".strip(),
                                         f"급여대장 {m or ''} 실지급액 합계 일치({len(rows)}명)")
                    return True
        return False

    def _pay_person(self, txn: Txn, rows: list[PayRow], emit: bool = True) -> bool:
        """급여대장에 있는 사람에게 나간 돈 → 미지급비용 털어내기.
        같은 달 급여부터 채우고(분할지급 허용), 그 달 급여가 없으면 가장 최근의 덜 지급된 달.
        금액이 남은 급여와 정확히(또는 +500원) 맞으면 그 달을 완납 처리."""
        ym = txn.date.strftime("%Y-%m")
        open_rows = [p for p in rows if p.remaining > 0]
        same = [p for p in open_rows if p.month == ym or p.month is None]
        earlier = sorted((p for p in open_rows if p.month and p.month < ym), key=lambda p: p.month, reverse=True)
        order = same + earlier
        target_row, fee = None, 0
        for amt, f in self._fee_targets(txn):
            target_row = next((p for p in order if p.remaining == amt), None)
            if target_row:
                fee = f
                break
        row = target_row or (order[0] if order else None)
        paid = txn.amount - fee
        applied = min(paid, row.remaining) if row else 0
        excess = paid - applied
        if row:
            row.paid_amt += applied
            row.paid = row.remaining <= 0
        if not emit:
            return True

        who = f"{row.month or ''} 급여 {row.name}".strip() if row else f"급여 {rows[0].name}"
        lines = []
        if applied:
            lines.append(Line("차변", self.cfg["payroll_payable_account"], applied, None, f"{who} 지급"))
        if fee:
            lines.append(self._fee_line(fee))
        if excess:
            lines.append(Line("차변", self.cfg["payroll_overpaid_account"], excess, None, f"{who} 초과 지급"))
        lines.append(self._bank_line(txn, memo=f"{who} 지급"))

        if row is None:
            basis = f"급여대장 {rows[0].name} — 지급할 급여 잔액 없음"
        elif row.remaining == 0 and applied == row.net:
            basis = f"급여대장 {row.month} {row.name} 차인지급액 일치"
        elif row.remaining == 0:
            basis = f"급여대장 {row.month} {row.name} 분할지급 완료 (차인지급액 {row.net:,})"
        else:
            basis = f"급여대장 {row.month} {row.name} 분할지급 — 남은 금액 {row.remaining:,}"
        if fee:
            basis += f" + 송금수수료 {fee:,}"
        review = ""
        if excess:
            review = (f"급여대장 차인지급액보다 {excess:,} 더 지급 — 상여·가불·퇴직금 등 확인 "
                      f"(임시로 {self.cfg['payroll_overpaid_account']})")
        review = "; ".join(x for x in (review, self._dual_review(txn.desc)) if x)
        self._note_partner(txn, None, "급여대장", rows[0].name)
        self._add(txn, lines, basis, review)
        return True

    def prime(self, txns: list[Txn]):
        """작업 시작일 이전 통장 거래로 급여 지급 누적만 반영(전표는 만들지 않음)."""
        for t in txns:
            if t.side == "출금" and not self._has(t.raw_text or t.desc, self.cfg.get("exclude_keywords", [])):
                mine = [p for p in self.payroll if self._name_in(p.name, t.desc)]
                if mine:
                    self._pay_person(t, mine, emit=False)
                    continue
                biz = [b for b in self.business if self._name_in(b.name, t.desc)]
                if biz:
                    self._pay_business(t, biz, emit=False)
                    continue
                self._try_purchase(t, emit=False)

    def _dual_review(self, desc_or_names) -> str:
        names = [n for n in self.dual_names if self._name_in(n, desc_or_names)]
        return f"급여·사업소득 양쪽에 있는 사람({', '.join(names)}) — 별도 확인" if names else ""

    def _settle_payable(self, txn: Txn, net: int, fee: int, who: str, basis: str, review: str = ""):
        """급여·사업소득 지급: 이미 잡힌 미지급비용(실지급액)을 털어낸다.
        차변 미지급비용 / (차변 지급수수료 500) / 대변 보통예금. 500원 외 차액은 확인필요."""
        acc = self.cfg["payroll_payable_account"]
        paid = txn.amount - fee
        lines = [Line("차변", acc, min(net, paid), None, f"{who} 지급")]
        if fee:
            lines.append(self._fee_line(fee))
            basis += f" + 송금수수료 {fee:,}"
        if paid > net:    # 실지급액보다 더 나감 → 초과분 가지급금
            lines.append(Line("차변", self.cfg["payroll_overpaid_account"], paid - net, None, f"{who} 초과 이체"))
        lines.append(self._bank_line(txn, memo=f"{who} 지급"))
        if paid < net:
            review = "; ".join(x for x in (review, f"미지급비용 {net - paid:,} 잔액 남음") if x)
        review = "; ".join(x for x in (review, self._dual_review(txn.desc)) if x)
        self._add(txn, lines, basis, review)

    def _book_payroll(self, txn: Txn, rows: list[PayRow], basis: str, review: str = "", fee: int = 0):
        for p in rows:
            p.paid = True
        names = ", ".join(p.name for p in rows)
        self._note_partner(txn, None, "급여대장", names)
        self._settle_payable(txn, sum(p.net for p in rows), fee, f"{rows[0].month or ''} 급여 {names}".strip(),
                             basis, review)

    def _try_business(self, txn: Txn) -> bool:
        mine = [b for b in self.business if self._name_in(b.name, txn.desc)]
        if not mine:
            return False
        return self._pay_business(txn, mine)

    def _pay_business(self, txn: Txn, rows: list[BizRow], emit: bool = True) -> bool:
        """사업소득자 송금 → 미지급비용 털어내기 (사용자 지정).
        - 끝자리가 500원인 송금은 건마다 500원을 지급수수료로 분리
        - 그달 금액과 안 맞아도 사람별로 누적해서 털어냄(오래된 달부터). 자료 전체 잔액을 넘는 금액만 확인필요"""
        fee_amt = int(self.cfg.get("transfer_fee", 0))
        fee = fee_amt if self.cfg.get("business_fee_per_transfer", True) and fee_amt \
            and txn.amount % 1000 == fee_amt % 1000 and txn.amount > fee_amt else 0
        if not self.cfg.get("business_fee_per_transfer", True):
            fee = next((f for amt, f in self._fee_targets(txn) if any(b.remaining == amt for b in rows)), 0)
        paid = txn.amount - fee
        left = paid
        months = []
        for b in sorted(rows, key=lambda b: b.month or ""):
            if left <= 0:
                break
            take = min(left, b.remaining)
            if take > 0:
                b.paid_amt += take
                b.paid = b.remaining <= 0
                left -= take
                months.append(b.month or "")
        applied, excess = paid - left, left
        if not emit:
            return True

        name = rows[0].name
        who = f"사업소득 {name}"
        lines = []
        if applied:
            lines.append(Line("차변", self.cfg["payroll_payable_account"], applied, None, f"{who} 지급"))
        if fee:
            lines.append(self._fee_line(fee))
        if excess:
            lines.append(Line("차변", self.cfg["payroll_overpaid_account"], excess, None, f"{who} 초과 지급"))
        lines.append(self._bank_line(txn, memo=f"{who} 지급"))
        rest = sum(b.remaining for b in rows)
        basis = f"사업소득 {name} 누적 털기" + (f" ({', '.join(months)}분)" if months else "") + \
            f" — 남은 미지급 {rest:,}" + (f" + 송금수수료 {fee:,}" if fee else "")
        review = ""
        if excess:
            review = (f"사업소득 자료의 차인지급액 합계보다 {excess:,} 더 지급 — 확인 필요 "
                      f"(임시로 {self.cfg['payroll_overpaid_account']})")
        review = "; ".join(x for x in (review, self._dual_review(txn.desc)) if x)
        self._note_partner(txn, None, "사업소득 지급내역", name)
        self._add(txn, lines, basis, review)
        return True

    # ----- 예수금 납부 -----
    def _withheld(self, month: str, attr: str, src) -> int | None:
        rows = [r for r in src if r.month == month]
        return sum(getattr(r, attr) for r in rows) if rows else None

    def _ref_month(self, txn: Txn, rows) -> str | None:
        """납부 대상 귀속월: 직전월 우선, 없으면 자료에 있는 가장 가까운 과거 월."""
        months = sorted({r.month for r in rows if r.month})
        pm = prev_month(txn.date.strftime("%Y-%m"))
        if pm in months:
            return pm
        past = [m for m in months if m <= txn.date.strftime("%Y-%m")]
        return past[-1] if past else (months[-1] if months else None)

    def _ledger_withheld(self, txn: Txn, typ: str) -> tuple[int, str] | None:
        """예수금 원장의 직전월(없으면 가장 가까운 과거 월) 예수액."""
        rows = [r for r in self.withheld if r.typ == typ and r.month]
        if not rows:
            return None
        m = self._ref_month(txn, rows)
        return sum(r.amount for r in rows if r.month == m), m

    def _expected(self, txn: Txn, attr: str) -> tuple[int | None, str]:
        led = self._ledger_withheld(txn, attr)
        if led:
            return led[0], f"예수금원장 {led[1]}"
        rows = self.payroll + self.business
        if not rows:
            return None, ""
        if all(r.month is None for r in rows):
            return sum(getattr(r, attr, 0) for r in rows), "자료 전체"
        m = self._ref_month(txn, rows)
        tot = sum(getattr(r, attr, 0) for r in rows if r.month == m)
        return tot, f"{m} 귀속"

    def _tax_account(self, item: str) -> str | None:
        """세목 → 계정 (config 'tax_item_accounts', 세목 이름에 키워드가 들어 있으면)."""
        for kw, acc in self.cfg.get("tax_item_accounts", {}).items():
            if kw in item.replace(" ", ""):
                return acc
        return None

    def _try_tax_certificate(self, txn: Txn) -> bool:
        """국세 납부내역증명·지방세 과세증명서와 금액·납부일(±3일)이 맞는 세금 출금 → 세목별 계정.
        금액이 0~5% 더 많은 출금(납부지연가산세 등)은 가장 가까운 납부와 연결하고 확인필요."""
        if not self._has(txn.desc, self.cfg["national_tax_keywords"] + self.cfg["local_tax_keywords"]):
            exact_only = True
        else:
            exact_only = False
        near = [tp for tp in self.tax_payments if tp.matched_row is None and tp.date
                and abs((tp.date - txn.date).days) <= 3]
        exact = [tp for tp in near if tp.amount == txn.amount]
        tp, review, surcharge = None, "", 0
        if exact:
            tp = exact[0]
        elif not exact_only:
            over = [t for t in near if t.amount < txn.amount <= t.amount * 1.05]
            if over:
                tp = min(over, key=lambda t: txn.amount - t.amount)
                surcharge = txn.amount - tp.amount   # 납부지연가산세 등 → 잡손실 (사용자 지정)
        if tp is None:
            return False
        tp.matched_row = txn.row
        acc = self._tax_account(tp.item)
        if not acc:
            acc = self.cfg["unknown_withdrawal_account"]
            review = "; ".join(x for x in (f"세목 '{tp.item}' 계정 미지정 — 확인 필요 (임시로 {acc})", review) if x)
        label = f"{tp.item} {tp.note}".strip()
        lines = [Line("차변", acc, tp.amount, None, f"{label} 납부")]
        if surcharge:
            lines.append(Line("차변", self.cfg.get("tax_surcharge_account", "잡손실"), surcharge, None,
                              f"{label} 가산세 등 차액"))
        lines.append(self._bank_line(txn, memo=f"{label} 납부"))
        self._note_partner(txn, None, "세금 납부내역", tp.item)
        basis = f"납부증명 {tp.date} {tp.year} {label} {tp.amount:,} 일치" + \
            (f" + 차액 {surcharge:,} 잡손실" if surcharge else "")
        self._add(txn, lines, basis, review)
        return True

    def _try_tax(self, txn: Txn) -> bool:
        if txn.side != "출금":
            return False
        if self.tax_payments and self._try_tax_certificate(txn):
            return True
        if self._has(txn.desc, ["재산세", "자동차세", "주민세", "면허세"]):
            return False  # 원천세가 아닌 지방세 → 키워드 규칙(세금과공과)
        is_local = self._has(txn.desc, self.cfg["local_tax_keywords"])
        is_nat = not is_local and self._has(txn.desc, self.cfg["national_tax_keywords"])
        if not (is_local or is_nat):
            return False
        attr = "local_tax" if is_local else "income_tax"
        label = "지방소득세(원천)" if is_local else "원천세(소득세)"
        exp, basis = self._expected(txn, attr)
        review, acc = "", self.cfg["withholding_account"]
        if exp is None:
            review = "예수금 원장·급여 자료가 없어 예수금 금액 대조 못함"
        elif exp != txn.amount:
            review = (f"{basis} 예수 {label} {exp:,} ≠ 납부액 {txn.amount:,} "
                      "— 부가세·법인세 등 다른 세금이거나 다른 월분일 수 있음")
        if review:   # 근거와 연결 안 됨 → 원칙대로 출금은 외상매입금
            acc = self.cfg["unknown_withdrawal_account"]
            review += f" (임시로 {acc})"
        lines = [Line("차변", acc, txn.amount, None, f"{label} 납부"),
                 self._bank_line(txn, memo=f"{label} 납부")]
        self._note_partner(txn, None, "예수금 납부", label)
        self._add(txn, lines, f"{label} 납부 ({basis} 대조)" if basis else f"{label} 납부", review)
        return True

    def _insurance_kind(self, desc: str):
        """(표시명, 예수금 종류 또는 None, 회사부담 계정 설정키, 확인사유)."""
        c = self.cfg
        if self._has(desc, c["pension_keywords"]):
            return "국민연금", "pension", "pension_company_account", ""
        if self._has(desc, c["health_keywords"]):
            return "건강보험", "health", "health_company_account", ""
        if self._has(desc, c["employment_keywords"]):      # '고용산재' 통합 납부도 여기(고용 예수금 차감)
            return "고용보험", "employment", "employment_company_account", ""
        if self._has(desc, c["accident_keywords"]):        # 산재는 근로자 부담 없음 → 전액 비용
            return "산재보험", None, "accident_account", ""
        if self._has(desc, c["comwel_keywords"]):          # 근로복지공단만 찍혀 고용/산재 구분 불가
            return "고용보험", "employment", "employment_company_account", \
                "근로복지공단 출금 — 고용/산재 구분 불가, 고용보험으로 보고 예수금 차감함. 산재면 전액 보험료로 수정"
        return None

    def _try_insurance(self, txn: Txn) -> bool:
        if txn.side != "출금":
            return False
        kind = self._insurance_kind(txn.desc)
        if not kind:
            return False
        label, attr, comp_acc, review = kind
        if attr is None:
            emp, basis = 0, "근로자 부담분 없음 — 전액 회사부담"
        else:
            # 근로자부담분 = 직전월 예수금 (예수금 원장 우선, 없으면 급여대장 공제액) 중 아직 안 쓴 금액
            led = self._ledger_withheld(txn, attr)
            if led:
                emp, m = led
                src = "예수금원장"
            elif self.payroll:
                m = self._ref_month(txn, self.payroll) if any(p.month for p in self.payroll) else None
                emp = sum(getattr(p, attr) for p in self.payroll if p.month == m or m is None)
                src = "급여대장 공제액"
            else:
                m, src = None, ""
                emp = txn.amount // 2 if attr != "employment" else 0
                review = "예수금 원장·급여대장 없이 근로자/회사부담 구분 — 확인 필요"
            if src:
                used = self._wh_used[(attr, m)]
                emp -= used
                basis = f"{src} {m or ''} 예수금 {emp + used:,}" + (f" (이미 차감 {used:,})" if used else "")
            else:
                basis = "예수금 자료 없음 — 근로자:회사 1:1 가정" if emp else "예수금 자료 없음"
            emp = max(emp, 0)
            if emp > txn.amount:
                review = f"예수금 {emp:,} > 납부액 {txn.amount:,} — 예수금 잔액이 남음, 미납·분납 여부 확인"
                emp = txn.amount
            if src:
                self._wh_used[(attr, m)] += emp
        lines = []
        if emp:
            lines.append(Line("차변", self.cfg["withholding_account"], emp, None, f"{label} 근로자부담분"))
        if txn.amount - emp:
            lines.append(Line("차변", self.cfg[comp_acc], txn.amount - emp, None, f"{label} 회사부담분"))
        lines.append(self._bank_line(txn, memo=f"{label} 납부"))
        self._note_partner(txn, None, "4대보험 납부", label)
        self._add(txn, lines, f"{label} 납부 ({basis})", review)
        return True

    # ----- 세금계산서 -----
    def _ap_key(self, p: Partner) -> str:
        return p.code or p.key

    def _try_purchase(self, txn: Txn, emit: bool = True) -> bool:
        """매입 세금계산서 거래처로 나간 돈 → 외상매입금(거래처코드) 누적 털기.
        계약금·잔금 분할, 여러 대 일괄 송금도 거래처별 세금계산서 합계 안이면 확정.
        끝자리 500원은 지급수수료. 세금계산서 합계를 넘는 금액은 확인필요(선급 등)."""
        if txn.side != "출금":
            return False
        all_hits = self.identify(txn.party or txn.desc)
        if len(all_hits) != 1:  # 동명 대표·같은 상호 여러 곳 → 세금계산서 유무로 고르지 않는다 (명확하지 않음)
            return False
        hits = [(p, m) for p, m in all_hits if any(i.kind == "매입" and i.partner is p for i in self.invoices)]
        if len(hits) != 1:     # 후보가 여럿이면 억지로 고르지 않는다 (사용자 원칙) → 미연결
            return False
        p, method = hits[0]
        fee_amt = int(self.cfg.get("transfer_fee", 0))
        fee = fee_amt if fee_amt and txn.amount > fee_amt and txn.amount % 1000 == fee_amt % 1000 else 0
        paid = txn.amount - fee
        open_before = self._ap_open(p)
        self.ap_paid[self._ap_key(p)] += paid
        if not emit:
            return True
        exact = next((i for i in self.invoices if i.kind == "매입" and i.partner is p and i.total == paid), None)
        lines = [Line("차변", self.cfg["payable_account"], paid, p, txn.desc)]
        if fee:
            lines.append(self._fee_line(fee))
        lines.append(self._bank_line(txn, p))
        basis = f"[{method}] 매입 {p.name}" + (f" 세금계산서 {exact.date} {exact.total:,} 일치" if exact else
                                               f" 누적 털기 (남은 외상매입금 {max(open_before - paid, 0):,})")
        if fee:
            basis += f" + 송금수수료 {fee:,}"
        review = ""
        if paid > open_before:
            review = (f"{p.name} 매입 세금계산서 합계보다 {paid - max(open_before, 0):,} 더 지급 — "
                      "선급금·계약금 또는 세금계산서 누락 확인")
        self._note_partner(txn, p, method, basis)
        self._add(txn, lines, basis, review)
        return True

    def _ap_open(self, p: Partner) -> int:
        total = sum(i.total for i in self.invoices if i.kind == "매입" and i.partner is p)
        return total - self.ap_paid[self._ap_key(p)]

    def _invoice_candidates(self, kind: str) -> list[Invoice]:
        return [i for i in self.invoices if i.kind == kind and not i.matched_rows]

    @staticmethod
    def _subset_sum(invs: list[Invoice], target: int, limit: int = 14) -> list[Invoice] | None:
        invs = [i for i in invs if i.total > 0][:limit]
        for r in range(2, min(len(invs), 6) + 1):
            for combo in itertools.combinations(invs, r):
                if sum(i.total for i in combo) == target:
                    return list(combo)
        return None

    def _try_invoice(self, txn: Txn, loose: bool = False) -> bool:
        kind = "매출" if txn.side == "입금" else "매입"
        hits = self.identify(txn.desc)
        cands = self._invoice_candidates(kind)
        ar_ap = self.cfg["receivable_account"] if kind == "매출" else self.cfg["payable_account"]

        # 매입대금 송금 시 고정 송금수수료가 같이 출금되는 경우: (세금계산서와 맞출 금액, 수수료)
        targets = self._fee_targets(txn)

        def book(p: Partner, invs: list[Invoice], method: str, basis: str, review: str = "", fee: int = 0):
            for i in invs:
                i.matched_rows.append(txn.row)
            if kind == "매출":
                lines = [self._bank_line(txn, p), Line("대변", ar_ap, txn.amount, p, txn.desc)]
            else:
                lines = [Line("차변", ar_ap, txn.amount - fee, p, txn.desc)]
                if fee:
                    lines.append(self._fee_line(fee))
                    basis += f" + 송금수수료 {fee:,}"
                lines.append(self._bank_line(txn, p))
            self._note_partner(txn, p, method, basis)
            self._add(txn, lines, basis, review)
            return True

        # 1) 거래처 인식 + 금액 일치(단건/복수)
        by_partner = []
        for target, f in targets:
            for p, method in hits:
                pinv = [i for i in cands if i.partner is p]
                one = [i for i in pinv if i.total == target]
                if one:
                    one.sort(key=lambda i: abs(((i.date or txn.date) - txn.date).days))
                    by_partner.append((p, method, [one[0]], f"{kind}세금계산서 {one[0].date} {one[0].total:,} 일치", f))
                    continue
                combo = self._subset_sum(pinv, target)
                if combo:
                    by_partner.append((p, method, combo, f"{kind}세금계산서 {len(combo)}건 합계 일치", f))
            if by_partner:
                break  # 수수료 없이 맞으면 수수료 분리 시도 안 함
        if not loose and len(by_partner) == 1:
            p, method, invs, basis, f = by_partner[0]
            return book(p, invs, method, f"[{method}] {basis}", fee=f)
        if len(by_partner) > 1:
            return False      # 후보 여러 곳 → 연결하지 않음 (거래처코드 공란, 미연결 처리)

        # 2) 거래처는 찾았으나 금액 불일치 / 세금계산서 없음
        if len(hits) > 1:
            return False      # 후보 여러 곳 → 연결하지 않음
        if hits:
            p, method = hits[0]
            review = f"동명 대표자/유사 상호 후보 {len(hits)}곳: " + ", ".join(h[0].name for h in hits) \
                if len(hits) > 1 else ""
            has_inv = any(i.partner is p for i in self.invoices if i.kind == kind)
            if has_inv == loose or (not has_inv and not p.in_ledger):
                return False  # 금액 불일치(계산서 있음)는 1차, 원장만 일치는 키워드 규칙 뒤 2차에서 처리
            if has_inv:
                why = f"{kind}세금계산서 금액과 불일치 (분할·일괄결제·선{'수' if kind == '매출' else '급'}금 가능)"
            else:
                why = f"이번 자료에 {kind}세금계산서 없음 (전기 이월 채권·채무 또는 계산서 누락 확인)"
            return book(p, [], method, f"[{method}] 거래처원장", "; ".join(x for x in (why, review) if x))

        # 3) 거래처명은 못 찾았지만 금액이 유일하게 일치
        if loose or (kind == "매입" and not self.cfg.get("purchase_amount_only_match", False)):
            return False
        for target, f in targets:
            same = [i for i in cands if i.total == target]
            if len(same) == 1:
                i = same[0]
                return book(i.partner, [i], "금액만 일치", f"{kind}세금계산서 {i.date} {i.total:,} 금액 일치",
                            f"통장 표시명 '{txn.desc}' 과 거래처 '{i.partner.name}' 연결 근거가 금액뿐 — 확인 필요",
                            fee=f)
        return False

    # ----- 키워드 규칙 / 대표자 / 미확인 -----
    def _try_owner(self, txn: Txn) -> bool:
        owners = self.cfg.get("owner_names") or []
        if not any(self._name_in(o, txn.desc) for o in owners):
            return False
        corp = self.cfg.get("entity_type", "법인") == "법인"
        if txn.side == "입금":
            acc, why = ("가수금", "대표자 입금") if corp else ("인출금", "사업주 입금")
            lines = [self._bank_line(txn), Line("대변", acc, txn.amount, None, why)]
        else:
            acc, why = ("가지급금", "대표자 출금") if corp else ("인출금", "사업주 인출")
            lines = [Line("차변", acc, txn.amount, None, why), self._bank_line(txn)]
        self._note_partner(txn, None, "대표자 본인", why)
        self._add(txn, lines, why, "대표자 거래 — 용도 확인" if corp else "")
        return True

    @staticmethod
    def _party_name(t: Txn) -> str:
        """통장 상대방 이름: 차량번호 등을 빼고, '적요'와 '의뢰인'이 같은 이름이면 한 번만."""
        words = [w for w in re.split(r"\s+", re.sub(r"\d{2,3}[가-힣]\d{3,4}", " ", t.party or t.desc)) if w]
        half = len(words) // 2
        if half and len(words) % 2 == 0 and words[:half] == words[half:]:
            words = words[:half]
        return " ".join(words)

    def _rule_match(self, rule: dict, txn: Txn) -> bool:
        if rule.get("keywords") and self._has(txn.desc, rule["keywords"]):
            return True
        if rule.get("pattern"):
            target = (txn.party or txn.desc) if rule.get("on") == "party" else txn.desc
            return bool(re.search(rule["pattern"], target.strip()))
        return False

    def _try_rules(self, txn: Txn) -> bool:
        # 고객 설정의 extra_rules 와 confirmed 규칙은 사용자가 정한 것 — 그 외 기본 규칙은 확인필요(노란색)
        for rule in [dict(r, confirmed=True) for r in self.cfg.get("extra_rules", [])] + self.cfg.get("rules", []):
            if rule.get("side") not in (txn.side, "both", None):
                continue
            if not self._rule_match(rule, txn):
                continue
            acc, memo = rule["account"], rule.get("memo") or txn.desc
            p = self.identify(txn.desc)
            partner = p[0][0] if len(p) == 1 else None      # 후보가 여럿이면 연결하지 않음
            if rule.get("partner_from_party") and not partner:
                partner = Partner(name=self._party_name(txn), in_ledger=False)
            if rule.get("partner_code"):      # 규칙에 거래처코드 지정 (예: 다른 보통예금 계좌의 금융거래처)
                partner = Partner(name=rule.get("partner_name", ""), code=str(rule["partner_code"]))
            bank_partner = None if rule.get("partner_code") else partner
            if txn.side == "입금" and rule.get("negative_debit"):
                # 비용 환급 등: 대변 대신 차변에 마이너스 (예: 차변 보통예금 / 차변 보험료 −금액)
                lines = [self._bank_line(txn, bank_partner), Line("차변", acc, -txn.amount, partner, memo)]
            elif txn.side == "입금":
                lines = [self._bank_line(txn, bank_partner), Line("대변", acc, txn.amount, partner, memo)]
            else:
                lines = [Line("차변", acc, txn.amount, partner, memo), self._bank_line(txn, bank_partner)]
            self._note_partner(txn, partner, "키워드 규칙",
                               "/".join(rule.get("keywords", [])[:3]) or rule.get("memo", rule.get("pattern", "")))
            review = ""
            if rule.get("review") or (self.cfg.get("flag_unconfirmed_rules", True) and not rule.get("confirmed")):
                review = f"기본 규칙으로 {acc} 처리 — 계정 확인 필요"
            self._add(txn, lines, f"키워드 규칙 → {acc}", review)
            return True
        return False

    def _fallback(self, txn: Txn):
        if txn.side == "입금":
            acc = self.cfg["unknown_deposit_account"]
            lines = [self._bank_line(txn), Line("대변", acc, txn.amount, None, txn.desc)]
        else:
            acc = self.cfg["unknown_withdrawal_account"]
            # 거래처를 못 찾아도 끝자리가 송금수수료(500원)면 수수료로 분리 (사용자 지정)
            fee = int(self.cfg.get("transfer_fee", 0))
            fee = fee if self.cfg.get("unlinked_fee_split", True) and fee and txn.amount > fee \
                and txn.amount % 1000 == fee % 1000 else 0
            lines = [Line("차변", acc, txn.amount - fee, None, txn.desc)]
            if fee:
                lines.append(self._fee_line(fee))
            lines.append(self._bank_line(txn))
        self._note_partner(txn, None, "미확인")
        review = "; ".join(x for x in ("거래처 연결 안 됨 — 거래처·계정 확인 필요", self._dual_review(txn.desc)) if x)
        self._add(txn, lines, f"매칭 실패 → {acc}" + (f" + 송금수수료 {fee:,}" if txn.side == "출금" and fee else ""),
                  review)

    @staticmethod
    def _party_keys(t: Txn) -> set[str]:
        text = t.party or t.desc
        keys = {norm(x) for x in hangul_tokens(text)}
        keys |= {w.upper() for w in re.findall(r"[A-Za-z]{4,}", text)}
        keys = {norm(re.sub("|".join(map(re.escape, CORP_TOKENS)), "", k)) for k in keys}
        return {k for k in keys if len(k) >= 3 and k not in GENERIC_WORDS}

    def _find_refund_pairs(self, txns: list[Txn]) -> dict[int, tuple]:
        """같은 상대에게 보낸 돈이 days 안에 그대로(또는 송금수수료만 빼고) 돌아온 거래 → 반환 쌍.
        {행번호: ('out'|'in', 돌려받은 금액, 수수료, 상대 행번호)}"""
        days = int(self.cfg.get("refund_pair_days", 30))
        fee = int(self.cfg.get("transfer_fee", 0))
        pairs, used = {}, set()
        outs = [t for t in txns if t.side == "출금"]
        ins = [t for t in txns if t.side == "입금"]
        for w in outs:
            wk = self._party_keys(w)
            for d in ins:
                if d.row in used or w.row in used or not (0 <= (d.date - w.date).days <= days):
                    continue
                if d.amount not in (w.amount, w.amount - fee) or not (wk & self._party_keys(d)):
                    continue
                f = w.amount - d.amount
                pairs[w.row] = ("out", d.amount, f, d.row)
                pairs[d.row] = ("in", d.amount, f, w.row)
                used |= {w.row, d.row}
                break
        return pairs

    def _book_refund_pair(self, t: Txn, info: tuple):
        kind, amt, fee, other = info
        acc = self.cfg.get("refund_pair_account", "예수금")
        who = (t.party or t.desc).strip()
        if kind == "out":
            lines = [Line("차변", acc, amt, None, f"{who} 계약금 지급")]
            if fee:
                lines.append(self._fee_line(fee))
            lines.append(self._bank_line(t))
            basis = f"계약금 지급 → {acc} (통장 {other}행에서 {amt:,} 반환)" + (f" + 송금수수료 {fee:,}" if fee else "")
        else:
            lines = [self._bank_line(t), Line("대변", acc, amt, None, f"{who} 계약금 반환")]
            basis = f"계약금 반환 → {acc} (통장 {other}행 지급분)"
        review = ""
        if amt > int(self.cfg.get("refund_pair_review_over", 10_000_000)):
            review = f"보냈다가 돌려받은 {amt:,} — 계약금 반환이 맞는지 확인 (임시로 {acc})"
        self._note_partner(t, None, "계약금 반환 쌍", who)
        self._add(t, lines, basis, review)

    def journalize_bank(self, txns: list[Txn]):
        self.excluded: list[Txn] = []
        excl = self.cfg.get("exclude_keywords", [])
        live = [t for t in txns if not (excl and self._has(t.raw_text or t.desc, excl))]
        pairs = self._find_refund_pairs(live)
        for t in txns:
            if excl and self._has(t.raw_text or t.desc, excl):
                self.excluded.append(t)  # 체크카드 등 — 세무사랑에서 따로 처리, 통장 전표에서 제외
                continue
            if t.row in pairs:
                self._book_refund_pair(t, pairs[t.row])
                continue
            if t.side == "출금":
                steps = (self._try_payroll, self._try_business, self._try_tax,
                         self._try_insurance, self._try_owner, self._try_purchase, self._try_invoice, self._try_rules,
                         lambda t: self._try_invoice(t, loose=True))
            else:
                steps = (self._try_owner, self._try_invoice, self._try_rules,
                         lambda t: self._try_invoice(t, loose=True))
            if not any(step(t) for step in steps):
                self._fallback(t)

    # ----- 세금계산서 자체 전표(매입매출전표) -----
    def invoice_entries(self) -> list[Entry]:
        out = []
        for i in self.invoices:
            p = i.partner
            if i.kind == "매출":
                sales = self.cfg["default_sales_account"]
                lines = [Line("차변", self.cfg["receivable_account"], i.total, p, i.item),
                         Line("대변", sales, i.supply, p, i.item)]
                if i.vat:
                    lines.append(Line("대변", "부가세예수금", i.vat, p, i.item))
                typ = "11.과세" if i.vat else "13.면세"
            else:
                acc = p.account if p.account in self.cfg["accounts"] else self.cfg["default_purchase_account"]
                lines = [Line("차변", acc, i.supply, p, i.item)]
                if i.vat:
                    lines.append(Line("차변", "부가세대급금", i.vat, p, i.item))
                lines.append(Line("대변", self.cfg["payable_account"], i.total, p, i.item))
                typ = "51.과세" if i.vat else "53.면세"
            out.append(Entry(i.date or date.today(), lines, f"{i.kind} {typ}", "",
                             f"{i.kind}세금계산서 {i.idx + 1}"))
        return out


# ---------------------------------------------------------------- 엑셀 출력

HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)


def _sheet(wb: Workbook, title: str, header: list[str], rows: list[list], warn=None,
           money_cols: tuple[int, ...] = ()):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, BOLD
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(r)
        for c in ws[ws.max_row]:
            if isinstance(c.value, (date, datetime)):
                c.number_format = "yyyy-mm-dd"
            if warn and warn(r):
                c.fill = WARN_FILL
    for col in money_cols:
        for c in ws.iter_cols(min_col=col + 1, max_col=col + 1, min_row=2):
            for cell in c:
                cell.number_format = "#,##0"
    for i, h in enumerate(header, start=1):
        width = max([len(str(h))] + [len(str(r[i - 1])) for r in rows[:300] if r[i - 1] is not None])
        ws.column_dimensions[get_column_letter(i)].width = min(max(8, width * 1.6), 60)
    ws.freeze_panes = "A2"
    return ws


def fmt_bizno(b: str) -> str:
    return f"{b[:3]}-{b[3:5]}-{b[5:]}" if len(b) == 10 else b


def voucher_rows(entries: list[Entry], cfg: dict) -> list[list]:
    rows, seq = [], defaultdict(int)
    for e in sorted(entries, key=lambda e: e.date):
        seq[e.date] += 1
        no = f"{seq[e.date]:05d}"
        for l in e.lines:
            p = l.partner
            rows.append([e.date.month, e.date.day, no, "3" if l.side == "차변" else "4",
                         cfg["accounts"].get(l.account, ""), l.account,
                         p.code if p else "", p.name if p else "", l.memo,
                         l.amount if l.side == "차변" else None, l.amount if l.side == "대변" else None,
                         e.review, e.basis, e.source])
    return rows


VOUCHER_HEADER = ["월", "일", "전표번호", "구분(3차/4대)", "계정코드", "계정과목", "거래처코드", "거래처명",
                  "적요", "차변", "대변", "확인필요", "분개근거", "원천"]


def write_workbook(path, j: Journalizer, txns: list[Txn], inv_entries: list[Entry]):
    cfg = j.cfg
    wb = Workbook()
    wb.remove(wb.active)

    # 요약
    flagged = [e for e in j.entries if e.review]
    dep = sum(t.deposit for t in txns)
    wd = sum(t.withdrawal for t in txns)
    by_basis = defaultdict(int)
    for m in j.match_log:
        by_basis[m["인식방법"]] += 1
    excluded = getattr(j, "excluded", [])
    summary = [["통장 거래 건수", len(txns)], ["입금 합계", dep], ["출금 합계", wd],
               ["전표 제외(체크카드 등) 건수", len(excluded)],
               ["전표 제외 금액", sum(t.amount for t in excluded)],
               ["자동 분개 건수", len(j.entries) - len(flagged)], ["확인 필요 건수", len(flagged)],
               ["신규 거래처(원장 미등록)", len([p for p in j.new_partners.values()])],
               ["미결 세금계산서", len([i for i in j.invoices if not i.matched_rows])],
               ["급여·사업소득 중복 인원", len(j.dual_names)]]
    summary += [[f"  인식방법: {k}", v] for k, v in sorted(by_basis.items(), key=lambda x: -x[1])]
    _sheet(wb, "요약", ["항목", "값"], summary, money_cols=(1,))

    _sheet(wb, "일반전표(통장)", VOUCHER_HEADER, voucher_rows(j.entries, cfg), warn=lambda r: r[11],
           money_cols=(9, 10))
    if inv_entries:
        _sheet(wb, "매입매출전표(세금계산서)", VOUCHER_HEADER, voucher_rows(inv_entries, cfg),
               money_cols=(9, 10))

    _sheet(wb, "제외(체크카드)", ["통장행", "일자", "구분", "금액", "통장표시", "사유"],
           [[t.row, t.date, t.side, t.amount, t.raw_text, "체크카드 — 세무사랑에서 따로 처리, 전표 업로드 제외"]
            for t in excluded], money_cols=(3,))

    log_header = list(j.match_log[0].keys()) if j.match_log else ["통장행"]
    _sheet(wb, "거래처매칭", log_header, [list(m.values()) for m in j.match_log],
           warn=lambda r: r[-2] == "미등록(신규)" or r[-3] in ("미확인", "금액만 일치"), money_cols=(3,))

    _sheet(wb, "확인필요", ["일자", "원천", "금액", "분개근거", "확인사유"],
           [[e.date, e.source, sum(l.amount for l in e.lines if l.side == "차변"), e.basis, e.review]
            for e in flagged], money_cols=(2,))

    dual_rows = []
    for n in j.dual_names:
        for p in j.payroll:
            if p.name == n:
                dual_rows.append([n, "급여", p.month, p.gross, p.net, "지급됨" if p.paid else "통장 출금 없음"])
        for b in j.business:
            if b.name == n:
                dual_rows.append([n, "사업소득", b.month, b.gross, b.net, "지급됨" if b.paid else "통장 출금 없음"])
    _sheet(wb, "급여지급대조", ["성명", "귀속월", "지급합계", "차인지급액", "통장지급누계", "남은금액"],
           [[p.name, p.month, p.gross, p.net, p.paid_amt, p.remaining]
            for p in sorted(j.payroll, key=lambda p: (p.name, p.month or ""))],
           warn=lambda r: r[5] != 0, money_cols=(2, 3, 4, 5))

    per = {}
    for b in j.business:
        x = per.setdefault(b.name, [b.name, 0, 0, 0, 0])
        x[1] += 1
        x[2] += b.net
        x[3] += b.paid_amt
        x[4] += b.remaining
    _sheet(wb, "사업소득지급대조", ["성명", "지급월수", "차인지급액 합계", "통장지급 누계(수수료 제외)", "남은 미지급"],
           list(per.values()), warn=lambda r: r[4] != 0, money_cols=(2, 3, 4))

    _sheet(wb, "급여·사업소득중복", ["성명", "구분", "귀속월", "지급총액", "실지급액", "통장"],
           dual_rows, warn=lambda r: True, money_cols=(3, 4))

    if j.tax_payments:
        _sheet(wb, "세금납부대조", ["납부일", "귀속", "세목", "비고", "금액", "계정", "통장행"],
               [[tp.date, tp.year, tp.item, tp.note, tp.amount, j._tax_account(tp.item) or "미지정",
                 tp.matched_row or "통장에 없음"] for tp in sorted(j.tax_payments, key=lambda t: t.date)],
               warn=lambda r: r[6] == "통장에 없음" or r[5] == "미지정", money_cols=(4,))

    _sheet(wb, "신규거래처등록", ["거래처명", "사업자번호", "대표자", "근거"],
           [[p.name, fmt_bizno(p.bizno), p.ceo, "세금계산서 상대방 — 거래처 원장에 없음"]
            for p in j.new_partners.values()])

    _sheet(wb, "미결세금계산서", ["구분", "작성일자", "거래처", "거래처코드", "공급가액", "세액", "합계"],
           [[i.kind, i.date, i.partner.name, i.partner.code, i.supply, i.vat, i.total]
            for i in j.invoices if not i.matched_rows], money_cols=(4, 5, 6))

    wb.save(path)


# ---------------------------------------------------------------- 세무사랑 업로드 파일

SEMUSARANG_TEMPLATE = HERE / "templates" / "세무사랑_일반전표전송.xls"
# 세무사랑 「엑셀자료 일반전표전송 v1.2」 열 위치 (0부터)
SR_COL = {"date": 0, "gubun": 1, "code": 2, "account": 3, "pcode": 4, "pname": 5, "memo_code": 6,
          "memo": 7, "amount": 8, "bizno": 15, "ceo": 21}
SR_FIRST_ROW = 10          # 머리글(10행) 다음 줄부터 입력
SR_COMPANY_CELL = (2, 11)  # 회사명 칸
SR_BIZNO_CELL = (2, 14)    # 사업자등록번호 칸


def write_semusarang(path, entries: list[Entry], cfg: dict, company: str = "", company_bizno: str = ""):
    """세무사랑 일반전표 엑셀 업로드 파일(.xls). 양식 파일을 복사해 데이터만 채운다.
    한 전표의 차변·대변 줄은 연속으로 쓰므로 일자별로 차대가 맞는 단위로 묶인다."""
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy

    rb = xlrd.open_workbook(str(SEMUSARANG_TEMPLATE), formatting_info=True)
    wb = xl_copy(rb)
    ws = wb.get_sheet(0)
    if company:
        ws.write(*SR_COMPANY_CELL, company)
    if company_bizno:
        ws.write(*SR_BIZNO_CELL, fmt_bizno(digits(company_bizno)))

    plain = xlwt.easyxf("")
    money = xlwt.easyxf("", num_format_str="#,##0")
    warn = xlwt.easyxf("pattern: pattern solid, fore_colour light_yellow")
    warn_money = xlwt.easyxf("pattern: pattern solid, fore_colour light_yellow", num_format_str="#,##0")

    r = SR_FIRST_ROW
    for e in sorted(entries, key=lambda e: e.date):
        ymd = int(e.date.strftime("%Y%m%d"))
        for l in e.lines:
            st, st_m = (warn, warn_money) if e.review else (plain, money)
            p = l.partner
            code = cfg["accounts"].get(l.account, "")
            vals = {"date": ymd, "gubun": 3 if l.side == "차변" else 4,
                    "code": int(code) if code.isdigit() else code, "account": l.account,
                    "pcode": p.code if p else "", "pname": p.name if p else "", "memo_code": "",
                    "memo": l.memo, "bizno": fmt_bizno(p.bizno) if p and p.bizno else "",
                    "ceo": p.ceo if p else ""}
            for k, c in SR_COL.items():
                if k == "amount":
                    ws.write(r, c, l.amount, st_m)
                else:
                    ws.write(r, c, vals[k], st)
            r += 1
    wb.save(str(path))
    return r - SR_FIRST_ROW


# ---------------------------------------------------------------- 실행

def load_config(path=None, overlay=None) -> dict:
    """기본 설정 + 고객별 설정(overlay: 거래처 폴더의 config.json — 이름 별칭 등 고객 정보는 여기에)."""
    with open(path or HERE / "config.json", encoding="utf-8") as f:
        cfg = json.load(f)
    if overlay:
        with open(overlay, encoding="utf-8") as f:
            extra = json.load(f)
        for k, v in extra.items():
            cfg[k] = {**cfg[k], **v} if isinstance(v, dict) and isinstance(cfg.get(k), dict) else v
    return cfg


def run(bank, ledger=None, sales=None, purchase=None, payroll=None, business=None,
        out="분개결과.xlsx", config=None, payroll_month=None, owners=None, withholding=None,
        upload=None, company="", company_bizno="", start=None, end=None, client_config=None,
        tax_payments=None, fin_partners=None) -> Journalizer:
    cfg = load_config(config, client_config)
    fins = load_fin_partners(fin_partners) if fin_partners else []
    if fins:
        acct = bank_account_number(bank)
        own = next((f for f in fins if acct and f.bizno == acct), None)
        if own:   # 통장 계좌번호 → 금융거래처 코드·이름
            if cfg.get("bank_partner_code") and str(cfg["bank_partner_code"]) != own.code:
                raise ValueError(f"설정의 통장 거래처코드 {cfg['bank_partner_code']} 와 금융거래처 목록의 "
                                 f"{own.code}({own.name}) 가 다릅니다.")
            cfg["bank_partner_code"], cfg["bank_partner_name"] = own.code, own.name
        names = {f.code: f.name for f in fins}
        for rule in cfg.get("extra_rules", []):   # 규칙에 코드만 있으면 이름 채우기 (예: 다른 계좌의 금융거래처)
            if rule.get("partner_code") and not rule.get("partner_name"):
                rule["partner_name"] = names.get(str(rule["partner_code"]), "")
    if owners:
        cfg["owner_names"] = list(cfg.get("owner_names", [])) + owners
    all_txns = load_bank(bank, None, to_date(end) if end else None)
    sd = to_date(start) if start else None
    txns = [t for t in all_txns if not sd or t.date >= sd]
    before = [t for t in all_txns if sd and t.date < sd]
    partners = load_ledger(ledger) if ledger else []
    invoices: list[Invoice] = []
    for paths, kind in ((sales, "매출"), (purchase, "매입")):
        for p in paths or []:
            invoices += load_invoices(p, kind, start_idx=len(invoices))
    pay = [r for p in payroll or [] for r in load_payroll(p, payroll_month)]
    biz = [r for p in business or [] for r in load_business(p, payroll_month)]

    year = txns[0].date.year if txns else None
    # 예수금 원장이 여러 개면 같은 종류(국민연금 등)는 뒤에 준 파일이 앞 파일을 대신한다
    wh_by_type: dict[str, list[WithheldRow]] = {}
    for p in withholding or []:
        rows = load_withholding(p, year)
        for typ in {r.typ for r in rows}:
            wh_by_type[typ] = [r for r in rows if r.typ == typ]
    wh = [r for rows in wh_by_type.values() for r in rows]

    j = Journalizer(cfg, partners, invoices, pay, biz, wh)
    j.tax_payments = [tp for p in tax_payments or [] for tp in load_tax_payments(p)
                      if not sd or (tp.date and tp.date >= sd)]
    j.prime(before)
    j.journalize_bank(txns)
    inv_entries = j.invoice_entries()
    write_workbook(out, j, txns, inv_entries)
    j.upload_path = Path(upload) if upload else Path(out).with_name(Path(out).stem + "_세무사랑업로드.xls")
    j.upload_rows = write_semusarang(j.upload_path, j.entries, cfg, company, company_bizno)
    return j


def main():
    ap = argparse.ArgumentParser(description="통장 거래내역 자동 분개")
    ap.add_argument("--bank", required=True, help="통장 거래내역 (xlsx/xls/csv)")
    ap.add_argument("--ledger", help="거래처 원장 (거래처명·사업자번호·대표자)")
    ap.add_argument("--sales", nargs="*", help="매출 세금계산서/계산서 목록 (홈택스 엑셀)")
    ap.add_argument("--purchase", nargs="*", help="매입 세금계산서/계산서 목록 (홈택스 엑셀)")
    ap.add_argument("--payroll", nargs="*", help="급여대장")
    ap.add_argument("--business", nargs="*", help="사업소득 지급내역")
    ap.add_argument("--withholding", nargs="*",
                    help="예수금 거래처원장/계정별원장 — 직전월 예수금으로 4대보험·원천세 근로자분 산정")
    ap.add_argument("--fin", help="금융기관거래처 등록 LIST — 통장 계좌번호로 보통예금 거래처코드 지정")
    ap.add_argument("--tax", nargs="*", help="국세 납부내역증명 (홈택스 PDF) — 세목별 계정 지정")
    ap.add_argument("--payroll-month", help="급여대장에 귀속월 열이 없을 때 귀속월 (예: 2026-09)")
    ap.add_argument("--owner", nargs="*", help="대표자(사업주) 본인 이름 — 가지급금/가수금 처리")
    ap.add_argument("--start", help="이 날짜부터 분개 (예: 2026-03-22) — 이미 작업한 기간 제외")
    ap.add_argument("--end", help="이 날짜까지 분개 (예: 2026-09-30)")
    ap.add_argument("--config", help="설정 파일 (기본: config.json)")
    ap.add_argument("--client-config", help="고객별 설정(이름 별칭 등) — input/<거래처>/config.json")
    ap.add_argument("--out", default="분개결과.xlsx", help="검토용 결과 엑셀")
    ap.add_argument("--upload", help="세무사랑 일반전표 업로드 파일(.xls) 경로 (기본: <out>_세무사랑업로드.xls)")
    ap.add_argument("--company", default="", help="업로드 양식 상단 회사명")
    ap.add_argument("--company-bizno", default="", help="업로드 양식 상단 사업자등록번호")
    a = ap.parse_args()

    j = run(a.bank, a.ledger, a.sales, a.purchase, a.payroll, a.business, a.out, a.config,
            a.payroll_month, a.owner, a.withholding, a.upload, a.company, a.company_bizno, a.start, a.end,
            a.client_config, a.tax, a.fin)
    flagged = sum(1 for e in j.entries if e.review)
    print(f"분개 {len(j.entries)}건 (확인필요 {flagged}건), 신규거래처 {len(j.new_partners)}곳 → {a.out}")
    print(f"세무사랑 업로드 {j.upload_rows}줄 → {j.upload_path}")


if __name__ == "__main__":
    main()
