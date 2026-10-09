#!/usr/bin/env python3
"""거래처별 자료 폴더로 통장 분개를 돌리는 도구.

  python client.py new   <거래처명>   # input/<거래처명>/ 폴더를 자료 틀(client_template)로 만든다
  python client.py check <거래처명>   # 넣은 자료를 점검하고 빠진 것을 알려준다
  python client.py add   <거래처명> <파일...>  # 파일 내용을 보고 알맞은 번호 폴더로 자동 분류
  python client.py run   <거래처명>   # 폴더 안 자료를 자동으로 찾아 분개 → input/<거래처명>/결과/

자료를 어느 폴더에 넣는지는 client_template/자료안내.md 와 각 폴더의 안내.txt 참고.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

import auto_journal as aj

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "client_template"
INPUT = HERE / "input"

DATA_EXT = {".xls", ".xlsx", ".csv", ".pdf"}
FOLDERS = {
    "bank": ("01_통장", "필수"),
    "ledger": ("02_일반거래처", "권장"),
    "fin": ("03_금융거래처", "권장"),
    "sales": ("04_매출", "있으면"),
    "purchase": ("05_매입", "권장"),
    "payroll": ("06_급여대장", "있으면"),
    "business": ("07_사업소득", "있으면"),
    "withholding": ("08_예수금원장", "있으면"),
    "tax": ("09_세금납부", "있으면"),
}


def client_dir(name: str) -> Path:
    return INPUT / name


def files_in(base: Path, key: str) -> list[Path]:
    folder = base / FOLDERS[key][0]
    if not folder.exists():
        return []
    fs = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in DATA_EXT and not p.name.startswith(".")]
    # 예수금 원장은 '나중에 넣은 파일 우선' → 수정 시각 순, 나머지는 이름 순
    return sorted(fs, key=(lambda p: p.stat().st_mtime) if key == "withholding" else (lambda p: p.name))


def load_info(base: Path) -> dict:
    f = base / "작업정보.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}


def pdf_has_text(p: Path) -> bool:
    try:
        out = subprocess.run(["pdftotext", "-l", "1", str(p), "-"], capture_output=True, text=True).stdout
        return len(out.strip()) > 20
    except OSError:
        return True


# ---------------------------------------------------------------- new

def cmd_new(name: str):
    dest = client_dir(name)
    if dest.exists():
        sys.exit(f"이미 있습니다: {dest}")
    shutil.copytree(TEMPLATE, dest)
    info = load_info(dest)
    info["회사명"] = name
    (dest / "작업정보.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"만들었습니다: {dest}")
    print((dest / "자료안내.md").read_text(encoding="utf-8").split("## 적용되는")[0])


# ---------------------------------------------------------------- add (자동 분류)

def _peek_text(p: Path) -> str:
    """파일 앞부분 글자 (분류용)."""
    if p.suffix.lower() == ".pdf":
        try:
            t = subprocess.run(["pdftotext", "-l", "2", str(p), "-"], capture_output=True, text=True).stdout
        except OSError:
            t = ""
        if len(t.strip()) < 20:
            try:
                from pypdf import PdfReader
                t = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages[:2])
            except Exception:
                t = ""
        return t
    try:
        raw = aj._read_raw(p).head(40).fillna("")
        return "\n".join(" ".join(str(v) for v in r) for r in raw.values.tolist())
    except Exception:
        return ""


def classify(p: Path) -> str | None:
    t = _peek_text(p)
    flat = t.replace(" ", "")
    if "금융기관거래처등록" in flat:
        return "fin"
    if "일반거래처등록" in flat:
        return "ledger"
    if "거래처명" in flat and "대표자" in flat and "사업자" in flat and "공급가액" not in flat:
        return "ledger"            # 제목 없는 거래처 목록
    if "급상여대장" in flat or "급여대장" in flat:
        return "payroll"
    if "소득자명" in flat and "차인지급액" in flat:
        return "business"
    paid = any(k in flat for k in ("차인지급액", "실지급액", "실수령액"))
    if paid and any(k in flat for k in ("국민연금", "건강보험", "고용보험")):
        return "payroll"           # 제목 없는 급여대장 (머리글로 판단)
    if paid and "소득세" in flat and "성명" in flat:
        return "business"          # 제목 없는 사업소득 지급내역
    if "거래처원장" in flat and "예수금" in flat:
        return "withholding"
    if "전기이월" in flat and "차변" in flat and "대변" in flat and \
            any(k in flat for k in ("국민연금", "건강", "장기요양", "고용", "소득세", "예수금")):
        return "withholding"      # 제목 없이 내려받은 예수금 거래처원장 엑셀
    if "납부내역증명" in flat or "과세증명서" in flat or "납부확인서" in flat:
        return "tax"
    if "구분" in flat and ("매출" in flat or "매입" in flat) and "공급가액" in flat:
        n_s, n_p = flat.count("매출"), flat.count("매입")
        return "sales" if n_s > n_p else "purchase"
    if "공급자사업자등록번호" in flat:
        return "sales" if "매출" in p.name else "purchase" if "매입" in p.name else None
    if ("입금" in flat or "맡기신" in flat) and ("출금" in flat or "찾으신" in flat) and "잔액" in flat:
        return "bank"
    if p.suffix.lower() == ".pdf" and len(t.strip()) < 20:
        return None   # 이미지 PDF — 내용을 사람이(Claude가) 보고 넣는다
    return None


def cmd_add(name: str, paths: list[str]):
    base = client_dir(name)
    if not base.exists():
        cmd_new(name)
    for src in map(Path, paths):
        key = classify(src)
        if key is None:
            print(f"  ? {src.name}: 종류를 알 수 없습니다 — 맞는 번호 폴더에 직접 넣어 주세요.")
            continue
        dest = base / FOLDERS[key][0] / src.name
        shutil.copy2(src, dest)
        dest.touch()   # 예수금 원장 '나중 파일 우선'을 위해 넣은 시각으로
        print(f"  → {FOLDERS[key][0]}/{src.name}")


# ---------------------------------------------------------------- check

def check(name: str) -> tuple[list[str], list[str]]:
    """(문제 — 실행 불가, 안내 — 실행은 되지만 확인할 것)"""
    base = client_dir(name)
    if not base.exists():
        return [f"폴더가 없습니다 — python client.py new {name}"], []
    errors, notes = [], []
    info = load_info(base)
    rows = []
    for key, (folder, need) in FOLDERS.items():
        fs = files_in(base, key)
        rows.append((folder, need, len(fs), ", ".join(p.name for p in fs)))
        if not fs and need == "필수":
            errors.append(f"{folder}: {need} 자료가 없습니다.")
        elif not fs and need == "권장":
            notes.append(f"{folder}: 없음 — 거래처코드 연결이 줄어듭니다.")
    print(f"\n[{name}] 자료 현황")
    for folder, need, n, names in rows:
        print(f"  {folder:<10} {need:<4} {n}개  {names}")

    if not info.get("시작일"):
        notes.append("작업정보.json 시작일이 비어 있습니다 — 통장 처음부터 분개합니다 (이미 작업한 기간이 있으면 꼭 적기).")
    for p in files_in(base, "tax"):
        if p.suffix.lower() == ".pdf" and not pdf_has_text(p):
            csvs = [c for c in files_in(base, "tax") if c.suffix.lower() in (".csv", ".xlsx", ".xls")]
            notes.append(f"09_세금납부/{p.name}: 이미지 PDF라 자동으로 못 읽습니다"
                         + (" (CSV/엑셀이 같이 있어 그걸 씁니다)" if csvs else
                            " — 세목,납부일,합계,비고 CSV로 옮겨 적어 같은 폴더에 두세요."))
    fins = [aj.load_fin_partners(p) for p in files_in(base, "fin")]
    fins = [f for fs in fins for f in fs]
    for b in files_in(base, "bank"):
        acct = aj.bank_account_number(b)
        own = next((f for f in fins if acct and f.bizno == acct), None)
        if fins and not own:
            notes.append(f"01_통장/{b.name}: 계좌번호 {acct or '(못 찾음)'} 가 금융거래처 목록에 없습니다 — 보통예금 거래처코드 공란.")
        elif own:
            print(f"  통장 {b.name} → 보통예금 거래처 {own.code} {own.name}")
    if files_in(base, "payroll") and info.get("시작일"):
        months = {r.month for p in files_in(base, "payroll") for r in aj.load_payroll(p)}
        start = aj.to_date(info["시작일"])
        prev = aj.prev_month(start.strftime("%Y-%m"))
        if prev not in months:
            notes.append(f"06_급여대장: 시작일 직전월({prev}) 급여대장이 없습니다 — 4대보험·원천세 대조용으로 있으면 좋습니다.")
    for e in errors:
        print("  ✖", e)
    for n in notes:
        print("  ·", n)
    return errors, notes


# ---------------------------------------------------------------- run

def cmd_run(name: str):
    errors, _ = check(name)
    if errors:
        sys.exit("필수 자료가 없어 실행하지 않습니다.")
    base = client_dir(name)
    info = load_info(base)
    out_dir = base / "결과"
    out_dir.mkdir(exist_ok=True)
    f = lambda k: [str(p) for p in files_in(base, k)]
    first = lambda k: (f(k) or [None])[0]
    client_cfg = base / "고객설정.json"
    start, end = info.get("시작일") or None, info.get("종료일") or None
    period = f"{(start or '처음').replace('-', '')}-{(end or '끝').replace('-', '')}"
    for bank in files_in(base, "bank"):
        acct = aj.bank_account_number(bank)
        tag = acct[-5:] if acct else bank.stem
        out = out_dir / f"{info.get('회사명') or name}_통장분개_{tag}_{period}.xlsx"
        j = aj.run(str(bank), ledger=first("ledger"), sales=f("sales"), purchase=f("purchase"),
                   payroll=f("payroll"), business=f("business"), out=str(out),
                   owners=info.get("대표자(본인 이름 입출금)") or None, withholding=f("withholding"),
                   company=info.get("회사명", ""), company_bizno=info.get("사업자번호", ""),
                   start=start, end=end, client_config=str(client_cfg) if client_cfg.exists() else None,
                   tax_payments=f("tax"), fin_partners=first("fin"))
        flagged = [e for e in j.entries if e.review]
        print(f"\n[{bank.name}] 전표 {len(j.entries)}건, 노란색(확인필요) {len(flagged)}건, "
              f"체크카드 제외 {len(j.excluded)}건")
        reasons = Counter(_reason(e.review) for e in flagged)
        for r, n in reasons.most_common():
            print(f"  - {r}: {n}건")
        print(f"  검토용: {out}")
        print(f"  업로드: {j.upload_path}")


def _reason(review: str) -> str:
    for key, label in (("연결 안 됨", "거래처 미연결"), ("매입 세금계산서 합계", "매입 세금계산서보다 더 지급"),
                       ("차인지급액", "급여·사업소득보다 더 지급"), ("양쪽", "급여·사업소득 중복 인원"),
                       ("세목", "세금 계정 미지정"), ("예수", "세금·예수금 대조 안 됨"),
                       ("기본 규칙", "기본 규칙 계정 확인"), ("사용자 확인", "사용자 확인 예정"),
                       ("돌려받은", "반환 거래 확인"), ("세금계산서 없음", "거래처는 찾았으나 세금계산서 없음")):
        if key in review:
            return label
    return review[:30]


def main():
    ap = argparse.ArgumentParser(description="거래처별 통장 분개 도구")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for c in ("new", "check", "run", "add"):
        sp = sub.add_parser(c)
        sp.add_argument("name", help="거래처명 (input/ 아래 폴더 이름)")
        if c == "add":
            sp.add_argument("files", nargs="+")
    a = ap.parse_args()
    if a.cmd == "add":
        cmd_add(a.name, a.files)
    else:
        {"new": cmd_new, "check": lambda n: check(n), "run": cmd_run}[a.cmd](a.name)


if __name__ == "__main__":
    main()
