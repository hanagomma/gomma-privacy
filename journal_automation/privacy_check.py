#!/usr/bin/env python3
"""고객 정보가 저장소 밖으로 나가지 않게 막는 검사.

  python privacy_check.py            # 커밋 대기(staged) 내용 검사
  python privacy_check.py --push     # 아직 푸시 안 된 커밋 전체 검사
  python privacy_check.py --hook     # Claude Code PreToolUse 훅: git commit / git push 직전에 자동 검사

막는 것
  - input/ · output/ 아래 파일 (고객 자료)
  - 주민·외국인등록번호 형식 (######-#######)
  - input/<거래처>/ 자료에서 뽑은 실제 회사명·사람 이름·사업자번호·계좌번호
문제가 있으면 종료코드 2 (훅에서는 명령이 차단된다).
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
INPUT = HERE / "input"

RRN = re.compile(r"(?<!\d)\d{6}-[1-8]\d{6}(?!\d)")          # 주민·외국인등록번호
BLOCKED_DIRS = ("journal_automation/input/", "journal_automation/output/")
# 흔한 말이라 이름으로 쓰여도 막지 않을 단어
COMMON = {"주식회사", "유한회사", "보통예금", "국민연금", "건강보험", "고용보험", "세무서", "합계", "소계", "대표자",
          "거래처", "테스트", "은행", "하나", "국민", "신한", "우리", "농협", "기업"}


def git(*args) -> str:
    return subprocess.run(["git", "-C", str(ROOT), *args], capture_output=True, text=True).stdout


def sensitive_terms() -> set[str]:
    """input/ 의 실제 자료에서 회사명·사람 이름·번호를 뽑는다 (로컬에서만 쓰고 어디에도 저장하지 않는다)."""
    terms: set[str] = set()
    if not INPUT.exists():
        return terms
    sys.path.insert(0, str(HERE))
    try:
        import auto_journal as aj
    except Exception:
        return terms
    add = lambda s: terms.add(str(s).strip()) if s and len(str(s).strip()) >= 3 else None
    for f in INPUT.rglob("*"):
        if not f.is_file():
            continue
        try:
            if f.name in ("작업정보.json", "config.json", "고객설정.json"):
                d = json.loads(f.read_text(encoding="utf-8"))
                for k in ("회사명", "사업자번호"):
                    add(d.get(k))
                for k, v in (d.get("name_aliases") or {}).items():
                    add(k)
                    [add(x) for x in v]
                continue
            if f.suffix.lower() not in (".xls", ".xlsx", ".csv"):
                continue
            kind = None
            for loader, k in ((aj.load_payroll, "pay"), (aj.load_business, "biz"), (aj.load_ledger, "ledger"),
                              (aj.load_fin_partners, "fin")):
                try:
                    rows = loader(f)
                except Exception:
                    continue
                if rows:
                    kind = k
                    for r in rows:
                        add(getattr(r, "name", ""))
                        add(getattr(r, "ceo", ""))
                        b = getattr(r, "bizno", "")
                        if b and len(b) >= 10:
                            terms.add(b)
                    break
            if kind is None and (f.parent.name.startswith("01_") or "bank" in f.name):
                acct = aj.bank_account_number(f)
                if acct:
                    terms.add(acct)
        except Exception:
            continue
    # 이미 저장소(HEAD)에 있는 단어는 일반 단어로 본다 — 저장소는 실제 자료 없이 정리된 상태를 기준으로 한다
    baseline = ""
    for f in git("-c", "core.quotepath=off", "ls-tree", "-r", "--name-only", "HEAD").split("\n"):
        p = ROOT / f
        if f and p.suffix.lower() not in (".xls", ".xlsx") and p.is_file():
            baseline += git("show", f"HEAD:{f}")
    out = set()
    for t in terms:
        if t in COMMON or re.fullmatch(r"\d{1,9}", t) or (not t.isdigit() and t in baseline):
            continue
        out.add(t)
        if re.fullmatch(r"\d{10,}", t):            # 번호는 하이픈 넣은 모양도
            out.add(f"{t[:3]}-{t[3:5]}-{t[5:]}")
    return out


def check_text(label: str, text: str, terms: set[str]) -> list[str]:
    probs = []
    if RRN.search(text):
        probs.append(f"{label}: 주민·외국인등록번호 형식")
    digits_only = re.sub(r"\D", "", text)
    for t in terms:
        if t.isdigit():
            if len(t) >= 10 and t in digits_only:
                probs.append(f"{label}: 실제 번호 ({t[:3]}…)")
        elif t in text:
            probs.append(f"{label}: 실제 이름/상호 '{t[0]}…'")
    return probs


def scan(push: bool = False) -> list[str]:
    if push:
        upstream = git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}").strip()
        rng = f"{upstream}..HEAD" if upstream else "HEAD"
        files = [f for f in git("diff", "--name-only", rng.split("..")[0], "HEAD").split("\n") if f] \
            if upstream else [f for f in git("ls-files").split("\n") if f]
        diff = git("diff", rng.split("..")[0], "HEAD") if upstream else ""
        msgs = git("log", "--format=%B", rng)
    else:
        files = [f for f in git("diff", "--cached", "--name-only").split("\n") if f]
        diff = git("diff", "--cached")
        msgs = ""
    probs = [f"{f}: 고객 자료 폴더 파일" for f in files if f.startswith(BLOCKED_DIRS)]
    terms = sensitive_terms()
    added = "\n".join(l[1:] for l in diff.split("\n") if l.startswith("+") and not l.startswith("+++"))
    probs += check_text("변경 내용", added, terms)
    probs += check_text("커밋 메시지", msgs, terms)
    # 바이너리(엑셀)는 diff 에 안 보이므로 직접 연다
    for f in files:
        p = ROOT / f
        if p.suffix.lower() in (".xls", ".xlsx") and p.exists():
            try:
                import pandas as pd
                t = " ".join(" ".join(df.fillna("").astype(str).values.ravel())
                             for df in pd.read_excel(p, sheet_name=None, header=None, dtype=str).values())
                probs += check_text(f, t, terms)
            except Exception:
                pass
    return sorted(set(probs))


OUTBOUND_SHELL = re.compile(r"\b(curl|wget|scp|rsync|sftp|ftp|nc|ncat|gh)\b")


def hook(event: dict) -> list[str]:
    """Claude Code PreToolUse 훅. 밖으로 나가는 도구 호출에 고객 정보가 있으면 문제 목록을 돌려준다."""
    tool = event.get("tool_name", "")
    inp = event.get("tool_input", {}) or {}
    if tool == "Bash":
        cmd = inp.get("command", "")
        if re.search(r"\bgit\b.*\bpush\b", cmd):
            return scan(push=True)
        if re.search(r"\bgit\b.*\bcommit\b", cmd):
            return scan()
        if OUTBOUND_SHELL.search(cmd):
            probs = []
            if re.search(r"journal_automation/(input|output)|/root/\.claude/uploads", cmd):
                probs.append("고객 자료 파일을 외부로 보내는 명령")
            return probs + check_text("명령", cmd, sensitive_terms())
        return []
    # 그 외 외부로 나가는 도구(GitHub·메일·문서·아티팩트·웹·다른 세션): 입력 전체와 첨부 파일 내용 검사
    terms = sensitive_terms()
    text = json.dumps(inp, ensure_ascii=False)
    probs = check_text(f"{tool} 입력", text, terms)
    for key in ("file_path", "path"):
        fp = inp.get(key)
        if isinstance(fp, str):
            if re.search(r"journal_automation/(input|output)", fp):
                probs.append(f"{tool}: 고객 자료 폴더 파일")
            p = Path(fp)
            if p.is_file() and p.suffix.lower() in (".html", ".md", ".txt", ".csv", ".json"):
                probs += check_text(f"{tool} 파일", p.read_text(encoding="utf-8", errors="ignore"), terms)
    return probs


def main():
    if "--hook" in sys.argv:
        try:
            event = json.load(sys.stdin)
        except Exception:
            return
        probs = hook(event)
        if probs:
            print("고객 정보가 섞여 있어 차단했습니다 (정보 보호 원칙 — CLAUDE.md):\n- " + "\n- ".join(probs[:20]),
                  file=sys.stderr)
            sys.exit(2)
        return
    probs = scan(push="--push" in sys.argv)
    if probs:
        print("문제:\n- " + "\n- ".join(probs))
        sys.exit(2)
    print("이상 없음")


if __name__ == "__main__":
    main()
