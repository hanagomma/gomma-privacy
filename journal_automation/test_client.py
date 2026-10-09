import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "samples"))

import client  # noqa: E402
import make_samples  # noqa: E402

S = HERE / "samples"


class ClientFlow(unittest.TestCase):
    """new → add(자동 분류) → check → run 이 가상 샘플로 돌아가는지."""

    @classmethod
    def setUpClass(cls):
        make_samples.main()
        cls.tmp = Path(tempfile.mkdtemp())
        client.INPUT = cls.tmp

    def test_flow(self):
        with redirect_stdout(StringIO()):
            client.cmd_new("테스트상회")
        base = self.tmp / "테스트상회"
        for folder, _ in client.FOLDERS.values():
            self.assertTrue((base / folder / "안내.txt").exists())
        files = ["통장_202609.xlsx", "거래처원장.xlsx", "금융거래처.xlsx", "매출세금계산서.xlsx",
                 "매입세금계산서.xlsx", "급여대장_202608.xlsx", "급여대장_202609.xlsx", "사업소득_202609.xlsx",
                 "예수금원장.xlsx"]
        out = StringIO()
        with redirect_stdout(out):
            client.cmd_add("테스트상회", [str(S / f) for f in files])
        self.assertEqual(len(client.files_in(base, "bank")), 1)
        self.assertEqual(len(client.files_in(base, "payroll")), 2)
        self.assertEqual(len(client.files_in(base, "withholding")), 1)
        self.assertEqual(len(client.files_in(base, "business")), 1)
        self.assertEqual(len(client.files_in(base, "fin")), 1)
        self.assertEqual(len(client.files_in(base, "ledger")), 1)
        self.assertEqual([p.name for p in client.files_in(base, "purchase")], ["매입세금계산서.xlsx"])

        info = json.loads((base / "작업정보.json").read_text(encoding="utf-8"))
        info.update({"시작일": "2026-09-01", "종료일": "2026-09-30"})
        (base / "작업정보.json").write_text(json.dumps(info, ensure_ascii=False), encoding="utf-8")
        with redirect_stdout(StringIO()):
            client.cmd_run("테스트상회")
        results = list((base / "결과").iterdir())
        self.assertTrue(any(p.name.endswith("_세무사랑업로드.xls") for p in results))
        self.assertTrue(any(p.suffix == ".xlsx" for p in results))

    def test_check_missing_bank(self):
        with redirect_stdout(StringIO()):
            client.cmd_new("빈거래처")
            errors, _ = client.check("빈거래처")
        self.assertTrue(any("01_통장" in e for e in errors))


if __name__ == "__main__":
    unittest.main()
