import ast
import re
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "rapport_FINAL_v25.py"
SOURCE = SCRIPT.read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)
FUNCTIONS = {
    node.name: node
    for node in TREE.body
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
}
ASSIGNMENTS = {
    target.id: node.value
    for node in TREE.body
    if isinstance(node, ast.Assign)
    for target in node.targets
    if isinstance(target, ast.Name)
}
NAMESPACE = {"re": re}
for name in ("BELOB_REGEX",):
    NAMESPACE[name] = ast.literal_eval(ASSIGNMENTS[name])
for name in (
    "dansk_til_float_str",
    "is_credit_nota",
    "extract_pdf_moms",
    "extract_total_ialt_from_pdf",
    "floats_equal",
):
    exec(compile(ast.Module(body=[FUNCTIONS[name]], type_ignores=[]), str(SCRIPT), "exec"), NAMESPACE)


class FakePage:
    def __init__(self, text):
        self.text = text

    def extract_text(self):
        return self.text


class FakePdf:
    def __init__(self, text):
        self.pages = [FakePage(text)]


class CreditNoteVatTests(unittest.TestCase):
    def extract_vat(self, text):
        return NAMESPACE["extract_pdf_moms"](text)

    def test_positive_and_negative_credit_note_vat_match_positive_excel(self):
        for amount in ("1.234,56", "-1.234,56"):
            with self.subTest(amount=amount):
                pdf_vat = self.extract_vat(f"KREDITNOTA\nHERAF MOMS {amount}")
                self.assertEqual(pdf_vat, "1234.56")
                self.assertTrue(NAMESPACE["floats_equal"](pdf_vat, "1234.56"))

    def test_all_supported_vat_layouts_and_danish_format(self):
        cases = (
            ("kReDiTnOtA\nSalgsmoms 25,00% af 1.234,56 308,64", "308.64"),
            ("Kreditnota\n(HERAF MOMS -4.560,00)", "4560.00"),
            ("Kreditnota\nSalgsmoms -123,45", "123.45"),
            ("Kreditnota\nMoms\n\n-1.234,56", "1234.56"),
        )
        for text, expected in cases:
            with self.subTest(text=text):
                self.assertEqual(self.extract_vat(text), expected)

    def test_zero_and_missing_credit_note_vat(self):
        self.assertEqual(self.extract_vat("Kreditnota\nMoms 0,00"), "0.00")
        self.assertEqual(self.extract_vat("Kreditnota\nIngen moms"), "")

    def test_real_vat_difference_is_still_detected(self):
        pdf_vat = self.extract_vat("Kreditnota\nMoms -100,00")
        self.assertFalse(NAMESPACE["floats_equal"](pdf_vat, "101.00"))

    def test_regular_invoice_keeps_vat_sign(self):
        self.assertEqual(self.extract_vat("Faktura\nMoms -100,00"), "-100.00")

    def test_credit_note_total_remains_negative(self):
        total = NAMESPACE["extract_total_ialt_from_pdf"](
            FakePdf("Kreditnota\nI alt 1.234,56")
        )
        self.assertEqual(total, "-1234.56")


if __name__ == "__main__":
    unittest.main()
