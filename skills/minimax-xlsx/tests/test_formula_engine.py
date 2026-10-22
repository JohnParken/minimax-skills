#!/usr/bin/env python3
"""Focused tests for the standard-library formula evaluator."""

import sys
import unittest
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from formula_engine import FormulaError, evaluate  # noqa: E402


class FormulaEngineTests(unittest.TestCase):
    def setUp(self):
        self.cells = {
            ("Current", "A1"): Decimal("2"),
            ("Current", "A2"): Decimal("3"),
            ("Current", "A3"): "label",
            ("Current", "A4"): True,
            ("Current", "B1"): None,
            ("Data Sheet", "B2"): Decimal("7.5"),
        }
        self.calls = []

        def resolve(sheet, address):
            self.calls.append((sheet, address))
            return self.cells.get((sheet, address))

        self.resolve = resolve

    def value(self, formula):
        return evaluate(formula, "Current", self.resolve)

    def assert_error(self, formula, code):
        with self.assertRaises(FormulaError) as caught:
            self.value(formula)
        self.assertEqual(caught.exception.code, code)
        self.assertTrue(str(caught.exception))

    def test_decimal_round_half_up_for_positive_and_negative_ties(self):
        self.assertEqual(self.value("ROUND(2.5,0)"), Decimal("3"))
        self.assertEqual(self.value("=ROUND(-2.5,0)"), Decimal("-3"))
        self.assertEqual(self.value("ROUND(1.245,2)"), Decimal("1.25"))

    def test_arithmetic_precedence_and_comparison(self):
        self.assertEqual(self.value("1 + 2 * 3 - 8 / 4"), Decimal("5"))
        self.assertTrue(self.value("(1 + 2) * 3 >= 9"))
        self.assertEqual(self.value("-$A$1 + A2"), Decimal("1"))
        self.assertEqual(self.value("TRUE + B1"), Decimal("1"))

    def test_mixed_boolean_comparison_is_not_numeric_coercion(self):
        self.assert_error("TRUE=1", "unsupported")
        self.assertTrue(self.value("TRUE=TRUE"))

    def test_if_evaluates_only_the_selected_branch(self):
        self.assertEqual(self.value("IF(TRUE, 9, 1/0)"), Decimal("9"))
        self.assertEqual(self.calls, [])
        self.assertEqual(self.value("IF(FALSE, A99, 4)"), Decimal("4"))
        self.assertEqual(self.calls, [])

    def test_cross_sheet_reference_and_quoted_sheet_name(self):
        self.assertEqual(self.value("'Data Sheet'!B2 + 0.5"), Decimal("8.0"))
        self.assertEqual(self.calls, [("Data Sheet", "B2")])
        self.calls.clear()
        self.assertEqual(self.value("SUM('Data Sheet'!B1:B2)"), Decimal("7.5"))
        self.assertEqual(self.calls, [("Data Sheet", "B1"), ("Data Sheet", "B2")])

    def test_ranges_and_aggregates(self):
        self.assertEqual(self.value("SUM(A1:A4)"), Decimal("5"))
        self.assertEqual(self.value("COUNT(A1:A4)"), Decimal("2"))
        self.assertEqual(self.value("AVERAGE(A1:A2)"), Decimal("2.5"))
        self.assertEqual(self.value("MIN(A1:A2)"), Decimal("2"))
        self.assertEqual(self.value("MAX(A1:A2)"), Decimal("3"))
        self.assertEqual(self.value("SUM('Current'!A2:A1)"), Decimal("5"))

    def test_scalar_boolean_aggregate_arguments_are_unsupported(self):
        for formula in ("COUNT(TRUE)", "SUM(A4)", "AVERAGE(FALSE)", "MIN(TRUE)", "MAX(FALSE)"):
            with self.subTest(formula=formula):
                self.assert_error(formula, "unsupported")

    def test_string_literals_are_not_rewritten_or_resolved(self):
        result = self.value('IF(TRUE, "A1", "Data Sheet!B2")')
        self.assertEqual(result, "A1")
        self.assertEqual(self.calls, [])
        self.assertEqual(self.value('"say ""hi"""'), 'say "hi"')

    def test_resolver_cycle_errors_propagate(self):
        def resolve(sheet, address):
            raise FormulaError("cycle", "Circular dependency")

        with self.assertRaises(FormulaError) as caught:
            evaluate("A1", "Current", resolve)
        self.assertEqual(caught.exception.code, "cycle")

    def test_division_by_zero_and_text_arithmetic(self):
        self.assert_error("1/0", "calculation_error")
        self.assert_error('"text" + 1', "unsupported")

    def test_unknown_functions_names_and_injection_are_rejected(self):
        self.assert_error("NO_SUCH_FUNCTION(1)", "unsupported")
        self.assert_error("NO_SUCH_FUNCTION(1/0)", "unsupported")
        self.assert_error("NamedRange + 1", "unsupported")
        self.assert_error('__import__("os")', "unsupported")
        self.assert_error("1 ^ 2", "unsupported")

    def test_malformed_formula_and_invalid_references_are_controlled(self):
        self.assert_error("(1 + 2", "calculation_error")
        self.assert_error('"unterminated', "calculation_error")
        self.assert_error("XFE1", "calculation_error")
        self.assert_error("A1048577", "calculation_error")
        self.assert_error("A" + "9" * 5000, "calculation_error")
        self.assert_error("² + 1", "calculation_error")

    def test_formula_and_range_size_limits(self):
        self.assert_error("1" + " " * 8192, "unsupported")
        self.assert_error("SUM(A1:CV1001)", "unsupported")

    def test_aggregate_cell_errors_propagate(self):
        def resolve(sheet, address):
            if address == "A2":
                raise FormulaError("calculation_error", "Cell calculation failed")
            return Decimal("1")

        with self.assertRaises(FormulaError) as caught:
            evaluate("SUM(A1:A3)", "Current", resolve)
        self.assertEqual(caught.exception.code, "calculation_error")


if __name__ == "__main__":
    unittest.main()
