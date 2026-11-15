###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

"""
Tests for list data validation encoding, in particular option literals that
contain a comma and the documented manual doubled-quote convention. These
cases exercise both the length-limit gate in data_validation() and the
<write_formula_1> output path, and they are repeated for the normal and
constant_memory workbook modes.
"""

import re
import unittest
from io import BytesIO, StringIO
from zipfile import ZipFile

from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


def _sheet_xml(xlsx_bytes):
    # Read the first worksheet XML out of an in-memory xlsx file.
    with ZipFile(BytesIO(xlsx_bytes)) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _formula1(xlsx_bytes):
    # Extract the data validation formula1 strings from an xlsx file.
    text = _sheet_xml(xlsx_bytes)
    return re.findall(r"<formula1>(.*?)</formula1>", text)


class TestEncodeListFormula(unittest.TestCase):
    """
    Unit tests for the Worksheet list encoding helpers.

    """

    def setUp(self):
        self.worksheet = Worksheet()

    def test_plain_items_unchanged(self):
        # Options without a comma keep the historical, unquoted encoding.
        self.assertEqual(
            self.worksheet._encode_list_formula(["Foo", "Bar", "Baz"]),
            '"Foo,Bar,Baz"',
        )

    def test_comma_item_is_quoted_as_one_option(self):
        # A literal comma no longer splits the option. This is the regression
        # that previously flattened two options into three silently.
        self.assertEqual(
            self.worksheet._encode_list_formula(["a,b", "c"]),
            '"""a,b"",c"',
        )

    def test_single_option_entirely_a_comma_list(self):
        # One option that contains commas is wrapped, not split into several.
        self.assertEqual(
            self.worksheet._encode_list_formula(["a,b,c"]),
            '"""a,b,c"""',
        )

    def test_comma_in_middle_option(self):
        self.assertEqual(
            self.worksheet._encode_list_formula(["x", "p,q", "r"]),
            '"x,""p,q"",r"',
        )

    def test_every_option_contains_a_comma(self):
        self.assertEqual(
            self.worksheet._encode_list_formula(["a,b", "c,d", "e,f"]),
            '"""a,b"",""c,d"",""e,f"""',
        )

    def test_option_is_a_single_comma(self):
        # An option that is just a literal comma still survives as one option.
        self.assertEqual(
            self.worksheet._encode_list_formula([",", "x"]),
            '""","",x"',
        )

    def test_quote_inside_a_comma_option_is_doubled(self):
        # When an option is wrapped for a comma its inner double quotes are
        # escaped by doubling them (Excel CSV rule).
        self.assertEqual(
            self.worksheet._encode_list_formula(['a"b,c', "z"]),
            '"""a""""b,c"",z"',
        )

    def test_manual_doubled_quote_convention_is_preserved(self):
        # The documented convention: callers pass '""Hello""' for the "Hello"
        # literal. Options without a comma are not rewritten at the item level;
        # the outer formula string then doubles every quote, exactly as Excel
        # does when saving the same source.
        self.assertEqual(
            self.worksheet._encode_list_formula(["x", '""Hello""']),
            '"x,""""Hello"""""',
        )

    def test_manual_doubled_quote_with_comma_option(self):
        # Manual doubling survives the automatic item wrapping applied for a
        # comma, and the outer formula layer doubles every quote once more.
        self.assertEqual(
            self.worksheet._encode_list_formula(["a,b", '""Say""', "z"]),
            '"""a,b"",""""Say"""",z"',
        )

    def test_non_string_items_are_stringified(self):
        self.assertEqual(
            self.worksheet._encode_list_formula([1, 2.5, "three"]),
            '"1,2.5,three"',
        )

    def test_single_empty_item(self):
        self.assertEqual(self.worksheet._encode_list_formula([""]), '""')

    def test_items_helper_excludes_outer_formula_quotes(self):
        # The item-level helper is what the 255 character gate measures.
        self.assertEqual(
            self.worksheet._encode_list_items(["a,b", "c"]),
            '"a,b",c',
        )


class TestWriteDataValidationListEncoding(unittest.TestCase):
    """
    End-to-end XML output tests for list data validation options.

    """

    def _write_and_get_xml(self, value):
        worksheet = Worksheet()
        fh = StringIO()
        worksheet._set_filehandle(fh)
        worksheet.data_validation(
            "B5",
            {
                "validate": "list",
                "source": value,
            },
        )
        worksheet._write_data_validations()
        return fh.getvalue()

    def test_write_plain_list(self):
        got = self._write_and_get_xml(["a", "bb", "ccc"])
        exp = (
            '<dataValidations count="1"><dataValidation type="list" '
            'allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            'sqref="B5"><formula1>"a,bb,ccc"</formula1>'
            "</dataValidation></dataValidations>"
        )
        self.assertEqual(_xml_to_list(exp), _xml_to_list(got))

    def test_write_list_with_comma_option(self):
        got = self._write_and_get_xml(["a,b", "c"])
        exp = (
            '<dataValidations count="1"><dataValidation type="list" '
            'allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            'sqref="B5"><formula1>"""a,b"",c"</formula1>'
            "</dataValidation></dataValidations>"
        )
        self.assertEqual(_xml_to_list(exp), _xml_to_list(got))

    def test_write_list_with_manual_quotes(self):
        got = self._write_and_get_xml(["x", '""Hello""'])
        exp = (
            '<dataValidations count="1"><dataValidation type="list" '
            'allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            'sqref="B5"><formula1>"x,""""Hello"""""</formula1>'
            "</dataValidation></dataValidations>"
        )
        self.assertEqual(_xml_to_list(exp), _xml_to_list(got))


class TestListLengthGate(unittest.TestCase):
    """
    Tests for the 255 character limit gate next to the validation entry point.

    """

    def setUp(self):
        self.worksheet = Worksheet()

    def _call(self, value):
        return self.worksheet.data_validation(
            "A1", {"validate": "list", "value": value}
        )

    def test_at_the_limit_is_accepted(self):
        # Exactly 255 joined characters (253 letters + 2 separators).
        value = ["a" * 100, "b" * 100, "c" * 53]
        self.assertEqual(len(",".join(value)), 255)
        self.assertEqual(self._call(value), 0)

    def test_over_the_limit_is_rejected_with_a_warning(self):
        # 256 joined characters is one past the documented limit.
        value = ["a" * 100, "b" * 100, "c" * 54]
        self.assertEqual(len(",".join(value)), 256)
        with self.assertWarns(UserWarning):
            result = self._call(value)
        self.assertEqual(result, -2)

    def test_comma_quotes_count_toward_the_limit(self):
        # The quoting characters added to protect embedded commas are part of
        # the stored item string and therefore count toward the same limit.
        value = ["a" * 126 + "," + "b" * 126]
        # One option: 252 literals + one comma + two wrapping quotes = 255.
        encoded = Worksheet()._encode_list_items(value)
        self.assertEqual(len(encoded), 255)
        self.assertEqual(self._call(value), 0)

        too_long = ["a" * 127 + "," + "b" * 127]
        self.assertEqual(len(Worksheet()._encode_list_items(too_long)), 257)
        with self.assertWarns(UserWarning):
            result = self._call(too_long)
        self.assertEqual(result, -2)

    def test_utf16_units_are_used_for_the_limit(self):
        # Excel's 255 limit is measured in UTF-16 units; an emoji counts as 2.
        # 127 emojis (254 units) plus a separator fits, 128 emojis does not.
        self.assertEqual(self._call(["\U0001f600" * 127]), 0)
        with self.assertWarns(UserWarning):
            result = self._call(["\U0001f600" * 128])
        self.assertEqual(result, -2)

    def test_formula_range_bypasses_the_list_gate(self):
        # A string source referencing a range is not a literal list and must
        # never be treated as a length-limited inline list.
        worksheet = Worksheet()
        result = worksheet.data_validation(
            "A1", {"validate": "list", "value": "=$D$1:$D$500"}
        )
        self.assertEqual(result, 0)


class TestBothExportModes(unittest.TestCase):
    """
    The same option intents must produce identical formula1 output in both the
    normal and constant_memory (in-line string) workbook modes.

    """

    def _build(self, constant_memory, value):
        buffer = BytesIO()
        workbook = Workbook(buffer, {"constant_memory": constant_memory})
        worksheet = workbook.add_worksheet()
        # In constant_memory mode rows must be written in order; writing a cell
        # before adding the validation keeps both modes on equal footing.
        worksheet.write("A1", "probe")
        worksheet.data_validation("B1", {"validate": "list", "value": value})
        workbook.close()
        return buffer.getvalue()

    def test_comma_option_identical_across_modes(self):
        value = ["a,b", "c"]
        normal = _formula1(self._build(False, value))
        optimized = _formula1(self._build(True, value))
        self.assertEqual(normal, ['"""a,b"",c"'])
        self.assertEqual(normal, optimized)

    def test_manual_quotes_identical_across_modes(self):
        value = ["x", '""Hello""']
        normal = _formula1(self._build(False, value))
        optimized = _formula1(self._build(True, value))
        self.assertEqual(normal, ['"x,""""Hello"""""'])
        self.assertEqual(normal, optimized)

    def test_mixed_options_identical_across_modes(self):
        value = ["plain", "p,q", '""Say""', "z"]
        normal = _formula1(self._build(False, value))
        optimized = _formula1(self._build(True, value))
        expected = ['"plain,""p,q"",""""Say"""",z"']
        self.assertEqual(normal, expected)
        self.assertEqual(normal, optimized)


if __name__ == "__main__":
    unittest.main()
