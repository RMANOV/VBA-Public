"""stdlib static/reference regressions. No native M/VBA evaluation.

Run: python -m unittest discover -s tests -v
Baseline: VBA_PQ_REF=c287c08fb640bec4d085b9a29bf07b22f05b89bf python -m unittest discover -s tests -v
Grammar: use an independent official M parser. Real Excel/Power BI remains
necessary for native evaluation, culture and external workbook integrations.
"""
import re
import unittest
from pq_reference import currency_beta, currency_eu, empty_schema, paginate, source


FULL_A = list(range(1, 101))
FULL_B = list(range(101, 201))
CASES = {
    'offset_empty': ('offset', {0: []}, 0, dict(rows=[], requests=[0], transitions=0)),
    'offset_first_short': ('offset', {0: [1, 2]}, 0, dict(rows=[1, 2], requests=[0], transitions=0)),
    'offset_full_short': ('offset', {0: FULL_A, 100: [101, 102]}, 0, dict(rows=list(range(1, 103)), requests=[0, 100], transitions=1)),
    'offset_exact_multiple_empty': ('offset', {0: FULL_A, 100: FULL_B, 200: []}, 0, dict(rows=list(range(1, 201)), requests=[0, 100, 200], transitions=2)),
}
for variant, first in [('cursor', None), ('link', 'base')]:
    CASES[variant + '_empty_final'] = (variant, {first: ([], None)}, first, dict(rows=[], requests=[first], transitions=0))
    CASES[variant + '_first_nonempty_null_next'] = (variant, {first: ([1, 2], None)}, first, dict(rows=[1, 2], requests=[first], transitions=0))
    CASES[variant + '_two_pages_final_null'] = (variant, {first: ([1, 2], 'next'), 'next': ([3], None)}, first, dict(rows=[1, 2, 3], requests=[first, 'next'], transitions=1))
    CASES[variant + '_empty_with_continuation'] = (variant, {first: ([], 'next'), 'next': ([3], None)}, first, dict(rows=[3], requests=[first, 'next'], transitions=1))


class ReferenceContracts(unittest.TestCase):
    def test_empty_typed_result_contract(self):
        self.assertEqual(empty_schema(), dict(rows=[], columns=['id', 'name', 'created_at'], missing=[]))

    def test_currency_beta_default_dot(self):
        self.assertEqual(currency_beta(), 850.0)

    def test_currency_explicit_eu_normalization(self):
        self.assertEqual(currency_eu(), 850.0)


def make_case(case):
    def test(self):
        kind, pages, initial, expected = case
        actual, internal = paginate(kind, pages, initial)
        self.assertEqual(actual, expected)
        self.assertLessEqual(internal, len(pages) + 1)
    return test


for name, case in CASES.items():
    setattr(ReferenceContracts, 'test_' + name, make_case(case))


class StaticContracts(unittest.TestCase):
    def test_null_checks_use_supported_grammar(self):
        text = source('05_NULL_HANDLING.pq')
        executable = '\n'.join(x for x in text.splitlines() if not x.lstrip().startswith('//'))
        self.assertNotRegex(executable, r'\bis\s+not\b')

    def test_missing_column_alternative_uses_official_null_behavior(self):
        text = source('14_MISSING_COLUMN_RESILIENCE.pq')
        match = re.search(r'Pattern2_SelectWithDefault\s*=\s*Table.SelectColumns\((.*?)\n\s*\)', text, re.S)
        self.assertIsNotNone(match)
        self.assertIn('MissingField.UseNull', match.group(1))

    def test_normalized_dot_currency_has_explicit_culture(self):
        self.assertIn('Number.From(normalized, "en-US")', source('06_CUSTOM_M_FUNCTIONS.pq'))

    def test_nonempty_pagination_keeps_extra_response_fields(self):
        self.assertRegex(source('16_API_PAGINATION.pq'), r'else\s+Table.FromRecords\(CombinedRecords\)')


if __name__ == '__main__':
    unittest.main()
