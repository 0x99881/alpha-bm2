from __future__ import annotations

import unittest
from decimal import Decimal

from bm2.domain.rules.daily_entry import IncompleteBalanceInput, find_invalid_score_member, resolve_wear_value


def parse_decimal(value: str, field_name: str) -> Decimal:
    return Decimal(value)


class DailyEntryRuleTests(unittest.TestCase):
    def test_empty_wear_inputs_return_none(self) -> None:
        self.assertIsNone(
            resolve_wear_value(
                before_text='',
                after_text='',
                manual_wear_text='',
                parse_decimal=parse_decimal,
                manual_field='manual',
                before_field='before',
                after_field='after',
            )
        )

    def test_balance_inputs_calculate_wear(self) -> None:
        self.assertEqual(
            Decimal('2.5'),
            resolve_wear_value(
                before_text='10',
                after_text='7.5',
                manual_wear_text='',
                parse_decimal=parse_decimal,
                manual_field='manual',
                before_field='before',
                after_field='after',
            ),
        )

    def test_manual_wear_overrides_balance_inputs(self) -> None:
        self.assertEqual(
            Decimal('1.25'),
            resolve_wear_value(
                before_text='10',
                after_text='7.5',
                manual_wear_text='1.25',
                parse_decimal=parse_decimal,
                manual_field='manual',
                before_field='before',
                after_field='after',
            ),
        )

    def test_partial_balance_inputs_are_invalid(self) -> None:
        with self.assertRaises(IncompleteBalanceInput):
            resolve_wear_value(
                before_text='10',
                after_text='',
                manual_wear_text='',
                parse_decimal=parse_decimal,
                manual_field='manual',
                before_field='before',
                after_field='after',
            )

    def test_score_validation_allows_empty_and_integer_scores(self) -> None:
        self.assertIsNone(find_invalid_score_member([{'name': 'a', 'score': ''}, {'name': 'b', 'score': '18'}]))

    def test_score_validation_returns_first_invalid_member(self) -> None:
        self.assertEqual('b', find_invalid_score_member([{'name': 'a', 'score': '1'}, {'name': 'b', 'score': '1.5'}]))


if __name__ == '__main__':
    unittest.main()
