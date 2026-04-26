from __future__ import annotations

import unittest
from pathlib import Path
import shutil
from uuid import uuid4

from bm2.store import DISABLED, ENABLED, ExcelStore


class StoreRegressionBehaviorTests(unittest.TestCase):
    def setUp(self) -> None:
        temp_root = Path.cwd() / '.tmp_test_workspaces'
        temp_root.mkdir(exist_ok=True)
        self.base_dir = temp_root / f'case_{uuid4().hex}'
        self.base_dir.mkdir()
        self.store = ExcelStore(self.base_dir)

    def tearDown(self) -> None:
        shutil.rmtree(self.base_dir, ignore_errors=True)

    def _member_names(self) -> list[str]:
        return [item['name'] for item in self.store.get_members()]

    def test_add_member_persists_new_member(self) -> None:
        self.store.add_member('new-member', 'note')

        members = self.store.get_members()
        added = next(item for item in members if item['name'] == 'new-member')

        self.assertEqual(ENABLED, added['status'])
        self.assertEqual('note', added['note'])
        self.assertIn('new-member', [item['name'] for item in self.store.get_active_members()])

    def test_disable_member_removes_member_from_active_list(self) -> None:
        target_name = self.store.get_active_members()[0]['name']

        self.store.update_member(target_name, 'disabled note', status=DISABLED)

        member = next(item for item in self.store.get_members() if item['name'] == target_name)
        self.assertEqual(DISABLED, member['status'])
        self.assertEqual('disabled note', member['note'])
        self.assertTrue(str(member['disabled_at']).strip())
        self.assertNotIn(target_name, [item['name'] for item in self.store.get_active_members()])

    def test_delete_member_removes_member_from_member_list(self) -> None:
        target_name = 'delete-me'
        self.store.add_member(target_name)

        self.store.delete_member(target_name)

        self.assertNotIn(target_name, self._member_names())

    def test_reorder_active_members_updates_active_member_order(self) -> None:
        active_names = [item['name'] for item in self.store.get_active_members()]
        reordered = list(reversed(active_names[:3]))

        self.store.reorder_active_members(reordered)

        updated_names = [item['name'] for item in self.store.get_active_members()]
        self.assertEqual(reordered, updated_names[:3])

    def test_default_score_date_can_be_set_and_read_back(self) -> None:
        self.store.set_default_score_date('2026-05-10')

        self.assertEqual('2026-05-10', self.store.get_default_score_date())

    def test_save_scores_rolls_date_forward_when_same_day_already_exists(self) -> None:
        member_name = self.store.get_active_members()[0]['name']
        entry = [{
            'name': member_name,
            'score': '3',
            'before_balance': '',
            'after_balance': '',
            'manual_wear': '',
            'income': '',
            'other_expense': '',
        }]

        first = self.store.save_scores_and_wear('2026-04-26', entry)
        second = self.store.save_scores_and_wear('2026-04-26', entry)

        self.assertEqual('2026-04-26', first['saved_date'])
        self.assertEqual('2026-04-27', second['saved_date'])

    def test_create_new_cycle_switches_workbook_and_updates_default_date(self) -> None:
        new_filename = self.store.create_new_cycle('2026-05-01')

        self.assertEqual(new_filename, self.store.workbook_path.name)
        self.assertEqual('2026-05-01', self.store.get_default_score_date())
        self.assertTrue(self.store.workbook_path.exists())

    def test_wear_threshold_can_be_saved_and_read_back(self) -> None:
        saved = self.store.set_wear_abnormal_threshold('3.7')

        self.assertEqual(3.7, saved)
        self.assertEqual(3.7, self.store.get_wear_abnormal_threshold())

    def test_profit_summary_aggregates_income_minus_wear_across_sheets(self) -> None:
        member_name = self.store.get_active_members()[0]['name']

        self.store.save_scores_and_wear(
            '2026-04-26',
            [{
                'name': member_name,
                'score': '6',
                'before_balance': '',
                'after_balance': '',
                'manual_wear': '1.5',
                'income': '10',
                'other_expense': '0',
            }],
        )

        profit_map = self.store.get_active_member_profit_map()
        calendar = self.store.get_member_profit_calendar('all', 2026, 4)

        self.assertEqual(8.5, profit_map[member_name])
        self.assertEqual(8.5, calendar['month_profit_total'])

    def test_workbook_can_be_created_saved_and_reopened(self) -> None:
        member_name = self.store.get_active_members()[0]['name']

        self.assertTrue(self.store.workbook_path.exists())

        self.store.save_scores_and_wear(
            '2026-04-26',
            [{
                'name': member_name,
                'score': '4',
                'before_balance': '',
                'after_balance': '',
                'manual_wear': '1.2',
                'income': '8',
                'other_expense': '2',
            }],
        )

        reopened = ExcelStore(self.base_dir)
        summary = reopened.get_score_summary()
        wear_records = reopened.get_member_wear_records(member_name)
        income_records = reopened.get_member_income_records(member_name)

        self.assertTrue(str(summary['latest_column']).strip())
        self.assertTrue(summary['rankings'])
        self.assertEqual(1.2, wear_records[-1]['wear'])
        self.assertEqual(8.0, income_records[-1]['income'])


if __name__ == '__main__':
    unittest.main()
