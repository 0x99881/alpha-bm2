from __future__ import annotations

import unittest

from scripts.architecture_check import run_checks


class ArchitectureRuleTests(unittest.TestCase):
    def test_architecture_rules_pass(self) -> None:
        self.assertEqual([], run_checks())


if __name__ == '__main__':
    unittest.main()
