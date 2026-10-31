"""
test_offline_demo.py
Runs examples/offline_demo.py (synthetic data, a few seconds) so the README's offline
example cannot silently rot.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'examples'))
from offline_demo import run_demo, make_synthetic_epc


def test_demo_structure_checks_pass():
    checks = run_demo(n=800, seed=1)
    assert checks['train_ends_before_test_starts']
    assert checks['no_uprn_in_both_sets']
    assert checks['outer_folds_scored'] == 3
    assert checks['train_rows'] > 0 and checks['test_rows'] > 0


def test_demo_cleaning_removes_planted_faults():
    raw = make_synthetic_epc(300, seed=2)
    checks = run_demo(n=300, seed=2)
    # bad rating row, negative floor area row and the older duplicate are all removed
    assert checks['rows_dropped_by_cleaning_and_dedup'] == 3
    assert checks['missing_floor_area_row_kept']
    assert len(raw) == 301
