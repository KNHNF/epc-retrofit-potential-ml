"""
test_prepare_data.py
Tests for src/00_prepare_data.py on tiny synthetic CSVs (no real EPC data needed).
The module name starts with a digit, so it is loaded by path rather than imported.
"""

import importlib.util
import os
import sys

import pandas as pd
import pytest

SRC = os.path.join(os.path.dirname(__file__), '..', 'src')
sys.path.insert(0, SRC)

_spec = importlib.util.spec_from_file_location('prepare_data', os.path.join(SRC, '00_prepare_data.py'))
prep = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(prep)

from preprocessing import engineer_target


def row(uprn, date, cur_rating='D', pot_rating='B', cur_eff=55, pot_eff=80, ptype='House'):
    return {
        'UPRN': uprn, 'LODGEMENT_DATE': date,
        'CURRENT_ENERGY_RATING': cur_rating, 'POTENTIAL_ENERGY_RATING': pot_rating,
        'CURRENT_ENERGY_EFFICIENCY': cur_eff, 'POTENTIAL_ENERGY_EFFICIENCY': pot_eff,
        'PROPERTY_TYPE': ptype, 'TOTAL_FLOOR_AREA': 80.0,
    }


def write_csv(path, rows, encoding='latin-1'):
    pd.DataFrame(rows).to_csv(path, index=False, encoding=encoding)


def test_newest_certificate_wins_across_chunks(tmp_path):
    # Same UPRN in two different chunks of one file: the newer one must be kept
    # even though it comes second. chunksize=1 puts each row in its own chunk.
    f = tmp_path / 'certificates-2024.csv'
    write_csv(f, [row(1, '2024-01-10', cur_eff=50), row(1, '2024-09-10', cur_eff=60)])
    out = prep.process_year_file(str(f), set(), chunksize=1)
    assert len(out) == 1
    assert out.iloc[0]['CURRENT_ENERGY_EFFICIENCY'] == 60


def test_uprn_in_later_year_is_dropped_from_earlier_year(tmp_path):
    f = tmp_path / 'certificates-2021.csv'
    write_csv(f, [row(1, '2021-05-01'), row(2, '2021-05-02')])
    out = prep.process_year_file(str(f), {'1'})
    assert out['UPRN'].tolist() == ['2']


def test_seen_uprns_updated_after_file(tmp_path):
    f = tmp_path / 'certificates-2021.csv'
    write_csv(f, [row(1, '2021-05-01'), row(2, '2021-05-02')])
    seen = set()
    prep.process_year_file(str(f), seen)
    assert seen == {'1', '2'}


def test_float_style_uprn_is_normalised(tmp_path):
    f = tmp_path / 'certificates-2021.csv'
    write_csv(f, [row(100.0, '2021-05-01')])
    out = prep.process_year_file(str(f), set())
    assert out['UPRN'].tolist() == ['100']


def test_missing_required_column_raises_clear_error(tmp_path):
    f = tmp_path / 'certificates-2021.csv'
    df = pd.DataFrame([row(1, '2021-05-01')]).drop(columns=['POTENTIAL_ENERGY_EFFICIENCY'])
    df.to_csv(f, index=False)
    with pytest.raises(ValueError, match='POTENTIAL_ENERGY_EFFICIENCY'):
        prep.process_year_file(str(f), set())


def test_invalid_rows_are_filtered(tmp_path):
    f = tmp_path / 'certificates-2021.csv'
    write_csv(f, [
        row(1, '2021-01-01'),
        row(2, '2021-01-01', cur_rating='Z'),            # bad rating
        row(3, '2021-01-01', cur_eff=0),                 # efficiency out of range
        row(4, '2021-01-01', ptype='Caravan'),           # property type not modelled
        row(5, 'not a date'),                            # unparseable date
    ])
    out = prep.process_year_file(str(f), set())
    assert out['UPRN'].tolist() == ['1']


def test_main_raises_when_no_raw_files(tmp_path, monkeypatch):
    monkeypatch.setattr(prep, 'DATA_DIR', str(tmp_path / 'raw'))
    monkeypatch.setattr(prep, 'OUTPUT_DIR', str(tmp_path / 'out'))
    with pytest.raises(FileNotFoundError, match='certificates'):
        prep.main()


def test_main_temporal_split_has_no_overlap_and_no_time_leak(tmp_path, monkeypatch):
    raw, out = tmp_path / 'raw', tmp_path / 'out'
    raw.mkdir()
    # UPRN 1 is certified in 2022 and again in 2025: only the 2025 certificate may
    # survive, and it must land in the test set, not train.
    write_csv(raw / 'certificates-2022.csv', [row(1, '2022-03-01'), row(2, '2022-04-01')])
    write_csv(raw / 'certificates-2024.csv', [row(3, '2024-06-01')])
    write_csv(raw / 'certificates-2025.csv', [row(1, '2025-02-01', cur_eff=40, pot_eff=70),
                                              row(4, '2025-03-01')])
    write_csv(raw / 'certificates-2026.csv', [row(5, '2026-01-15')])
    monkeypatch.setattr(prep, 'DATA_DIR', str(raw))
    monkeypatch.setattr(prep, 'OUTPUT_DIR', str(out))
    prep.main()

    train = pd.read_parquet(out / 'epc_train_full.parquet')
    test = pd.read_parquet(out / 'epc_test_full.parquet')

    assert set(train['UPRN']) == {'2', '3'}
    assert set(test['UPRN']) == {'1', '4', '5'}
    assert set(train['UPRN']).isdisjoint(test['UPRN'])
    assert pd.to_datetime(train['LODGEMENT_DATE']).max() < pd.to_datetime(test['LODGEMENT_DATE']).min()
    # the 2025 certificate for UPRN 1 won, not the 2022 one
    assert test.loc[test['UPRN'] == '1', 'CURRENT_ENERGY_EFFICIENCY'].iloc[0] == 40


def test_target_matches_preprocessing_engineer_target(tmp_path, monkeypatch):
    # The target is defined twice (here and in src/preprocessing.py). Guard against drift.
    raw, out = tmp_path / 'raw', tmp_path / 'out'
    raw.mkdir()
    write_csv(raw / 'certificates-2024.csv', [
        row(1, '2024-01-01', cur_rating='D', cur_eff=55, pot_eff=75),   # gap 20, D: positive
        row(2, '2024-01-01', cur_rating='D', cur_eff=55, pot_eff=74),   # gap 19: negative
        row(3, '2024-01-01', cur_rating='C', pot_rating='A', cur_eff=70, pot_eff=95),  # C: negative
    ])
    write_csv(raw / 'certificates-2025.csv', [row(9, '2025-01-01')])
    monkeypatch.setattr(prep, 'DATA_DIR', str(raw))
    monkeypatch.setattr(prep, 'OUTPUT_DIR', str(out))
    prep.main()
    train = pd.read_parquet(out / 'epc_train_full.parquet').sort_values('UPRN').reset_index(drop=True)
    expected = engineer_target(train.drop(columns=['RETROFIT_POTENTIAL', 'EFFICIENCY_GAP']))
    assert train['RETROFIT_POTENTIAL'].tolist() == expected['RETROFIT_POTENTIAL'].tolist() == [1, 0, 0]
