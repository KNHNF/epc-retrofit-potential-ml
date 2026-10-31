"""
test_data_loader.py
Tests for the CSV loaders, using small temporary files.
"""

import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from data_loader import load_epc_csv, load_epc_chunked


def make_csv(path, n=20):
    # latin-1 content: the pound sign and an accented character are not valid utf-8 bytes
    df = pd.DataFrame({'UPRN': range(n), 'WALLS_DESCRIPTION': ['Cavity wall, £ café'] * n})
    df.to_csv(path, index=False, encoding='latin-1')


def test_load_epc_csv_reads_latin1(tmp_path):
    f = tmp_path / 'c.csv'
    make_csv(f)
    df = load_epc_csv(str(f))
    assert len(df) == 20
    assert '£' in df.loc[0, 'WALLS_DESCRIPTION']


def test_load_epc_csv_sampling_is_reproducible(tmp_path):
    f = tmp_path / 'c.csv'
    make_csv(f)
    a = load_epc_csv(str(f), sample_n=5, seed=1)
    b = load_epc_csv(str(f), sample_n=5, seed=1)
    assert len(a) == 5
    assert a['UPRN'].tolist() == b['UPRN'].tolist()


def test_load_epc_csv_sample_larger_than_file_returns_all(tmp_path):
    f = tmp_path / 'c.csv'
    make_csv(f, n=5)
    assert len(load_epc_csv(str(f), sample_n=100)) == 5


def test_load_epc_chunked_matches_full_load(tmp_path):
    f = tmp_path / 'c.csv'
    make_csv(f)
    full = load_epc_csv(str(f))
    chunked = load_epc_chunked(str(f), chunksize=7)
    assert chunked['UPRN'].tolist() == full['UPRN'].tolist()
