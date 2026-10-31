"""
test_preprocessing.py
Unit tests for cleaning and feature engineering functions.
Run with: pytest tests/
"""

import pandas as pd
import numpy as np
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from preprocessing import (
    drop_duplicates_by_uprn,
    filter_valid_records,
    encode_age_band,
    extract_wall_type,
    engineer_target,
    build_feature_matrix,
)


def make_sample_df():
    """Minimal valid EPC dataframe for testing."""
    return pd.DataFrame({
        'UPRN': ['A', 'A', 'B'],
        'LODGEMENT_DATE': ['2022-01-01', '2020-06-01', '2021-03-15'],
        'CURRENT_ENERGY_RATING': ['D', 'D', 'C'],
        'POTENTIAL_ENERGY_RATING': ['B', 'B', 'B'],
        'CURRENT_ENERGY_EFFICIENCY': [55, 55, 70],
        'POTENTIAL_ENERGY_EFFICIENCY': [80, 80, 80],
        'PROPERTY_TYPE': ['House', 'House', 'Flat'],
        'BUILT_FORM': ['Detached', 'Detached', 'Mid-Terrace'],
        'TENURE': ['Owner-occupied', 'Owner-occupied', 'Rental (private)'],
        'WALLS_DESCRIPTION': ['Cavity wall, as built, no insulation', 'Solid brick', 'Cavity wall, filled'],
        'CONSTRUCTION_AGE_BAND': ['England and Wales: 1930-1949', 'England and Wales: 1967-1975',
                                   'England and Wales: 2003-2006'],
        'TOTAL_FLOOR_AREA': [90.0, 90.0, 55.0],
        'NUMBER_HABITABLE_ROOMS': [4, 4, 2],
    })


def test_drop_duplicates_keeps_most_recent():
    df = make_sample_df()
    result = drop_duplicates_by_uprn(df)
    assert len(result) == 2
    # UPRN 'A' should keep the 2022 record
    a_row = result[result['UPRN'] == 'A']
    assert pd.to_datetime(a_row['LODGEMENT_DATE'].values[0]).year == 2022


def test_filter_removes_invalid_ratings():
    df = make_sample_df()
    df.loc[0, 'CURRENT_ENERGY_RATING'] = 'Z'  # invalid
    result = filter_valid_records(df)
    assert 'Z' not in result['CURRENT_ENERGY_RATING'].values


def test_filter_removes_out_of_range_efficiency():
    df = make_sample_df()
    df.loc[0, 'CURRENT_ENERGY_EFFICIENCY'] = 0  # invalid
    result = filter_valid_records(df)
    assert len(result) == 2


def test_filter_removes_negative_physical_values():
    df = make_sample_df()
    df.loc[0, 'TOTAL_FLOOR_AREA'] = -90.0  # physically impossible
    result = filter_valid_records(df)
    assert (result['TOTAL_FLOOR_AREA'] >= 0).all()
    assert len(result) == 2


def test_filter_keeps_missing_physical_values():
    # NaN is missing, not invalid: it must survive filtering (imputed downstream).
    df = make_sample_df()
    df.loc[0, 'TOTAL_FLOOR_AREA'] = np.nan
    result = filter_valid_records(df)
    assert len(result) == 3


def test_encode_age_band_no_nulls():
    df = make_sample_df()
    result = encode_age_band(df)
    assert result['AGE_BAND_YEAR'].isnull().sum() == 0


def test_extract_wall_type_cavity():
    df = make_sample_df()
    result = extract_wall_type(df)
    assert result.loc[0, 'WALL_TYPE'] == 'cavity'
    assert result.loc[1, 'WALL_TYPE'] == 'solid'


def test_engineer_target_labels():
    df = make_sample_df()
    result = engineer_target(df, gap_threshold=20)
    # Row 0: D rating, gap=25. Should be 1.
    assert result.loc[0, 'RETROFIT_POTENTIAL'] == 1
    # Row 2: C rating. Should be 0 even though gap is 10.
    assert result.loc[2, 'RETROFIT_POTENTIAL'] == 0


def test_engineer_target_gap_column():
    df = make_sample_df()
    result = engineer_target(df)
    assert 'EFFICIENCY_GAP' in result.columns
    assert result.loc[0, 'EFFICIENCY_GAP'] == 25


def test_build_feature_matrix_shape():
    df = make_sample_df()
    df = encode_age_band(df)
    df = extract_wall_type(df)
    df = engineer_target(df)
    X, y = build_feature_matrix(df)
    assert X.shape[0] == len(df)
    assert len(y) == len(df)
    assert X.isnull().sum().sum() == 0


def test_drop_duplicates_keeps_all_rows_with_missing_uprn():
    # Missing UPRN cannot be matched to a property, so these must not collapse to one row.
    df = make_sample_df()
    df['UPRN'] = [None, None, 'B']
    result = drop_duplicates_by_uprn(df)
    assert len(result) == 3


def test_filter_missing_column_gives_clear_error():
    df = make_sample_df().drop(columns=['PROPERTY_TYPE'])
    with pytest.raises(ValueError, match='PROPERTY_TYPE'):
        filter_valid_records(df)


def test_engineer_target_missing_column_gives_clear_error():
    df = make_sample_df().drop(columns=['POTENTIAL_ENERGY_EFFICIENCY'])
    with pytest.raises(ValueError, match='POTENTIAL_ENERGY_EFFICIENCY'):
        engineer_target(df)


def test_engineer_target_threshold_is_inclusive_and_d_to_g_only():
    df = make_sample_df()
    df['CURRENT_ENERGY_EFFICIENCY'] = [55, 55, 70]
    df['POTENTIAL_ENERGY_EFFICIENCY'] = [75, 74, 100]   # gaps 20, 19, 30
    result = engineer_target(df, gap_threshold=20)
    assert result['RETROFIT_POTENTIAL'].tolist() == [1, 0, 0]   # gap 20 counts, C never does


def test_extract_wall_type_missing_column_gives_clear_error():
    with pytest.raises(ValueError, match='WALLS_DESCRIPTION'):
        extract_wall_type(make_sample_df().drop(columns=['WALLS_DESCRIPTION']))


def test_encode_age_band_unmapped_values_get_median():
    df = make_sample_df()
    df['CONSTRUCTION_AGE_BAND'] = ['England and Wales: 1930-1949', 'NO DATA!',
                                   'England and Wales: 2003-2006']
    result = encode_age_band(df)
    assert result.loc[1, 'AGE_BAND_YEAR'] == (1940 + 2005) / 2
