"""
preprocessing.py
Cleaning and feature engineering for UK EPC domestic data.
"""

import pandas as pd
import numpy as np


VALID_ENERGY_RATINGS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
VALID_PROPERTY_TYPES = ['House', 'Flat', 'Maisonette', 'Bungalow', 'Park home']

AGE_BAND_ORDER = {
    'England and Wales: before 1900': 1880,
    'England and Wales: 1900-1929': 1915,
    'England and Wales: 1930-1949': 1940,
    'England and Wales: 1950-1966': 1958,
    'England and Wales: 1967-1975': 1971,
    'England and Wales: 1976-1982': 1979,
    'England and Wales: 1983-1990': 1987,
    'England and Wales: 1991-1995': 1993,
    'England and Wales: 1996-2002': 1999,
    'England and Wales: 2003-2006': 2005,
    'England and Wales: 2007-2011': 2009,
    'England and Wales: 2012 onwards': 2015,
}


def load_sample(filepath: str, n: int = 200_000, seed: int = 42) -> pd.DataFrame:
    """Load and randomly sample n rows from an EPC CSV file."""
    df = pd.read_csv(filepath, encoding='latin-1', low_memory=False)
    if len(df) > n:
        df = df.sample(n=n, random_state=seed)
    return df


def drop_duplicates_by_uprn(df: pd.DataFrame) -> pd.DataFrame:
    """Keep most recent certificate per UPRN (unique property reference number)."""
    if 'UPRN' not in df.columns or 'LODGEMENT_DATE' not in df.columns:
        return df
    df = df.copy()
    df['LODGEMENT_DATE'] = pd.to_datetime(df['LODGEMENT_DATE'], errors='coerce')
    df = df.sort_values('LODGEMENT_DATE', ascending=False)
    df = df.drop_duplicates(subset='UPRN', keep='first')
    return df


def filter_valid_records(df: pd.DataFrame) -> pd.DataFrame:
    """Remove records with invalid energy ratings or out-of-range efficiency scores."""
    df = df.copy()
    df = df[df['CURRENT_ENERGY_RATING'].isin(VALID_ENERGY_RATINGS)]
    df = df[df['POTENTIAL_ENERGY_RATING'].isin(VALID_ENERGY_RATINGS)]
    df = df[df['PROPERTY_TYPE'].isin(VALID_PROPERTY_TYPES)]
    # Efficiency scores must be in range 1-100
    for col in ['CURRENT_ENERGY_EFFICIENCY', 'POTENTIAL_ENERGY_EFFICIENCY']:
        if col in df.columns:
            df = df[(df[col] >= 1) & (df[col] <= 100)]
    return df


def encode_age_band(df: pd.DataFrame) -> pd.DataFrame:
    """Map CONSTRUCTION_AGE_BAND to approximate midpoint year (ordinal numeric)."""
    df = df.copy()
    df['AGE_BAND_YEAR'] = df['CONSTRUCTION_AGE_BAND'].map(AGE_BAND_ORDER)
    # Fill unmapped values (unknown, NO DATA!, etc.) with median
    median_year = df['AGE_BAND_YEAR'].median()
    df['AGE_BAND_YEAR'] = df['AGE_BAND_YEAR'].fillna(median_year)
    return df


def extract_wall_type(df: pd.DataFrame) -> pd.DataFrame:
    """Extract broad wall type from WALLS_DESCRIPTION string: cavity, solid, other."""
    df = df.copy()
    desc = df['WALLS_DESCRIPTION'].str.lower().fillna('')
    df['WALL_TYPE'] = 'other'
    df.loc[desc.str.contains('cavity'), 'WALL_TYPE'] = 'cavity'
    df.loc[desc.str.contains('solid'), 'WALL_TYPE'] = 'solid'
    return df


def engineer_target(df: pd.DataFrame, gap_threshold: int = 20) -> pd.DataFrame:
    """
    Create binary target: 1 if property is D-G rated AND efficiency gap >= gap_threshold.
    This identifies properties with the most headroom for retrofit improvement.
    """
    df = df.copy()
    low_rated = df['CURRENT_ENERGY_RATING'].isin(['D', 'E', 'F', 'G'])
    gap = df['POTENTIAL_ENERGY_EFFICIENCY'] - df['CURRENT_ENERGY_EFFICIENCY']
    df['RETROFIT_POTENTIAL'] = ((low_rated) & (gap >= gap_threshold)).astype(int)
    df['EFFICIENCY_GAP'] = gap
    return df


def build_feature_matrix(df: pd.DataFrame) -> tuple:
    """
    Select and encode features for modelling.
    Returns X (feature DataFrame) and y (target Series).
    """
    df = df.copy()

    # Numeric features
    numeric_cols = [
        'CURRENT_ENERGY_EFFICIENCY',
        'TOTAL_FLOOR_AREA',
        'NUMBER_HABITABLE_ROOMS',
        'AGE_BAND_YEAR',
    ]

    # Categorical features to one-hot encode
    categorical_cols = [
        'PROPERTY_TYPE',
        'BUILT_FORM',
        'TENURE',
        'WALL_TYPE',
    ]

    # Impute missing numerics with median
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].fillna(df[col].median())

    # Impute missing categoricals with 'unknown'
    for col in categorical_cols:
        if col in df.columns:
            df[col] = df[col].fillna('unknown')

    # One-hot encode categoricals
    df_encoded = pd.get_dummies(df[numeric_cols + categorical_cols], drop_first=True)

    X = df_encoded
    y = df['RETROFIT_POTENTIAL']
    return X, y
