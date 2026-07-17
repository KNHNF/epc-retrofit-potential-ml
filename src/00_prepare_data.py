"""
00_prepare_data.py
Process year-by-year raw EPC CSV files, clean, deduplicate by UPRN,
engineer the retrofit potential target, and save train/test splits as Parquet.
"""

import os
import pandas as pd
import numpy as np

# Config
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(_PROJECT_ROOT, "data", "raw")        # source: raw certificates-*.csv
OUTPUT_DIR = os.path.join(_PROJECT_ROOT, "data", "processed")  # destination: parquet outputs

COLUMNS_TO_KEEP = [
    'UPRN', 'LODGEMENT_DATE',
    'CURRENT_ENERGY_RATING', 'POTENTIAL_ENERGY_RATING',
    'CURRENT_ENERGY_EFFICIENCY', 'POTENTIAL_ENERGY_EFFICIENCY',
    'PROPERTY_TYPE', 'BUILT_FORM', 'CONSTRUCTION_AGE_BAND',
    'WALLS_DESCRIPTION', 'TOTAL_FLOOR_AREA', 'NUMBER_HABITABLE_ROOMS',
    'TENURE', 'TRANSACTION_TYPE', 'MAINS_GAS_FLAG', 'REGION', 'COUNTRY',
    'CO2_EMISS_CURR_PER_FLOOR_AREA', 'CO2_EMISSIONS_CURRENT', 'CO2_EMISSIONS_POTENTIAL',
    'ENERGY_CONSUMPTION_CURRENT', 'ENERGY_CONSUMPTION_POTENTIAL',
    'HEATING_COST_CURRENT', 'HEATING_COST_POTENTIAL',
    'HOT_WATER_COST_CURRENT', 'LIGHTING_COST_CURRENT',
    'EXTENSION_COUNT', 'FIXED_LIGHTING_OUTLETS_COUNT',
    'MULTI_GLAZE_PROPORTION', 'LOW_ENERGY_LIGHTING', 'NUMBER_HEATED_ROOMS',
    'WALLS_ENERGY_EFF', 'ROOF_ENERGY_EFF', 'FLOOR_ENERGY_EFF',
    'WINDOWS_ENERGY_EFF', 'MAINHEAT_ENERGY_EFF', 'HOT_WATER_ENERGY_EFF',
    'LIGHTING_ENERGY_EFF', 'MAINHEATC_ENERGY_EFF'
]

VALID_ENERGY_RATINGS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
VALID_PROPERTY_TYPES = ['House', 'Flat', 'Maisonette', 'Bungalow', 'Park home']


def process_year_file(filepath: str, seen_uprns: set) -> pd.DataFrame:
    """
    Process a single year CSV file in chunks, clean, and filter.
    seen_uprns is updated in-place.
    """
    print(f"Processing {os.path.basename(filepath)}...")
    chunks = []
    
    # Read in chunks to prevent memory overload
    chunksize = 100_000
    for chunk in pd.read_csv(filepath, encoding='latin-1', low_memory=False, chunksize=chunksize):
        # Convert all columns to uppercase to avoid case mismatches
        chunk.columns = chunk.columns.str.upper()
        
        # Check that required columns exist
        missing_cols = [c for c in ['UPRN', 'LODGEMENT_DATE', 'CURRENT_ENERGY_RATING', 
                                    'POTENTIAL_ENERGY_RATING', 'PROPERTY_TYPE',
                                    'CURRENT_ENERGY_EFFICIENCY', 'POTENTIAL_ENERGY_EFFICIENCY'] 
                        if c not in chunk.columns]
        if missing_cols:
            print(f"Skipping chunk, missing columns: {missing_cols}")
            continue
            
        # Select columns that are present
        cols_to_select = [c for c in COLUMNS_TO_KEEP if c in chunk.columns]
        chunk = chunk[cols_to_select]
        
        # Drop rows with null UPRN or LODGEMENT_DATE
        chunk = chunk.dropna(subset=['UPRN', 'LODGEMENT_DATE'])
        
        # Clean and standardise string values for filtering
        chunk['CURRENT_ENERGY_RATING'] = chunk['CURRENT_ENERGY_RATING'].str.strip().str.upper()
        chunk['POTENTIAL_ENERGY_RATING'] = chunk['POTENTIAL_ENERGY_RATING'].str.strip().str.upper()
        chunk['PROPERTY_TYPE'] = chunk['PROPERTY_TYPE'].str.strip()
        
        # Filter valid records
        chunk = chunk[
            chunk['CURRENT_ENERGY_RATING'].isin(VALID_ENERGY_RATINGS) &
            chunk['POTENTIAL_ENERGY_RATING'].isin(VALID_ENERGY_RATINGS) &
            chunk['PROPERTY_TYPE'].isin(VALID_PROPERTY_TYPES)
        ]
        
        # Convert efficiency scores to numeric and filter valid range (1-100)
        chunk['CURRENT_ENERGY_EFFICIENCY'] = pd.to_numeric(chunk['CURRENT_ENERGY_EFFICIENCY'], errors='coerce')
        chunk['POTENTIAL_ENERGY_EFFICIENCY'] = pd.to_numeric(chunk['POTENTIAL_ENERGY_EFFICIENCY'], errors='coerce')
        chunk = chunk.dropna(subset=['CURRENT_ENERGY_EFFICIENCY', 'POTENTIAL_ENERGY_EFFICIENCY'])
        chunk = chunk[
            (chunk['CURRENT_ENERGY_EFFICIENCY'] >= 1) & (chunk['CURRENT_ENERGY_EFFICIENCY'] <= 100) &
            (chunk['POTENTIAL_ENERGY_EFFICIENCY'] >= 1) & (chunk['POTENTIAL_ENERGY_EFFICIENCY'] <= 100)
        ]

        # Drop physically impossible negative values in energy, emissions and cost fields
        # (sentinel/entry errors in the source register). Coerce to numeric first so
        # string junk becomes NaN and is not treated as a valid non-negative value.
        NON_NEGATIVE_COLS = [
            'CO2_EMISS_CURR_PER_FLOOR_AREA', 'CO2_EMISSIONS_CURRENT',
            'ENERGY_CONSUMPTION_CURRENT', 'HEATING_COST_CURRENT',
            'HOT_WATER_COST_CURRENT', 'LIGHTING_COST_CURRENT',
            'TOTAL_FLOOR_AREA', 'NUMBER_HABITABLE_ROOMS', 'NUMBER_HEATED_ROOMS',
            'EXTENSION_COUNT', 'FIXED_LIGHTING_OUTLETS_COUNT',
            'MULTI_GLAZE_PROPORTION', 'LOW_ENERGY_LIGHTING',
        ]
        neg_present = [c for c in NON_NEGATIVE_COLS if c in chunk.columns]
        for c in neg_present:
            chunk[c] = pd.to_numeric(chunk[c], errors='coerce')
        if neg_present:
            # keep rows where every present column is either missing (imputed later) or >= 0
            chunk = chunk[((chunk[neg_present] >= 0) | chunk[neg_present].isna()).all(axis=1)]

        # Convert UPRN to clean string to standardise comparison
        # Remove trailing decimals if parsed as floats
        chunk['UPRN'] = chunk['UPRN'].astype(str).str.split('.').str[0].str.strip()
        chunk = chunk[chunk['UPRN'] != 'nan']
        
        # Parse dates and sort chunk descending
        chunk['LODGEMENT_DATE_DT'] = pd.to_datetime(chunk['LODGEMENT_DATE'], errors='coerce')
        chunk = chunk.dropna(subset=['LODGEMENT_DATE_DT'])
        chunk = chunk.sort_values('LODGEMENT_DATE_DT', ascending=False)
        
        # Drop duplicates within the chunk
        chunk = chunk.drop_duplicates(subset='UPRN', keep='first')
        
        # Filter out records where UPRN has already been seen (newer record exists)
        chunk = chunk[~chunk['UPRN'].isin(seen_uprns)]
        
        # Add new UPRNs to the seen set
        seen_uprns.update(chunk['UPRN'].tolist())
        
        # Drop helper date column
        chunk = chunk.drop(columns=['LODGEMENT_DATE_DT'])
        
        chunks.append(chunk)
        
    if not chunks:
        return pd.DataFrame()
        
    df = pd.concat(chunks, ignore_index=True)
    # Final deduplication just in case
    df = df.sort_values('LODGEMENT_DATE', ascending=False)
    df = df.drop_duplicates(subset='UPRN', keep='first')
    
    return df


def main():
    years = [2026, 2025, 2024, 2023, 2022, 2021, 2020]
    seen_uprns = set()
    
    train_dfs = []
    test_dfs = []
    
    for year in years:
        filename = f"certificates-{year}.csv"
        filepath = os.path.join(DATA_DIR, filename)
        if not os.path.exists(filepath):
            print(f"Warning: File {filename} not found, skipping.")
            continue
            
        df_year = process_year_file(filepath, seen_uprns)
        if df_year.empty:
            continue
            
        # Temporal split: 2025-2026 is test set, 2020-2024 is train set
        if year in [2025, 2026]:
            test_dfs.append(df_year)
            print(f"Added {len(df_year):,} unique records from {year} to TEST set")
        else:
            train_dfs.append(df_year)
            print(f"Added {len(df_year):,} unique records from {year} to TRAIN set")
            
    print("\nConcatenating train and test sets...")
    
    if test_dfs:
        test_full = pd.concat(test_dfs, ignore_index=True)
    else:
        test_full = pd.DataFrame()
        
    if train_dfs:
        train_full = pd.concat(train_dfs, ignore_index=True)
    else:
        train_full = pd.DataFrame()
        
    print(f"Train set full unique records: {len(train_full):,}")
    print(f"Test set full unique records: {len(test_full):,}")
    
    # Engineer target variable for train and test
    def engineer_target_in_place(df):
        if df.empty:
            return df
        df['EFFICIENCY_GAP'] = df['POTENTIAL_ENERGY_EFFICIENCY'] - df['CURRENT_ENERGY_EFFICIENCY']
        low_rated = df['CURRENT_ENERGY_RATING'].isin(['D', 'E', 'F', 'G'])
        df['RETROFIT_POTENTIAL'] = ((low_rated) & (df['EFFICIENCY_GAP'] >= 20)).astype(int)
        return df

    print("Engineering target variables...")
    train_full = engineer_target_in_place(train_full)
    test_full = engineer_target_in_place(test_full)
    
    # Report class balances
    if not train_full.empty:
        train_balance = train_full['RETROFIT_POTENTIAL'].value_counts(normalize=True)
        print(f"Train class balance: \n{train_balance}")
    if not test_full.empty:
        test_balance = test_full['RETROFIT_POTENTIAL'].value_counts(normalize=True)
        print(f"Test class balance: \n{test_balance}")
        
    # Save full files as Parquet
    print("Saving full datasets to Parquet...")
    train_path = os.path.join(OUTPUT_DIR, "epc_train_full.parquet")
    test_path = os.path.join(OUTPUT_DIR, "epc_test_full.parquet")
    
    if not train_full.empty:
        train_full.to_parquet(train_path, index=False)
        print(f"Saved {train_path}")
    if not test_full.empty:
        test_full.to_parquet(test_path, index=False)
        print(f"Saved {test_path}")
        
    # Extract smaller stratified samples for quick development
    print("Extracting stratified samples for development...")
    
    def get_stratified_sample(df, n_samples=200_000, seed=42):
        if df.empty or len(df) <= n_samples:
            return df
        # Stratify by target variable
        pos = df[df['RETROFIT_POTENTIAL'] == 1]
        neg = df[df['RETROFIT_POTENTIAL'] == 0]
        
        pos_ratio = len(pos) / len(df)
        n_pos = int(n_samples * pos_ratio)
        n_neg = n_samples - n_pos
        
        # Sample with fallback if class sizes are too small
        pos_sample = pos.sample(n=min(n_pos, len(pos)), random_state=seed)
        neg_sample = neg.sample(n=min(n_neg, len(neg)), random_state=seed)
        
        sample = pd.concat([pos_sample, neg_sample], ignore_index=True)
        # Shuffle
        sample = sample.sample(frac=1, random_state=seed).reset_index(drop=True)
        return sample

    if not train_full.empty:
        train_sample = get_stratified_sample(train_full, n_samples=200_000)
        train_sample_path = os.path.join(OUTPUT_DIR, "epc_train_sample_200k.parquet")
        train_sample.to_parquet(train_sample_path, index=False)
        print(f"Saved 200k train sample to {train_sample_path}")
        
    if not test_full.empty:
        test_sample = get_stratified_sample(test_full, n_samples=50000)
        test_sample_path = os.path.join(OUTPUT_DIR, "epc_test_sample_50k.parquet")
        test_sample.to_parquet(test_sample_path, index=False)
        print(f"Saved 50k test sample to {test_sample_path}")
        
    print("Data preparation complete.")


if __name__ == "__main__":
    main()
