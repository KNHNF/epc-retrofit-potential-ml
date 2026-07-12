"""
data_loader.py
Load EPC CSV data with correct encoding and optional sampling.
"""

import pandas as pd


def load_epc_csv(filepath: str, sample_n: int = None, seed: int = 42) -> pd.DataFrame:
    """
    Load a domestic EPC CSV file.
    EPC CSVs use latin-1 encoding (not utf-8).
    Set sample_n to load a random sample of that many rows.
    """
    df = pd.read_csv(filepath, encoding='latin-1', low_memory=False)
    print(f"Loaded {len(df):,} records from {filepath}")

    if sample_n and len(df) > sample_n:
        df = df.sample(n=sample_n, random_state=seed).reset_index(drop=True)
        print(f"Sampled down to {len(df):,} records")

    return df


def load_epc_chunked(filepath: str, chunksize: int = 50_000) -> pd.DataFrame:
    """
    Load a very large EPC CSV in chunks to avoid memory errors.
    Concatenates all chunks into one DataFrame.
    Use this if load_epc_csv causes a MemoryError.
    """
    chunks = []
    for chunk in pd.read_csv(filepath, encoding='latin-1', low_memory=False, chunksize=chunksize):
        chunks.append(chunk)
    df = pd.concat(chunks, ignore_index=True)
    print(f"Loaded {len(df):,} records in chunks")
    return df
