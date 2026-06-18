"""Operasyon tarama servisi."""

import pandas as pd

from recipe_automation.core.config import settings


def extract_operations(df: pd.DataFrame) -> tuple[set[str], list[str]]:
    """DataFrame'den 'Operasyon' kelimesi geçen sütunları bulur ve unique değerleri döner."""
    operasyonlar: set[str] = set()
    op_cols = [
        str(c) for c in df.columns if settings.col_operasyon_keyword.lower() in str(c).lower()
    ]
    for col in op_cols:
        operasyonlar.update(df[col].dropna().unique())
    return operasyonlar, op_cols
