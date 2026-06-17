"""Excel dosya okuma ve yazma işlemleri."""

import glob
import os

import pandas as pd


def find_excel_files(path: str) -> list[str]:
    """Verilen yoldaki tüm Excel dosyalarını bulur."""
    if os.path.isfile(path) and path.lower().endswith(".xlsx"):
        return [path]
    elif os.path.isdir(path):
        files = glob.glob(os.path.join(path, "*.xlsx"))
        return [f for f in files if not os.path.basename(f).startswith("~$")]
    return []


def read_excel_safe(
    file_path: str, headers_to_try: tuple[int, ...] = (0, 2)
) -> pd.DataFrame | None:
    """Excel dosyasını güvenle okur, farklı header satırlarını dener."""
    for h in headers_to_try:
        try:
            df = pd.read_excel(file_path, header=h, dtype=str)
            return df
        except Exception:
            pass
    return None
