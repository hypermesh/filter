"""
Stok Rezervasyon Motoru
-----------------------
Öncelik sırasına göre sıralanmış Excel reçetelerini işler ve
StokListesi.xlsx verilerini kullanarak FIFO stok rezervasyonu yapar.

Her kaynak dosya için:
  - 'Aktif Rezerve Edilen Miktar', 'Kalan', 'Çıkan', 'Rezerve Edilecek Miktar'
    sütunları önce sıfırlanır.
  - İhtiyaç miktarına göre mevcut stoktan düşülür (kısmi rezervasyon: Seçenek A).
  - Çıktı dosyası orijinal adın sonuna '_' eklenerek aynı klasöre kaydedilir.
"""

from __future__ import annotations

import os
import warnings
from typing import Any

import pandas as pd


# ── Türkçe karakter normalleştirici ──────────────────────────────────────────

_REPLACEMENTS: dict[str, str] = {
    "ı": "i", "İ": "i", "I": "i",
    "ş": "s", "Ş": "s",
    "ğ": "g", "Ğ": "g",
    "ü": "u", "Ü": "u",
    "ö": "o", "Ö": "o",
    "ç": "c", "Ç": "c",
}


def norm_col(text: str) -> str:
    """Sütun adını Türkçe karakter ve boşluklardan arındırarak küçük harfe çevirir."""
    val = str(text)
    for tr, eng in _REPLACEMENTS.items():
        val = val.replace(tr, eng)
    return val.strip().lower().replace(" ", "").replace("_", "")


def safe_float(val: Any) -> float:
    """Herhangi bir değeri float'a güvenli şekilde çevirir."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def safe_kod(val: Any) -> str:
    """Parça kodunu temiz string'e çevirir; '.0' uzantısını kaldırır."""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


# ── StokListesi okuma ─────────────────────────────────────────────────────────

def load_stok_dict(stok_xlsx_path: str) -> dict[str, float]:
    """
    StokListesi.xlsx dosyasından {malzeme_kodu: mevcut_stok} sözlüğü döndürür.
    Kullanılabilir stok = Fiziki Stok - Rezerve olarak hesaplanır.
    Dosya yoksa boş sözlük döner.
    """
    if not os.path.exists(stok_xlsx_path):
        return {}

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df_stok = pd.read_excel(stok_xlsx_path, header=0)

    kod_col: str | None = None
    stok_col: str | None = None
    rezerve_col: str | None = None

    for c in df_stok.columns:
        nc = norm_col(str(c))
        if nc in ("malzemekodu", "kod"):
            kod_col = c
        elif nc == "stok" and stok_col is None:
            stok_col = c
        elif "rezerve" in nc and rezerve_col is None:
            rezerve_col = c

    if not kod_col or not stok_col:
        return {}

    stok_dict: dict[str, float] = {}
    for _, row in df_stok.iterrows():
        k = safe_kod(row[kod_col])
        if k and k.lower() != "nan":
            # Sadece fiziki stok değerini alıyoruz.
            # Rezerve hesabı yapılmıyor — tüm işlemler kaynak dosyalar üzerinde gerçekleşecek.
            stok_dict[k] = safe_float(row[stok_col])

    return stok_dict


# ── Sütun tespit yardımcıları ─────────────────────────────────────────────────

# Sütun arama hedefleri (normalleştirilmiş)
_KOD_TARGETS = ("malzemekodu", "kod")
_IHTYAC_NORMALIZED = "rezerveedilecekmiktar"  # norm_col("Rezerve Edilecek Miktar")
_IHTYAC_FALLBACKS = ("carpilmismiktar", "miktar")  # Alternatif ihtiyaç sütunları

_AKTIF_REZERVE_LABEL = "Aktif Rezerve Edilen Miktar"
_REZERVE_EDILECEK_LABEL = "Rezerve Edilecek Miktar"
_KALAN_LABEL = "Kalan"
_CIKAN_LABEL = "Çıkan"


def _find_col(df: pd.DataFrame, *targets: str) -> str | None:
    """Normalleştirilmiş hedef listesinden ilk eşleşen sütun adını döndürür."""
    for c in df.columns:
        nc = norm_col(str(c))
        if nc in targets:
            return c
    return None


def _find_col_contains(df: pd.DataFrame, keyword: str) -> str | None:
    """Normalleştirilmiş sütun adında keyword geçen ilk sütunu döndürür."""
    for c in df.columns:
        nc = norm_col(str(c))
        if keyword in nc:
            return c
    return None


# ── Tek dosya rezervasyonu ────────────────────────────────────────────────────

def reserve_single_file(
    src_path: str,
    stok_dict: dict[str, float],
    output_suffix: str = "_",
) -> tuple[str, dict[str, Any]]:
    """
    Tek bir Excel dosyasını işler:
      1. İlgili sütunları sıfırlar.
      2. Stok sözlüğünü kullanarak kısmi rezervasyon hesaplar.
      3. Stok sözlüğünü günceller (in-place).
      4. Çıktı dosyasını 'isim_ .xlsx' olarak kaydeder.

    Döndürür:
        (output_path, stats_dict)
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(src_path, header=0, dtype=str)

    stats: dict[str, Any] = {
        "satir": len(df),
        "rezerve_yapilan": 0,
        "eksik_olan": 0,
        "uyari": [],
    }

    # ── Sütunları tespit et ──────────────────────────────────────────────────

    # Parça kodu sütunu
    kod_col = _find_col(df, *_KOD_TARGETS)
    if kod_col is None:
        # Fallback: "Sıra No" / "Sira No" gibi ID sütunları
        for c in df.columns:
            nc = norm_col(str(c))
            if nc in ("sirano", "sira"):
                kod_col = c
                break
    if kod_col is None:
        stats["uyari"].append("Parça kodu sütunu bulunamadı, dosya atlandı.")
        return "", stats

    # İhtiyaç (istek) miktarı sütunu — önce "Rezerve Edilecek Miktar", sonra alternatifler
    ihtyac_col: str | None = None
    for c in df.columns:
        if norm_col(str(c)) == _IHTYAC_NORMALIZED:
            ihtyac_col = c
            break
    if ihtyac_col is None:
        for fb in _IHTYAC_FALLBACKS:
            ihtyac_col = _find_col_contains(df, fb)
            if ihtyac_col:
                break

    if ihtyac_col is None:
        stats["uyari"].append("İhtiyaç miktarı sütunu bulunamadı, dosya atlandı.")
        return "", stats

    # Aktif rezerve edilen miktar sütunu (yoksa oluştur)
    aktif_col = _find_col_contains(df, "aktif")
    if aktif_col is None:
        aktif_col = _AKTIF_REZERVE_LABEL
        df[aktif_col] = "0"

    # Rezerve edilecek miktar — ihtiyaç sütunu ile aynı olabilir (ID-tabanlı dosyalar)
    rezerve_col: str | None = None
    for c in df.columns:
        if norm_col(str(c)) == _IHTYAC_NORMALIZED and c != ihtyac_col:
            rezerve_col = c
            break
    if rezerve_col is None:
        # İhtyaç sütunu ile aynıysa veya bulunamazsa — ihtyac_col'u da rezerve sütunu yap
        rezerve_col = ihtyac_col

    # Kalan sütunu (varsa sıfırla, yoksa ekle)
    kalan_col = _find_col_contains(df, "kalan")
    if kalan_col is None:
        kalan_col = _KALAN_LABEL
        df[kalan_col] = "0"

    # Çıkan sütunu (varsa sıfırla, yoksa ekle)
    cikan_col: str | None = None
    for c in df.columns:
        nc = norm_col(str(c))
        if nc in ("cikan", "cikanmiktar"):
            cikan_col = c
            break
    if cikan_col is None:
        cikan_col = _CIKAN_LABEL
        df[cikan_col] = "0"

    # ── İHTİYAÇ MİKTARLARINI ÖNCEDEN KAYDET ──────────────────────────────────
    # Sütun sıfırlamadan ÖNCE ihtiyaç değerlerini ayrı bir sözlükte sakla.
    # (ihtyac_col ile rezerve_col aynı sütun olabilir — sıfırlamadan önce kaydet!)
    ihtyac_values: dict[int, float] = {}
    for idx, row in df.iterrows():
        ihtyac_values[int(idx)] = safe_float(row[ihtyac_col])

    # ── Tüm hedef sütunları sıfırla ─────────────────────────────────────────
    for col in (aktif_col, rezerve_col, kalan_col, cikan_col):
        if col in df.columns:
            df[col] = 0.0

    # ── Satır satır rezervasyon hesabı ─────────────────────────────────────
    for idx, row in df.iterrows():
        k = safe_kod(row[kod_col])
        if not k or k.lower() == "nan":
            continue

        # İhtiyaç miktarını önceden kaydettiğimiz sözlükten al
        ihtyac = ihtyac_values.get(int(idx), 0.0)
        if ihtyac <= 0:
            continue

        mevcut_stok = stok_dict.get(k, 0.0)

        if mevcut_stok >= ihtyac:
            # Yeterli stok: tamamı rezerve edilir
            aktif_rezerve = ihtyac
            yeni_rezerve_edilecek = 0.0
            stok_dict[k] = mevcut_stok - ihtyac
            stats["rezerve_yapilan"] += 1
        elif mevcut_stok > 0:
            # Kısmi stok: Seçenek A — olan kadarı rezerve, kalanı eksik
            aktif_rezerve = mevcut_stok
            yeni_rezerve_edilecek = ihtyac - mevcut_stok
            stok_dict[k] = 0.0
            stats["rezerve_yapilan"] += 1
            stats["eksik_olan"] += 1
        else:
            # Stok yok: hiç rezerve edilemez
            aktif_rezerve = 0.0
            yeni_rezerve_edilecek = ihtyac
            stats["eksik_olan"] += 1

        df.at[idx, aktif_col] = aktif_rezerve
        df.at[idx, rezerve_col] = yeni_rezerve_edilecek
        df.at[idx, kalan_col] = 0.0
        df.at[idx, cikan_col] = 0.0

    # ── Çıktı dosyası ────────────────────────────────────────────────────────
    dirname = os.path.dirname(src_path)
    basename = os.path.basename(src_path)
    name_no_ext = os.path.splitext(basename)[0]
    out_name = f"{name_no_ext}{output_suffix}.xlsx"
    out_path = os.path.join(dirname, out_name)

    df.to_excel(out_path, index=False)
    return out_path, stats
