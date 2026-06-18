"""Filtreleme servisleri."""

import json
import os
from typing import Any

import pandas as pd
from rich.console import Console

from recipe_automation.core.config import settings

console = Console()


def apply_exclusions(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str], int]:
    """haric_tutulacak_parcalar.json dosyasını okur ve eşleşen KODLARI DataFrame'den çıkarır."""
    db_dir = os.path.join(os.getcwd(), settings.db_dir_name)
    json_path = os.path.join(db_dir, "haric_tutulacak_parcalar.json")

    if not os.path.exists(json_path):
        console.print(
            f"[yellow]Uyarı:[/yellow] {json_path} bulunamadı, hariç tutma işlemi atlanıyor."
        )
        return df, [], 0

    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        console.print(f"[red]JSON Okuma Hatası:[/red] {e}")
        return df, [], 0

    if not isinstance(data, list) or len(data) == 0:
        return df, [], 0

    df_filtered = df.copy()

    kod_col = None
    sira_col = None

    # Sütunları bul
    for c in df_filtered.columns:
        c_lower = str(c).strip().lower()
        if c_lower == "kod":
            kod_col = c
        elif c_lower == "sıra no" or c_lower == "sira no":
            sira_col = c

    if kod_col and sira_col:
        tam_kodlar_upper = [str(k).strip().upper() for k in data]
        df_filtered["__temp_kod__"] = df_filtered[kod_col].astype(str).str.strip().str.upper()

        # PANDAS KÜSURAT DÜZELTMESİ ("1.010" -> "1.01" sorununu engellemek için)
        # Eğer sıfırla biten ondalıklı bir sayı gelirse string'e çevirirken dikkatli olmalıyız
        # Ancak en güvenlisi direkt string formatına zorlamaktır.
        df_filtered["__temp_sira__"] = df_filtered[sira_col].astype(str).str.strip()

        # 1. Adım: Verilen kodların Sıra Numaralarını bul
        hedef_satirlar = df_filtered[df_filtered["__temp_kod__"].isin(tam_kodlar_upper)]
        hedef_sira_nolar = hedef_satirlar["__temp_sira__"].tolist()

        # 2. Adım: Alt kırılımları bulmak için sonuna "." ekle
        ust_sira_prefixleri = [
            str(s) + "." for s in hedef_sira_nolar if str(s).strip() and str(s).strip() != "nan"
        ]

        def is_deleted(kod: str, sira: str) -> bool:
            if kod in tam_kodlar_upper:
                return True
            if pd.isna(sira) or str(sira).strip() == "nan":
                return False

            sira_str = str(sira).strip()
            for prefix in ust_sira_prefixleri:
                if sira_str.startswith(prefix):
                    return True
            return False

        silinecek_maske = df_filtered.apply(
            lambda r: is_deleted(r["__temp_kod__"], r["__temp_sira__"]), axis=1
        )
        silinen_adet = silinecek_maske.sum()
        silinen_kodlar = []

        if silinen_adet > 0:
            silinen_kodlar = df_filtered[silinecek_maske]["__temp_kod__"].unique().tolist()
            df_filtered = df_filtered[~silinecek_maske]

        df_filtered = df_filtered.drop(columns=["__temp_kod__", "__temp_sira__"], errors="ignore")
        return df_filtered, silinen_kodlar, silinen_adet

    elif kod_col:
        # Eski mantık: Eğer Sıra No sütunu bir sebeple yoksa sadece kodları sil
        tam_kodlar_upper = [str(k).strip().upper() for k in data]
        kod_series = df_filtered[kod_col].astype(str).str.strip().str.upper()

        silinecek_maske = kod_series.isin(tam_kodlar_upper)
        silinen_adet = silinecek_maske.sum()
        silinen_kodlar = []

        if silinen_adet > 0:
            silinen_kodlar = df_filtered[silinecek_maske][kod_col].astype(str).unique().tolist()
            df_filtered = df_filtered[~silinecek_maske]

        return df_filtered, silinen_kodlar, silinen_adet

    return df_filtered, [], 0


def _is_target_operation(row: pd.Series, op_cols: list[str], target_group: set[str]) -> bool:
    """Verilen satırın herhangi bir operasyonunda hedef gruptan bir operasyon var mı kontrol eder."""
    for col in op_cols:
        val = str(row.get(col, "")).strip().upper()
        if val and val in target_group:
            return True
    return False


def aggregate_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Filtrelenen listede aynı koddan (Kod veya Sıra No) birden fazla varsa,
    ilk satırın bilgilerini koruyarak sadece 'Çarpılmış Miktar' sütunlarını toplar.
    Çarpılmış miktar sütunu bulunamasa bile mükerrerleri temizler."""
    kod_col = None
    for c in df.columns:
        if str(c).strip().lower() == "kod":
            kod_col = c
            break

    if not kod_col and settings.col_sira_no_id in df.columns:
        kod_col = settings.col_sira_no_id

    if not kod_col:
        return df

    df = df.copy()
    # Excel'den gelen görünmez boşlukları ve tip farklarını (int/str) eşitle!
    df[kod_col] = df[kod_col].astype(str).str.strip().str.upper()

    carp_col = None
    exact_targets = [
        "Çarpılmış Miktar",
        "Carpilmis Miktar",
        settings.col_rezerve_miktar,
        settings.col_kullanilabilir_stok,
    ]
    for target in exact_targets:
        if target in df.columns:
            carp_col = target
            break

    agg_dict = {}
    for col in df.columns:
        if col == kod_col:
            continue
        elif carp_col and col == carp_col:
            agg_dict[col] = "sum"
        else:
            agg_dict[col] = "first"

    if carp_col:
        df[carp_col] = pd.to_numeric(
            df[carp_col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)

    df_grouped = df.groupby(kod_col, as_index=False, dropna=False).agg(agg_dict)

    # Orijinal sütun sıralamasını koru
    return df_grouped[df.columns]


def filter_id_based(
    df: pd.DataFrame, target_group: set[str]
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Sıra No ve Rezerve Edilecek Miktar bazlı filtreleme."""
    kod_sutunu = settings.col_sira_no_id
    hedef_sutun = settings.col_rezerve_miktar

    if kod_sutunu not in df.columns or hedef_sutun not in df.columns:
        raise ValueError(f"Gerekli sütunlar bulunamadı: '{kod_sutunu}' veya '{hedef_sutun}'")

    op_cols = [
        str(c) for c in df.columns if settings.col_operasyon_keyword.lower() in str(c).lower()
    ]

    metadata = {
        "kod_sutunu": kod_sutunu,
        "hedef_sutun": hedef_sutun,
        "op_cols": op_cols,
        "satir_ilk": len(df),
    }

    df = df.copy()
    df["__target_num__"] = (
        df[hedef_sutun].astype(str).str.replace(",", ".").astype(float).fillna(0.0)
    )

    sifir_olanlar = df[df["__target_num__"] == 0.0]
    sifir_kodlar = [str(kod).strip() + "." for kod in sifir_olanlar[kod_sutunu] if pd.notna(kod)]

    def is_deleted(sira_no: str, miktar: float) -> bool:
        if miktar == 0.0:
            return True
        if pd.isna(sira_no):
            return False
        sira_str = str(sira_no).strip()
        for ust_kod in sifir_kodlar:
            if sira_str.startswith(ust_kod):
                return True
        return False

    silinecek_maske = df.apply(lambda r: is_deleted(r[kod_sutunu], r["__target_num__"]), axis=1)
    tim_maskesi = df.apply(lambda r: _is_target_operation(r, op_cols, target_group), axis=1)

    # Orijinal Toplamı Bulma (Kırılım öncesi, sadece ilgili gruba ait parçaların toplamı)
    df_orijinal_grup = df[tim_maskesi].copy()
    if not df_orijinal_grup.empty:
        df_orijinal_grup = aggregate_duplicates(df_orijinal_grup)

    carp_col = None
    exact_targets = [
        "Çarpılmış Miktar",
        "Carpilmis Miktar",
        settings.col_rezerve_miktar,
        settings.col_kullanilabilir_stok,
    ]
    for target in exact_targets:
        if target in df.columns:
            carp_col = target
            break

    brut_col = None
    brut_targets = ["Miktar", "Çarpılmış Miktar", "Carpilmis Miktar"]
    for target in brut_targets:
        if target in df.columns:
            brut_col = target
            break

    orijinal_grup_toplami = 0
    if brut_col and not df_orijinal_grup.empty and brut_col in df_orijinal_grup.columns:
        brut_series = pd.to_numeric(
            df_orijinal_grup[brut_col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)
        orijinal_grup_toplami = brut_series.sum()
    elif carp_col and not df_orijinal_grup.empty and carp_col in df_orijinal_grup.columns:
        # Fallback
        carp_series = pd.to_numeric(
            df_orijinal_grup[carp_col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)
        orijinal_grup_toplami = carp_series.sum()

    filtered_df = df[tim_maskesi & ~silinecek_maske].copy()
    filtered_df = filtered_df.drop(columns=["__target_num__"])

    # Çift kayıtları temizle ve miktarlarını topla
    filtered_df = aggregate_duplicates(filtered_df)

    kalan_grup_toplami = 0
    if carp_col and not filtered_df.empty and carp_col in filtered_df.columns:
        kalan_grup_toplami = filtered_df[carp_col].sum()

    metadata["satir_son"] = len(filtered_df)
    metadata["orijinal_grup_toplami"] = float(orijinal_grup_toplami)
    metadata["kalan_grup_toplami"] = float(kalan_grup_toplami)
    metadata["orijinal_kalem_sayisi"] = len(df_orijinal_grup) if not df_orijinal_grup.empty else 0

    return filtered_df, metadata


def filter_stock_based(
    df: pd.DataFrame, target_group: set[str]
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Sira ve Kullanilabilir Stok bazlı filtreleme."""
    kod_sutunu = settings.col_sira_no_stock
    hedef_sutun = settings.col_kullanilabilir_stok

    if kod_sutunu not in df.columns or hedef_sutun not in df.columns:
        raise ValueError(f"Gerekli sütunlar bulunamadı: '{kod_sutunu}' veya '{hedef_sutun}'")

    op_cols = [
        str(c) for c in df.columns if settings.col_operasyon_keyword.lower() in str(c).lower()
    ]

    metadata = {
        "kod_sutunu": kod_sutunu,
        "hedef_sutun": hedef_sutun,
        "op_cols": op_cols,
        "satir_ilk": len(df),
    }

    df = df.copy()
    df["__stok_num__"] = df[hedef_sutun].astype(str).str.replace(",", ".").astype(float).fillna(0.0)

    stok_olanlar = df[df["__stok_num__"] > 0.0]
    stok_kodlar = [str(kod).strip() + "." for kod in stok_olanlar[kod_sutunu] if pd.notna(kod)]

    def is_deleted(sira_no: str, stok: float) -> bool:
        if stok > 0.0:
            return True
        if pd.isna(sira_no):
            return False
        sira_str = str(sira_no).strip()
        for ust_kod in stok_kodlar:
            if sira_str.startswith(ust_kod):
                return True
        return False

    silinecek_maske = df.apply(lambda r: is_deleted(r[kod_sutunu], r["__stok_num__"]), axis=1)
    tim_maskesi = df.apply(lambda r: _is_target_operation(r, op_cols, target_group), axis=1)

    # Orijinal Toplamı Bulma (Kırılım öncesi, sadece ilgili gruba ait parçaların toplamı)
    df_orijinal_grup = df[tim_maskesi].copy()
    if not df_orijinal_grup.empty:
        df_orijinal_grup = aggregate_duplicates(df_orijinal_grup)

    carp_col = None
    exact_targets = [
        "Çarpılmış Miktar",
        "Carpilmis Miktar",
        settings.col_rezerve_miktar,
        settings.col_kullanilabilir_stok,
    ]
    for target in exact_targets:
        if target in df.columns:
            carp_col = target
            break

    brut_col = None
    brut_targets = ["Miktar", "Çarpılmış Miktar", "Carpilmis Miktar"]
    for target in brut_targets:
        if target in df.columns:
            brut_col = target
            break

    orijinal_grup_toplami = 0
    if brut_col and not df_orijinal_grup.empty and brut_col in df_orijinal_grup.columns:
        brut_series = pd.to_numeric(
            df_orijinal_grup[brut_col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)
        orijinal_grup_toplami = brut_series.sum()
    elif carp_col and not df_orijinal_grup.empty and carp_col in df_orijinal_grup.columns:
        # Fallback
        carp_series = pd.to_numeric(
            df_orijinal_grup[carp_col].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)
        orijinal_grup_toplami = carp_series.sum()

    filtered_df = df[tim_maskesi & ~silinecek_maske].copy()
    filtered_df = filtered_df.drop(columns=["__stok_num__"])

    # Çift kayıtları temizle ve miktarlarını topla
    filtered_df = aggregate_duplicates(filtered_df)

    kalan_grup_toplami = 0
    if carp_col and not filtered_df.empty and carp_col in filtered_df.columns:
        kalan_grup_toplami = filtered_df[carp_col].sum()

    metadata["satir_son"] = len(filtered_df)
    metadata["orijinal_grup_toplami"] = float(orijinal_grup_toplami)
    metadata["kalan_grup_toplami"] = float(kalan_grup_toplami)
    metadata["orijinal_kalem_sayisi"] = len(df_orijinal_grup) if not df_orijinal_grup.empty else 0

    return filtered_df, metadata
