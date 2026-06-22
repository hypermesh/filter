import json
import os
import re

import pandas as pd

from recipe_automation.core.config import settings


def load_priority_mapping(db_dir: str) -> dict:
    """
    veritabanlari klasörü içindeki oncelik_sirasi.json dosyasını okur.
    Dosya yoksa boş dict döner.
    """
    json_path = os.path.join(db_dir, "oncelik_sirasi.json")
    if not os.path.exists(json_path):
        return {}

    try:
        with open(json_path, encoding="utf-8") as f:
            mapping = json.load(f)
            # Tüm anahtarları string olarak tut ve boşlukları temizle
            return {str(k).strip(): int(v) for k, v in mapping.items()}
    except Exception as e:
        print(f"Uyarı: Öncelik sırası dosyası okunamadı: {e}")
        return {}


def extract_file_names(kaynak_metin: str) -> list[str]:
    """
    '2241 (4), 2242 (5)' veya '2241, 2245' gibi metinlerden
    saf dosya isimlerini çıkarır: ['2241', '2242']
    """
    if pd.isna(kaynak_metin):
        return []

    metin = str(kaynak_metin)
    # Virgülle ayır
    parts = metin.split(",")

    names = []
    for part in parts:
        part = part.strip()
        if not part:
            continue

        # Parantez içindeki sayıları temizle: '2241 (4)' -> '2241'
        # Regex: boşluk ve ardından gelen parantez içindeki her şeyi sil
        clean_name = re.sub(r"\s*\([^)]*\)", "", part).strip()

        # Önündeki öncelik numarasını temizle: '2 - 1946' -> '1946'
        clean_name = re.sub(r"^\d+\s*-\s*", "", clean_name).strip()

        # Eğer 'Sadece_TIM_2241' veya '.xlsx' varsa onları da temizle
        clean_name = (
            clean_name.replace("Sadece_TIM_", "")
            .replace("Sadece_KZM5_", "")
            .replace(".xlsx", "")
            .strip()
        )
        names.append(clean_name)

    return names


def calculate_row_priority(kaynak_metin: str, mapping: dict) -> int:
    """
    Bir satırdaki kaynak dosya isimlerini alır,
    öncelik haritasında en küçük değere (en yüksek öncelik) sahip olanı bulur.
    Hiçbiri bulunamazsa 9999 döner.
    """
    names = extract_file_names(kaynak_metin)

    if not names:
        return 9999

    min_priority = 9999
    for name in names:
        # Tam eşleşme (Örn: '2241')
        if name in mapping:
            prio = mapping[name]
            if prio < min_priority:
                min_priority = prio
        else:
            # Kısmi eşleşme (Örn: mapping'de '2241' var, name '2241_REV2' ise)
            for k, v in mapping.items():
                if k in name or name in k:
                    if v < min_priority:
                        min_priority = v

    return min_priority


def sort_dataframe(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    DataFrame'i KAYNAK DOSYA sütunundaki değerlere göre önceliklendirip sıralar.
    """
    if df is None or df.empty:
        return df

    if "KAYNAK DOSYA" not in df.columns:
        return df

    # Her satırın önceliğini hesapla
    df_sorted = df.copy()
    oncelik_degerleri = df_sorted["KAYNAK DOSYA"].apply(
        lambda x: calculate_row_priority(x, mapping)
    )

    # "Öncelik Sırası" sütununu uygun konuma ekle (Örn: Kod sütununun soluna)
    hedef_index = 1  # Varsayılan olarak KAYNAK DOSYA'dan hemen sonra
    if settings.col_depo_kod in df_sorted.columns:
        hedef_index = df_sorted.columns.get_loc(settings.col_depo_kod)

    df_sorted.insert(hedef_index, "Öncelik Sırası", oncelik_degerleri)

    # Öncelik Sırası'na ve Kaynak Dosya adına göre sırala
    df_sorted = df_sorted.sort_values(by=["Öncelik Sırası", "KAYNAK DOSYA"])

    return df_sorted
