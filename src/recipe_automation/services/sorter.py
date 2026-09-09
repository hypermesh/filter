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


def load_completed_production_records(db_dir: str) -> set[tuple[str, str]]:
    """
    veritabanlari/uretim_gecmisi.json dosyasını okur.
    Tamamlanmış (kaynak_dosya, parca_kodu) çiftlerini set olarak döner.
    """
    json_path = os.path.join(db_dir, "uretim_gecmisi.json")
    if not os.path.exists(json_path):
        return set()

    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
            completed_set = set()
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict):
                        k_dosya = (
                            str(item.get("kaynak", item.get("kaynak_dosya", ""))).strip().upper()
                        )
                        k_kod = str(item.get("kod", item.get("parca_kodu", ""))).strip().upper()
                        is_done = item.get("tamamlandi", True)
                        if k_kod and is_done:
                            completed_set.add((k_dosya, k_kod))
            return completed_set
    except Exception as e:
        print(f"Uyarı: Üretim geçmişi dosyası okunamadı: {e}")
        return set()


def is_row_completed(row: pd.Series, completed_set: set[tuple[str, str]]) -> bool:
    """
    Bir satırın uretim_gecmisi.json veritabanına göre tamamlanıp tamamlanmadığını kontrol eder.
    """
    if not completed_set:
        return False

    kod_col = (
        settings.col_depo_kod if settings.col_depo_kod in row else ("Kod" if "Kod" in row else None)
    )
    if not kod_col:
        return False

    kod_val = str(row.get(kod_col, "")).strip().upper()
    if not kod_val:
        return False

    kaynak_val = str(row.get("KAYNAK DOSYA", "")).strip().upper()
    kaynak_names = extract_file_names(kaynak_val)

    # 1. Tam eşleşme veya genel kod eşleşmesi
    if ("", kod_val) in completed_set:
        return True

    # 2. Kaynak dosya bazlı eşleşme
    for k_name in kaynak_names:
        k_upper = k_name.strip().upper()
        if (k_upper, kod_val) in completed_set:
            return True

    return False


def sort_dataframe(
    df: pd.DataFrame,
    mapping: dict,
    completed_set: set[tuple[str, str]] | None = None,
    sort_completed_to_top: bool = False,
) -> pd.DataFrame:
    """
    DataFrame'i KAYNAK DOSYA sütunundaki değerlere göre önceliklendirip sıralar.
    Eğer sort_completed_to_top=True ise, tamamlanmış olanlar (Yapıldı) en üste (Grup 0),
    üretilecek aktif olanlar ise normal öncelik sırasına göre (Grup 1) sıralanır.
    """
    if df is None or df.empty:
        return df

    if "KAYNAK DOSYA" not in df.columns:
        return df

    df_sorted = df.copy()

    # Öncelik değerlerini hesapla
    oncelik_degerleri = df_sorted["KAYNAK DOSYA"].apply(
        lambda x: calculate_row_priority(x, mapping)
    )

    # "Öncelik Sırası" sütununu uygun konuma ekle (Örn: Kod sütununun soluna)
    if "Öncelik Sırası" not in df_sorted.columns:
        hedef_index = 1  # Varsayılan olarak KAYNAK DOSYA'dan hemen sonra
        if settings.col_depo_kod in df_sorted.columns:
            hedef_index = df_sorted.columns.get_loc(settings.col_depo_kod)
        df_sorted.insert(hedef_index, "Öncelik Sırası", oncelik_degerleri)
    else:
        df_sorted["Öncelik Sırası"] = oncelik_degerleri

    if sort_completed_to_top and completed_set:
        # Tamamlanma durumunu hesapla: 0 = Tamamlandı (En üst), 1 = Aktif Üretilecek
        df_sorted["__is_completed_group__"] = df_sorted.apply(
            lambda r: 0 if is_row_completed(r, completed_set) else 1, axis=1
        )
        df_sorted = df_sorted.sort_values(
            by=["__is_completed_group__", "Öncelik Sırası", "KAYNAK DOSYA"]
        )
        df_sorted = df_sorted.drop(columns=["__is_completed_group__"])
    else:
        # Öncelik Sırası'na ve Kaynak Dosya adına göre sırala
        df_sorted = df_sorted.sort_values(by=["Öncelik Sırası", "KAYNAK DOSYA"])

    return df_sorted
