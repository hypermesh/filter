import pandas as pd

from recipe_automation.core.config import settings


def match_with_depo(filtered_df: pd.DataFrame, depo_df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtrelenmiş DataFrame ile Depo DataFrame'ini 'Kod' sütunu üzerinden eşleştirir.
    Sadece depoda bulunan (eşleşen) satırları döner. Eşleşmeyenler silinir.
    """
    # Filtrelenmiş listede aranacak muhtemel Kod sütunları
    # ID bazlı dosyalarda parça kodu "Sıra No" sütununda, KZM5 (Stok) bazlılarda ise "Kod" sütununda yer alıyor.
    kod_sutunlari_hedef = ["Kod", settings.col_sira_no_id]
    mevcut_kod_sutunu = None

    # Filtrelenmiş listedeki ID sütununu bul
    for c in kod_sutunlari_hedef:
        if c in filtered_df.columns:
            mevcut_kod_sutunu = c
            break

    if not mevcut_kod_sutunu:
        raise ValueError(
            f"Filtrelenmiş dosyada arama yapılacak ID sütunu bulunamadı. Arananlar: {kod_sutunlari_hedef}"
        )

    depo_kod_sutunu = settings.col_depo_kod
    if depo_kod_sutunu not in depo_df.columns:
        raise ValueError(
            f"Depo dosyasında eşleştirme için '{depo_kod_sutunu}' sütunu bulunamadı! (Lütfen Excel'i kontrol edin)"
        )

    # 1. Filtrelenmiş listedeki (Örn: Sadece_TIM_...) geçerli kodları temizle ve havuza (set) al
    aranacak_kodlar = set(
        filtered_df[mevcut_kod_sutunu].dropna().astype(str).str.strip().str.upper()
    )

    # 2. Depo Excelindeki (liste4) kodları temizle
    depo_kodlar = depo_df[depo_kod_sutunu].dropna().astype(str).str.strip().str.upper()

    # 3. Depo Excelinde maske oluştur (Depodaki bu kod, benim 'aranacak_kodlar' havuzumda var mı?)
    eslesen_maske = depo_kodlar.isin(aranacak_kodlar)

    # 4. DEPO dosyasını filtrele ve sadece bizim istediğimiz kodların olduğu yeni depo excelini dön
    matched_depo_df = depo_df.loc[depo_kodlar[eslesen_maske].index].copy()

    return matched_depo_df


def append_hammadde(matched_df: pd.DataFrame, hammadde_df: pd.DataFrame) -> pd.DataFrame:
    """
    Eşleştirilmiş ana depoya (matched_df), hammadde veritabanından (hammadde_df)
    'Hammadde Kod' ve 'Hammadde' sütunlarını 'Kod' üzerinden Left Join ile ekler.
    Bulunamayan değerler 'Bulunamadı' olarak işaretlenir.
    """
    depo_kod_sutunu = settings.col_depo_kod  # "Kod"

    if depo_kod_sutunu not in matched_df.columns:
        return matched_df  # Ana listede kod sütunu yoksa bir şey yapma
    if depo_kod_sutunu not in hammadde_df.columns:
        raise ValueError(f"Hammadde dosyasında '{depo_kod_sutunu}' sütunu bulunamadı!")

    # Sadece gerekli sütunları seçelim
    gerekli_sutunlar = [depo_kod_sutunu]
    hedef_sutunlar = [settings.col_hammadde_kod, settings.col_hammadde_isim, "Hammadde Miktar"]
    for col in hedef_sutunlar:
        if col in hammadde_df.columns:
            gerekli_sutunlar.append(col)

    if len(gerekli_sutunlar) == 1:
        return matched_df  # Eklenecek hammadde sütunu yok

    hammadde_sub = hammadde_df[gerekli_sutunlar].copy()

    # Küçük/büyük harf ve boşluk farklarını yok etmek için geçici bir eşleştirme sütunu yarat
    matched_df["_JOIN_KEY"] = matched_df[depo_kod_sutunu].astype(str).str.strip().str.upper()
    hammadde_sub["_JOIN_KEY"] = hammadde_sub[depo_kod_sutunu].astype(str).str.strip().str.upper()

    # Duplicate (Çift) hammaddeleri önle
    hammadde_sub = hammadde_sub.drop_duplicates(subset=["_JOIN_KEY"], keep="first")

    # Anahtar dışındaki orjinal 'Kod' sütununu silelim ki isim çakışması olmasın
    hammadde_sub = hammadde_sub.drop(columns=[depo_kod_sutunu])

    # VLOOKUP (Left Join)
    result_df = pd.merge(matched_df, hammadde_sub, on="_JOIN_KEY", how="left")

    # Eşleşmeyenleri doldur
    if settings.col_hammadde_kod in result_df.columns:
        result_df[settings.col_hammadde_kod] = result_df[settings.col_hammadde_kod].fillna(
            "Bulunamadı"
        )
    if settings.col_hammadde_isim in result_df.columns:
        result_df[settings.col_hammadde_isim] = result_df[settings.col_hammadde_isim].fillna(
            "Bulunamadı"
        )
    if "Hammadde Miktar" in result_df.columns:
        # Metin olarak kalmasını engelle, gerçek sayıya dönüştür ki Excel formülleri çalışsın
        result_df["Hammadde Miktar"] = pd.to_numeric(
            result_df["Hammadde Miktar"].astype(str).str.replace(",", "."), errors="coerce"
        ).fillna(0)

    # Geçici sütunu sil
    result_df = result_df.drop(columns=["_JOIN_KEY"])

    # "Hammadde Kod", "Hammadde" ve "Hammadde Miktar" sütunlarını "Kod" sütununun hemen sağına taşı
    cols = list(result_df.columns)

    # Sütunları mevcut yerinden çıkar
    mevcut_eklenenler = [c for c in hedef_sutunlar if c in cols]
    for c in mevcut_eklenenler:
        cols.remove(c)

    # Yerleştirilecek index'i bul
    insert_idx = len(cols)  # Varsayılan en sağ
    if depo_kod_sutunu in cols:
        insert_idx = cols.index(depo_kod_sutunu) + 1
        # Eğer 'Malzeme Adı' sütunu varsa, hammaddeleri onun da sağına alalım
        if "Malzeme Adı" in cols:
            insert_idx = cols.index("Malzeme Adı") + 1

    # Sırayla "Kod" sütununun sağına ekle
    for i, c in enumerate(mevcut_eklenenler):
        cols.insert(insert_idx + i, c)

    result_df = result_df[cols]

    return result_df


def append_carpimis_miktar(matched_df: pd.DataFrame, filtered_df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtrelenmiş ana listeden 'Çarpılmış Miktar' sütununu bulup eşleşen koda göre
    ana tabloya (matched_df) ekler. Eğer aynı koddan çok varsa miktarları toplar.
    """
    depo_kod_sutunu = settings.col_depo_kod  # "Kod"

    if depo_kod_sutunu not in matched_df.columns:
        return matched_df

    mevcut_kod_sutunu = None
    # Sadece_TIM dosyasındaki Kod sütununu görünmez boşluklara aldırış etmeden bul
    for c in filtered_df.columns:
        if str(c).strip().lower() == "kod":
            mevcut_kod_sutunu = c
            break

    if not mevcut_kod_sutunu:
        for c in filtered_df.columns:
            if str(c).strip().lower() == str(settings.col_sira_no_id).strip().lower():
                mevcut_kod_sutunu = c
                break

    if not mevcut_kod_sutunu:
        return matched_df

    carpimis_col = None
    exact_targets = [
        "Çarpılmış Miktar",
        "Carpilmis Miktar",
        settings.col_rezerve_miktar,
        settings.col_kullanilabilir_stok,
    ]
    for target in exact_targets:
        if target in filtered_df.columns:
            carpimis_col = target
            break

    if not carpimis_col:
        return matched_df

    extra_cols = [mevcut_kod_sutunu, carpimis_col]
    if "KAYNAK DOSYA" in filtered_df.columns:
        extra_cols.append("KAYNAK DOSYA")

    df_sub = filtered_df[extra_cols].copy()
    df_sub["_JOIN_KEY"] = df_sub[mevcut_kod_sutunu].astype(str).str.strip().str.upper()

    # Sayısal değere çevirip topla (Kümülatif)
    df_sub[carpimis_col] = pd.to_numeric(
        df_sub[carpimis_col].astype(str).str.replace(",", "."), errors="coerce"
    ).fillna(0)

    def combine_sources(series):
        items = []
        for val in series.dropna():
            val_str = str(val).strip()
            if val_str and val_str not in items:
                items.append(val_str)
        return ", ".join(items)

    agg_dict = {carpimis_col: "sum"}
    if "KAYNAK DOSYA" in extra_cols:
        agg_dict["KAYNAK DOSYA"] = combine_sources

    df_grouped = df_sub.groupby("_JOIN_KEY", as_index=False).agg(agg_dict)

    matched_df["_JOIN_KEY"] = matched_df[depo_kod_sutunu].astype(str).str.strip().str.upper()
    result_df = pd.merge(matched_df, df_grouped, on="_JOIN_KEY", how="left")

    # Sütunu Hammadde'nin yanına taşı
    cols = list(result_df.columns)
    if carpimis_col in cols:
        cols.remove(carpimis_col)
        insert_idx = len(cols)
        if settings.col_hammadde_isim in cols:
            insert_idx = cols.index(settings.col_hammadde_isim) + 1
        elif depo_kod_sutunu in cols:
            insert_idx = cols.index(depo_kod_sutunu) + 1

        cols.insert(insert_idx, carpimis_col)
        result_df = result_df[cols]

    # --- YENİ EKLENTİ: ÜRETİLECEK MİKTAR VE TOPLAM HAMMADDE MİKTARI ---
    hammadde_miktar_col = "Hammadde Miktar"
    uretilecek_miktar_col = "Üretilecek Miktar"
    toplam_miktar_col = "Toplam Hammadde Miktarı"

    if carpimis_col in result_df.columns:
        result_df[uretilecek_miktar_col] = result_df[carpimis_col]

        # Üretilecek Miktar sütununu 'Çarpılmış Miktar'ın hemen sağına al
        cols = list(result_df.columns)
        cols.remove(uretilecek_miktar_col)
        insert_idx = cols.index(carpimis_col) + 1
        cols.insert(insert_idx, uretilecek_miktar_col)
        result_df = result_df[cols]

        if hammadde_miktar_col in result_df.columns:
            # Hesaplamayı artık OpenPyXL ile dinamik formül atayarak yapacağız.
            # Buraya Pandas şimdilik sadece yer tutucu olarak 0 koysun.
            result_df[toplam_miktar_col] = 0

            # Toplam Hammadde Miktarı sütununu 'Hammadde Miktar'ın hemen sağına al
            cols = list(result_df.columns)
            cols.remove(toplam_miktar_col)

            insert_idx = cols.index(hammadde_miktar_col) + 1
            cols.insert(insert_idx, toplam_miktar_col)

            result_df = result_df[cols]

    result_df = result_df.drop(columns=["_JOIN_KEY"])
    return result_df
