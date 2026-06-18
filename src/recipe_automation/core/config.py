from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    debug_mode: bool = False
    default_input_dir: str = "./data"

    # Hedef Sütun İsimleri (Environment Variable ile ezilebilir)
    col_sira_no_id: str = "Sıra No"
    col_rezerve_miktar: str = "Rezerve Edilecek Miktar"

    col_sira_no_stock: str = "Sira"
    col_kullanilabilir_stok: str = "Kullanilabilir Stok"

    col_operasyon_keyword: str = "Operasyon"

    # Veritabanı (DB) Ayarları
    db_dir_name: str = "veritabanlari"
    col_depo_kod: str = "Kod"
    col_hammadde_kod: str = "Hammadde Kod"
    col_hammadde_isim: str = "Hammadde"

    class Config:
        env_file = ".env"


settings = Settings()
