"""Reçete Veritabanı ve Yıllık Tüketim Projeksiyonu Oluşturucu."""

import json
from pathlib import Path
from typing import Any

import pandas as pd


class RecipeDatabaseBuilder:
    def __init__(self, db_dir: Path | None = None):
        self.db_dir = db_dir or Path("veritabanlari")
        self.db_dir.mkdir(parents=True, exist_ok=True)
        self.recete_file = self.db_dir / "parca_makine_receteleri.json"
        self.yillik_file = self.db_dir / "makine_yillik_tahminler.json"

    def parse_combined_excel(self, excel_path: Path) -> dict[str, Any]:
        """Tek birleştirilmiş Excel dosyasından reçeteleri çıkarır."""
        df = pd.read_excel(excel_path)
        col_src = df.columns[0]
        col_code = df.columns[1]
        col_qty = df.columns[2]

        parca_db: dict[str, Any] = {}
        machines = set()

        for _, row in df.iterrows():
            src = str(row[col_src]).strip()
            code = str(row[col_code]).strip().upper()
            qty = float(row[col_qty]) if pd.notna(row[col_qty]) else 0.0

            if not code or code == "NAN":
                continue

            machines.add(src)

            if code not in parca_db:
                parca_db[code] = {
                    "kod": code,
                    "makineler": {},
                }

            if src not in parca_db[code]["makineler"]:
                parca_db[code]["makineler"][src] = qty
            else:
                parca_db[code]["makineler"][src] += qty

        for data in parca_db.values():
            data["makine_sayisi"] = len(data["makineler"])
            data["toplam_birim_adet"] = sum(data["makineler"].values())

        with open(self.recete_file, "w", encoding="utf-8") as f:
            json.dump(parca_db, f, ensure_ascii=False, indent=4)

        if not self.yillik_file.exists():
            yillik = {m: 50 for m in sorted(machines)}
            with open(self.yillik_file, "w", encoding="utf-8") as f:
                json.dump(yillik, f, ensure_ascii=False, indent=4)

        return parca_db

    def calculate_annual_demand(self) -> dict[str, float]:
        """Yıllık makine üretim adetleri ile reçete adetlerini çarpar."""
        if not self.recete_file.exists() or not self.yillik_file.exists():
            return {}

        with open(self.recete_file, encoding="utf-8") as f:
            receteler = json.load(f)

        with open(self.yillik_file, encoding="utf-8") as f:
            yillik_tahminler = json.load(f)

        annual_demand: dict[str, float] = {}

        for code, data in receteler.items():
            total_annual = 0.0
            for m_name, qty in data.get("makineler", {}).items():
                m_annual = yillik_tahminler.get(m_name, 50)
                total_annual += qty * m_annual
            annual_demand[code] = total_annual

        return annual_demand


if __name__ == "__main__":
    builder = RecipeDatabaseBuilder()
    if Path("test.xlsx").exists():
        db = builder.parse_combined_excel(Path("test.xlsx"))
        demand = builder.calculate_annual_demand()
        print(f"Toplam {len(db)} parça için yıllık talep başarıyla hesaplandı.")
