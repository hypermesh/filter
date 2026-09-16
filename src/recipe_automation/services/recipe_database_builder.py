"""Reçete Veritabanı ve Yıllık Tüketim Projeksiyonu Oluşturucu."""

import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd


class RecipeDatabaseBuilder:
    def __init__(self, db_dir: Path | None = None):
        self.db_dir = db_dir or Path("veritabanlari")
        self.db_dir.mkdir(parents=True, exist_ok=True)
        self.recete_file = self.db_dir / "parca_makine_receteleri.json"
        self.yillik_file = self.db_dir / "makine_yillik_tahminler.json"

    def load_existing_db(self) -> dict[str, Any]:
        """Mevcut veritabanını okur (yoksa boş döner)."""
        if self.recete_file.exists():
            try:
                with open(self.recete_file, encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
            except Exception:
                return {}
        return {}

    def parse_excel_file(
        self, file_path: Path, parca_db: dict[str, Any], machines: set[str]
    ) -> None:
        """Tek bir Excel dosyasını ayrıştırıp parca_db'ye işler."""
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                return

            col_count = len(df.columns)
            default_machine_name = file_path.stem

            if col_count >= 3 and "Kaynak" in str(df.columns[0]):
                col_src = df.columns[0]
                col_code = df.columns[1]
                col_qty = df.columns[2]

                for _, row in df.iterrows():
                    src = str(row[col_src]).strip()
                    code = str(row[col_code]).strip().upper()
                    qty = float(row[col_qty]) if pd.notna(row[col_qty]) else 0.0

                    if not code or code in ("NAN", "-"):
                        continue

                    machines.add(src)
                    if code not in parca_db:
                        parca_db[code] = {"kod": code, "makineler": {}}

                    parca_db[code]["makineler"][src] = qty
            else:
                col_code = df.columns[0]
                col_qty_name = df.columns[1] if col_count > 1 else None

                machines.add(default_machine_name)
                for _, row in df.iterrows():
                    code = str(row[col_code]).strip().upper()
                    qty = (
                        float(row[col_qty_name])
                        if col_qty_name and pd.notna(row[col_qty_name])
                        else 1.0
                    )

                    if not code or code in ("NAN", "-"):
                        continue

                    if code not in parca_db:
                        parca_db[code] = {"kod": code, "makineler": {}}

                    parca_db[code]["makineler"][default_machine_name] = qty
        except Exception as e:
            print(f"Hata ({file_path.name}): {e}")

    def update_recipes(self, paths: Path | list[Path], incremental: bool = True) -> dict[str, Any]:
        """Excel dosyalarından veya klasörden reçeteleri günceller."""
        parca_db = self.load_existing_db() if incremental else {}
        machines: set[str] = set()

        if isinstance(paths, Path):
            if paths.is_dir():
                file_list = list(paths.glob("*.xlsx")) + list(paths.glob("*.xls"))
            else:
                file_list = [paths]
        else:
            file_list = [p for p in paths if p.is_file()]

        if not file_list:
            print("İşlenecek Excel dosyası bulunamadı.")
            return parca_db

        for fp in file_list:
            print(f"İşleniyor: {fp.name}...")
            self.parse_excel_file(fp, parca_db, machines)

        for data in parca_db.values():
            data["makine_sayisi"] = len(data.get("makineler", {}))
            data["toplam_birim_adet"] = sum(data.get("makineler", {}).values())

        with open(self.recete_file, "w", encoding="utf-8") as f:
            json.dump(parca_db, f, ensure_ascii=False, indent=4)

        yillik: dict[str, Any] = {}
        if self.yillik_file.exists():
            try:
                with open(self.yillik_file, encoding="utf-8") as f:
                    loaded_yillik = json.load(f)
                    if isinstance(loaded_yillik, dict):
                        yillik = loaded_yillik
            except Exception:
                yillik = {}

        for m in machines:
            if m not in yillik:
                yillik[m] = 50

        with open(self.yillik_file, "w", encoding="utf-8") as f:
            json.dump(yillik, f, ensure_ascii=False, indent=4)

        print(f"\nBaşarılı! Toplam {len(parca_db)} parça, {len(yillik)} makine güncellendi.")
        return parca_db


if __name__ == "__main__":
    builder = RecipeDatabaseBuilder()
    if len(sys.argv) > 1:
        target_path = Path(sys.argv[1])
        builder.update_recipes(target_path, incremental=True)
    else:
        recete_dir = Path("Receteler")
        if recete_dir.exists():
            builder.update_recipes(recete_dir, incremental=True)
        elif Path("test.xlsx").exists():
            builder.update_recipes(Path("test.xlsx"), incremental=True)
