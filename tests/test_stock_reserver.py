"""
Stok rezervasyon motoru için birim testleri.
"""
import os
import tempfile

import pandas as pd
import pytest

from recipe_automation.utils.stock_reserver import reserve_single_file, load_stok_dict


def make_test_excel(rows: list[dict], path: str) -> None:
    """Verilen satırlarla test Excel dosyası oluşturur."""
    df = pd.DataFrame(rows)
    df.to_excel(path, index=False)


class TestReserveSingleFile:
    """reserve_single_file fonksiyonunun birim testleri."""

    def test_yeterli_stok(self, tmp_path):
        """Stok yeterliyse Aktif Rezerve = İhtiyaç, Rezerve Edilecek = 0."""
        src = str(tmp_path / "test.xlsx")
        make_test_excel([
            {"Kod": "A001", "Rezerve Edilecek Miktar": 5},
            {"Kod": "A002", "Rezerve Edilecek Miktar": 3},
        ], src)

        stok_dict = {"A001": 10.0, "A002": 8.0}
        out_path, stats = reserve_single_file(src, stok_dict, output_suffix="_")

        assert out_path.endswith("_.xlsx")
        df_out = pd.read_excel(out_path)

        a001 = df_out[df_out["Kod"].astype(str) == "A001"].iloc[0]
        assert float(a001["Aktif Rezerve Edilen Miktar"]) == 5.0
        assert float(a001["Rezerve Edilecek Miktar"]) == 0.0

        # Stok düşmeli
        assert stok_dict["A001"] == 5.0
        assert stok_dict["A002"] == 5.0

    def test_stok_yetersiz_kısmi(self, tmp_path):
        """Kısmi stokta Aktif Rezerve = Kalan Stok, Rezerve Edilecek = İhtiyaç - Stok."""
        src = str(tmp_path / "test2.xlsx")
        make_test_excel([
            {"Kod": "B001", "Rezerve Edilecek Miktar": 10},
        ], src)

        stok_dict = {"B001": 4.0}
        out_path, stats = reserve_single_file(src, stok_dict, output_suffix="_")

        df_out = pd.read_excel(out_path)
        b001 = df_out[df_out["Kod"].astype(str) == "B001"].iloc[0]

        assert float(b001["Aktif Rezerve Edilen Miktar"]) == 4.0
        assert float(b001["Rezerve Edilecek Miktar"]) == 6.0
        assert stok_dict["B001"] == 0.0

    def test_stok_yok(self, tmp_path):
        """Stok yoksa Aktif Rezerve = 0, Rezerve Edilecek = İhtiyaç."""
        src = str(tmp_path / "test3.xlsx")
        make_test_excel([
            {"Kod": "C001", "Rezerve Edilecek Miktar": 7},
        ], src)

        stok_dict = {}  # Bu parça stokta yok
        out_path, stats = reserve_single_file(src, stok_dict, output_suffix="_")

        df_out = pd.read_excel(out_path)
        c001 = df_out[df_out["Kod"].astype(str) == "C001"].iloc[0]

        assert float(c001["Aktif Rezerve Edilen Miktar"]) == 0.0
        assert float(c001["Rezerve Edilecek Miktar"]) == 7.0

    def test_oncelik_fifo(self, tmp_path):
        """İki dosya aynı parçaya ihtiyaç duyuyorsa önce gelen daha fazla stok alır."""
        src1 = str(tmp_path / "2254.xlsx")
        src2 = str(tmp_path / "2255.xlsx")

        make_test_excel([{"Kod": "X001", "Rezerve Edilecek Miktar": 8}], src1)
        make_test_excel([{"Kod": "X001", "Rezerve Edilecek Miktar": 5}], src2)

        stok_dict = {"X001": 10.0}

        # Önce 2254 işle
        reserve_single_file(src1, stok_dict, output_suffix="_")
        # Kalan stok: 10 - 8 = 2
        assert stok_dict["X001"] == 2.0

        # Sonra 2255 işle (sadece 2 stok kaldı, 5 lazım)
        out2, _ = reserve_single_file(src2, stok_dict, output_suffix="_")
        df_out = pd.read_excel(out2)
        x001 = df_out[df_out["Kod"].astype(str) == "X001"].iloc[0]

        assert float(x001["Aktif Rezerve Edilen Miktar"]) == 2.0
        assert float(x001["Rezerve Edilecek Miktar"]) == 3.0
        assert stok_dict["X001"] == 0.0

    def test_sifir_ihtyac_satirlari_atlanir(self, tmp_path):
        """İhtiyaç miktarı 0 olan satırlar değiştirilmez."""
        src = str(tmp_path / "test4.xlsx")
        make_test_excel([
            {"Kod": "D001", "Rezerve Edilecek Miktar": 0},
            {"Kod": "D002", "Rezerve Edilecek Miktar": 5},
        ], src)

        stok_dict = {"D001": 100.0, "D002": 3.0}
        reserve_single_file(src, stok_dict, output_suffix="_")

        # D001'in stoku değişmemeli
        assert stok_dict["D001"] == 100.0
        # D002'den 3 çekilmeli
        assert stok_dict["D002"] == 0.0
