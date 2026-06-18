import datetime
import os
import sys

import pandas as pd


def merge_excels(folder_path):
    # 1. Klasör kontrolü
    if not os.path.isdir(folder_path):
        print(f"\n[HATA] '{folder_path}' geçerli bir klasör değil.")
        return

    # 2. Excel dosyalarını listele
    try:
        all_files = os.listdir(folder_path)
    except Exception as e:
        print(f"\n[HATA] Klasör içeriği okunamadı: {e}")
        return

    excel_files = [f for f in all_files if f.lower().endswith((".xlsx", ".xls"))]

    # Çıktı dosyasını hariç tut
    output_filename = "BIRLESTIRILMIS_EXCEL.xlsx"
    log_filename = "birlesme_raporu.txt"

    excel_files = [f for f in excel_files if f != output_filename]

    if not excel_files:
        print(f"\n[UYARI] Klasörde birleştirilecek excel dosyası bulunamadı: {folder_path}")
        return

    print(
        f"\nToplam {len(excel_files)} adet Excel dosyası bulundu. Birleştirme işlemi başlatılıyor..."
    )
    print("=" * 70)
    print(f"{'Dosya Adı':<50} | {'Satır Sayısı':<15}")
    print("-" * 70)

    dfs = []
    log_entries = []
    total_rows = 0
    success_count = 0

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        try:
            # header=2 ile ilk 2 satırı (0 ve 1. indeks) atlayıp 3. satırı başlık olarak kullanıyoruz
            df = pd.read_excel(file_path, header=2)
            row_count = len(df)

            print(f"{file_name:<50} | {row_count:<15}")
            log_entries.append((file_name, row_count))

            if row_count > 0:
                # Dosya adını bir kolon olarak eklemek verinin nereden geldiğini izlemek açısından çok faydalıdır.
                # Ancak bunu kullanıcı istemediyse opsiyonel tutabiliriz. Şimdilik sadece veriyi ekleyelim.
                dfs.append(df)
                total_rows += row_count
            success_count += 1
        except PermissionError:
            print(f"{file_name:<50} | [HATA] Kilitli (Açık)")
            log_entries.append((file_name, "HATA: Dosya açık veya kilitli"))
        except Exception as e:
            print(f"{file_name:<50} | [HATA] {str(e)[:15]}")
            log_entries.append((file_name, f"HATA: {e}"))

    print("=" * 70)
    print(
        f"Başarılı: {success_count}/{len(excel_files)} dosya okundu. Toplam veri satırı: {total_rows}"
    )
    print("=" * 70)

    if not dfs:
        print("\n[HATA] Birleştirilecek geçerli veri bulunamadı.")
        return

    # 3. DataFrame'leri birleştir
    try:
        combined_df = pd.concat(dfs, ignore_index=True)
    except Exception as e:
        print(f"\n[HATA] Tablolar birleştirilirken hata oluştu: {e}")
        return

    # 4. Birleştirilmiş Excel'i kaydet
    output_path = os.path.join(folder_path, output_filename)
    try:
        combined_df.to_excel(output_path, index=False)
        print(f"\n[BAŞARILI] Birleştirilmiş Excel kaydedildi: {output_path}")
    except PermissionError:
        print(
            f"\n[HATA] Kayıt Başarısız! '{output_filename}' dosyası şu an açık olduğu için üzerine yazılamıyor."
        )
        print("Lütfen dosyayı kapatıp tekrar deneyin.")
        return
    except Exception as e:
        print(f"\n[HATA] Birleştirilmiş Excel kaydedilemedi: {e}")
        return

    # 5. Rapor/Log dosyasını oluştur
    log_path = os.path.join(folder_path, log_filename)
    try:
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write("========================================================\n")
            log_file.write("             EXCEL BİRLEŞTİRME RAPORU\n")
            log_file.write("========================================================\n")
            log_file.write(f"Tarih: {now}\n")
            log_file.write(f"Klasör: {folder_path}\n")
            log_file.write(f"Oluşturulan Excel: {output_filename}\n")
            log_file.write(f"Toplam Klasördeki Excel Sayısı: {len(excel_files)}\n")
            log_file.write(f"Başarıyla Okunan Dosya Sayısı: {success_count}\n")
            log_file.write(f"Toplam Birleştirilen Satır Sayısı: {total_rows}\n")
            log_file.write("-" * 70 + "\n")
            log_file.write(f"{'Dosya Adı':<50} | {'Satır Sayısı':<15}\n")
            log_file.write("-" * 70 + "\n")
            for name, count in log_entries:
                log_file.write(f"{name:<50} | {count:<15}\n")
            log_file.write("========================================================\n")
        print(f"[BAŞARILI] Rapor dosyası oluşturuldu: {log_path}")
    except Exception as e:
        print(f"[UYARI] Rapor dosyası oluşturulamadı: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("[HATA] Lütfen hedef klasör yolunu argüman olarak verin.")
        sys.exit(1)

    target_folder = sys.argv[1]
    merge_excels(target_folder)
