import json
import os
import sys


def main():
    # Proje kök dizinini bul
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    json_path = os.path.join(root_dir, "veritabanlari", "oncelik_sirasi.json")

    # Sürükle-bırak desteği: Eğer argüman olarak bir .txt dosyası verilmişse onu kullan
    if len(sys.argv) > 1 and str(sys.argv[1]).lower().endswith(".txt"):
        txt_path = os.path.abspath(sys.argv[1])
    else:
        txt_path = os.path.join(root_dir, "oncelik_sirasi.txt")

    if not os.path.exists(txt_path):
        if len(sys.argv) > 1:
            print(f"Hata: Sürüklenen dosya bulunamadı: '{txt_path}'")
            return
        # Eğer varsayılan dosya yoksa örnek şablon oluştur
        default_content = (
            "# Bu dosyaya öncelikli kaynak kodlarını sıralı bir şekilde yazın.\n"
            "# Her satıra tırnaksız, virgülsüz tek bir kod yazılmalıdır.\n"
            "# Listenin en üstündeki kod en yüksek önceliğe (1) sahip olur.\n"
            "# '#' ile başlayan satırlar açıklama satırıdır ve yoksayılır.\n\n"
            "2254\n"
            "2235\n"
            "2273\n"
            "2238\n"
        )
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(default_content)
        print(f"Bilgi: '{txt_path}' dosyası bulunamadı. Örnek bir şablon oluşturuldu.")
        print("Lütfen bu dosyayı düzenleyip komutu tekrar çalıştırın.")
        return

    print(f"'{txt_path}' dosyası okunuyor...")
    priorities = {}
    priority_counter = 1

    with open(txt_path, encoding="utf-8") as f:
        for line in f:
            line_clean = line.strip()
            # Boş veya yorum satırlarını atla
            if not line_clean or line_clean.startswith("#"):
                continue

            # Tırnak, çift tırnak, virgül gibi karakterleri temizle
            clean_code = line_clean.replace("'", "").replace('"', "").replace(",", "").strip()
            if clean_code:
                priorities[clean_code] = priority_counter
                priority_counter += 1

    # Dizini oluştur
    os.makedirs(os.path.dirname(json_path), exist_ok=True)

    # JSON dosyasına yaz (üzerine yazar)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(priorities, f, indent=4, ensure_ascii=False)

    print(f"\n[BAŞARILI] Toplam {len(priorities)} adet kod öncelik sırasına göre kaydedildi.")
    print("Hedef dosya güncellendi: 'veritabanlari/oncelik_sirasi.json'")

    print("\nOluşturulan Öncelik Sıralaması (İlk 10):")
    for code, prio in list(priorities.items())[:10]:
        print(f"  - Sıra {prio:2d}: {code}")
    if len(priorities) > 10:
        print(f"  - ... ve diğer {len(priorities) - 10} kod.")


if __name__ == "__main__":
    main()
