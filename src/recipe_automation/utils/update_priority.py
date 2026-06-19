import json
import os
import sys
import codecs

def detect_and_read(file_path):
    # UTF-8 with BOM or UTF-16 may need codecs.open
    encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'cp1254', 'latin-1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                content = f.readlines()
            return content
        except (UnicodeDecodeError, LookupError):
            continue
    raise UnicodeDecodeError(f"Dosya kodlaması tespit edilemedi: {file_path}")

def main():
    root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    json_path = os.path.join(root_dir, "veritabanlari", "oncelik_sirasi.json")

    if len(sys.argv) > 1 and str(sys.argv[1]).lower().endswith(".txt"):
        txt_path = os.path.abspath(sys.argv[1])
    else:
        txt_path = os.path.join(root_dir, "oncelik_sirasi.txt")

    if not os.path.exists(txt_path):
        if len(sys.argv) > 1:
            print(f"Hata: Belirttiğiniz '{txt_path}' dosyası sistemde bulunamadı.")
            return
        default_content = (
            "# Bu dosyaya öncelikli kaynak kodlarını sıralı bir şekilde yazın.\n"
            "# Listenin en üstündeki kod en yüksek önceliğe (1) sahip olur.\n\n"
            "2254\n2235\n2273\n2238\n"
        )
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(default_content)
        print(f"Bilgi: '{txt_path}' oluşturuldu. Lütfen düzenleyip tekrar çalıştırın.")
        return

    print(f"'{txt_path}' okunuyor...")
    priorities = {}
    priority_counter = 1

    try:
        lines = detect_and_read(txt_path)
    except Exception as e:
        print(f"Dosya okuma hatası: {e}")
        return

    for line in lines:
        line_clean = line.strip()
        if not line_clean or line_clean.startswith("#"):
            continue
        clean_code = line_clean.replace("'", "").replace('"', "").replace(",", "").strip()
        if clean_code:
            priorities[clean_code] = priority_counter
            priority_counter += 1

    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(priorities, f, indent=4, ensure_ascii=False)

    # Verification step
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if data == priorities:
                print("[DOĞRULAMA] JSON dosyası başarıyla yazıldı ve veriler eşleşiyor.")
            else:
                print("[HATA] Yazılan JSON verisi kaynakla eşleşmiyor!")
    except Exception as e:
        print(f"[HATA] JSON doğrulama sırasında bir hata oluştu: {e}")

    print(f"\n[BAŞARILI] Toplam {len(priorities)} adet kod kaydedildi.")
    print("Oluşturulan Öncelik Sıralaması (İlk 5):")
    for code, prio in list(priorities.items())[:5]:
        print(f"  - Sıra {prio:2d}: {code}")

if __name__ == "__main__":
    main()
