import http.server
import json
import os
import re
import shutil
import socketserver
import sys
import tempfile
import threading
import time
import webbrowser

import pandas as pd

# Global değişkenler
global_last_modified = 0.0
excel_path = ""
html_path = ""


def find_latest_filtered_excel():
    """Çıktı klasörlerinde veya ana dizinde en son oluşturulan Filtered_*.xlsx dosyasını bulur."""
    search_paths = ["outputs", "."]
    candidates = []
    for p in search_paths:
        if os.path.exists(p):
            for f in os.listdir(p):
                if f.startswith("Filtered_") and f.endswith(".xlsx") and not f.startswith("~$"):
                    full_path = os.path.abspath(os.path.join(p, f))
                    candidates.append((full_path, os.path.getmtime(full_path)))
    if not candidates:
        return None
    # En son değiştirileni döndür
    return max(candidates, key=lambda x: x[1])[0]


def find_corresponding_html(excel_path_val):
    """Excel dosyasına karşılık gelen HTML ağaç dosyasını bulur."""
    dir_name = os.path.dirname(excel_path_val)
    base_name = os.path.splitext(os.path.basename(excel_path_val))[0]  # örn "Filtered_2229"

    # Aday 1: Doğrudan eşleşme
    c1 = os.path.join(dir_name, f"Makine_Agaci_{base_name}.html")
    if os.path.exists(c1):
        return c1

    # Aday 2: Toplu reçeteler ismi
    c2 = os.path.join(dir_name, "Makine_Agaci_Receteler.html")
    if os.path.exists(c2):
        return c2

    # Aday 3: Filtered_ ön ekini kaldırıp deneme (örn Makine_Agaci_2229.html)
    if base_name.startswith("Filtered_"):
        short_name = base_name.replace("Filtered_", "")
        c3 = os.path.join(dir_name, f"Makine_Agaci_{short_name}.html")
        if os.path.exists(c3):
            return c3

    # Aday 4: Klasördeki herhangi bir Makine_Agaci_*.html dosyası
    for f in os.listdir(dir_name):
        if f.startswith("Makine_Agaci_") and f.endswith(".html"):
            return os.path.abspath(os.path.join(dir_name, f))

    return None


def parse_excel_takip_safe(path):
    """Excel dosyasını kilitleme hatası (PermissionError) olmadan güvenli bir şekilde okur."""
    temp_dir = tempfile.gettempdir()
    temp_path = os.path.join(temp_dir, f"temp_watch_{int(time.time())}.xlsx")

    # Okuma için geçici bir kopya oluştur
    for _ in range(5):
        try:
            shutil.copy2(path, temp_path)
            break
        except Exception:
            time.sleep(0.5)

    try:
        # Excel'deki formül sonuçlarını (data_only=True) almak için openpyxl kullanacağız
        # Pandas read_excel varsayılan olarak cached formül sonuçlarını okur.
        df = pd.read_excel(temp_path, sheet_name="Üretim Takip")

        # Sütunları normalize et
        kod_col = None
        req_col = None
        prod_col = None

        for col in df.columns:
            col_norm = str(col).strip().lower()
            col_norm = (
                col_norm.replace("ı", "i")
                .replace("ş", "s")
                .replace("ç", "c")
                .replace("ğ", "g")
                .replace("ü", "u")
                .replace("ö", "o")
            )
            if col_norm == "kod":
                kod_col = col
            elif col_norm in ["uretilecek miktar", "uretilecek", "gereken miktar", "miktar"]:
                if col_norm != "kod":
                    req_col = col
            elif col_norm in ["uretilen miktar", "uretilen"]:
                prod_col = col

        if not kod_col or not req_col or not prod_col:
            # Sütun isimleri tam eşleşmiyorsa ilk satırlardan tahmin et
            print(
                "[UYARI] Üretim Takip sütunları otomatik çözümlenemedi. Varsayılanlar aranıyor..."
            )
            return None

        data_map = {}
        for _, row in df.iterrows():
            kod = str(row[kod_col]).strip().upper()
            if not kod or kod == "NAN" or kod == "":
                continue

            try:
                req = float(str(row[req_col]).replace(",", "."))
                if pd.isna(req):
                    req = 0.0
            except Exception:
                req = 0.0

            try:
                prod = float(str(row[prod_col]).replace(",", "."))
                if pd.isna(prod):
                    prod = 0.0
            except Exception:
                prod = 0.0

            if kod not in data_map:
                data_map[kod] = {"uretilecek": 0.0, "uretilen": 0.0}
            data_map[kod]["uretilecek"] += req
            data_map[kod]["uretilen"] += prod

        return data_map
    except Exception as e:
        print(f"[HATA] Excel okunamadı: {e}")
        return None
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def update_html_with_excel_data(html_path_val, excel_data):
    """HTML içindeki JSON verisini günceller."""
    if not os.path.exists(html_path_val):
        return False

    try:
        with open(html_path_val, encoding="utf-8") as f:
            html_content = f.read()

        # const allTreesData = [...]; yapısını yakala
        pattern = r"(const allTreesData\s*=\s*)(.*?)(;\s*\n|\s*;\s*//|;\s*$)"
        match = re.search(pattern, html_content)
        if not match:
            # Tek satırda olmayan formatı dene (re.DOTALL ile)
            pattern_dotall = r"(const allTreesData\s*=\s*)(.*?)(;)"
            match = re.search(pattern_dotall, html_content, re.DOTALL)

        if not match:
            print("[HATA] HTML içerisinde 'const allTreesData' verisi bulunamadı.")
            return False

        json_str = match.group(2).strip()
        all_trees = json.loads(json_str)

        # Ağacı rekürsif dolaşıp adetleri güncelle
        def update_node(node):
            kod = str(node.get("kod", "")).strip().upper()
            if kod and kod in excel_data:
                info = excel_data[kod]
                node["uretilecek"] = info["uretilecek"]
                node["uretilen"] = info["uretilen"]

                if info["uretilecek"] > 0:
                    if info["uretilen"] >= info["uretilecek"]:
                        node["prodStatus"] = "green"
                    elif info["uretilen"] > 0:
                        node["prodStatus"] = "yellow"
                    else:
                        node["prodStatus"] = "red"
                else:
                    node["prodStatus"] = "none"

            # Çocukları güncelle
            if "children" in node:
                for child in node["children"]:
                    update_node(child)

        for item in all_trees:
            if "tree" in item:
                update_node(item["tree"])

        # JSON'u geri serileştir
        new_json_str = json.dumps(all_trees, ensure_ascii=False)
        new_html_content = (
            html_content[: match.start(2)] + new_json_str + html_content[match.end(2) :]
        )

        with open(html_path_val, "w", encoding="utf-8") as f:
            f.write(new_html_content)

        return True
    except Exception as e:
        print(f"[HATA] HTML dosyası güncellenirken hata oluştu: {e}")
        return False


def excel_watcher():
    """Excel dosyasını izleyen arka plan döngüsü."""
    global global_last_modified

    last_mtime = os.path.getmtime(excel_path)
    global_last_modified = last_mtime

    # İlk çalıştırmada bir kez güncelle
    excel_data = parse_excel_takip_safe(excel_path)
    if excel_data:
        update_html_with_excel_data(html_path, excel_data)

    print(f"[İZLEYİCİ] '{os.path.basename(excel_path)}' izleniyor...")

    while True:
        try:
            if os.path.exists(excel_path):
                current_mtime = os.path.getmtime(excel_path)
                if current_mtime > last_mtime:
                    # Excel dosyası güncellenmiş!
                    time.sleep(0.5)  # Kaydetme işleminin bitmesini bekle
                    excel_data = parse_excel_takip_safe(excel_path)
                    if excel_data:
                        success = update_html_with_excel_data(html_path, excel_data)
                        if success:
                            last_mtime = current_mtime
                            global_last_modified = current_mtime
                            print(
                                f"[BİLGİ] {time.strftime('%H:%M:%S')} - Excel verileri okundu ve HTML güncellendi."
                            )
        except Exception as e:
            print(f"[UYARI] İzleyici döngüsünde hata: {e}")
        time.sleep(1.0)


class WatcherHTTPHandler(http.server.SimpleHTTPRequestHandler):
    """Yenileme isteklerini ve statik dosyaları yöneten HTTP istek yöneticisi."""

    def log_message(self, format, *args):
        # Konsol kirliliğini önlemek için GET istek loglarını gizle
        pass

    def do_GET(self):
        if self.path == "/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.end_headers()
            response = {"last_modified": global_last_modified}
            self.wfile.write(json.dumps(response).encode("utf-8"))
        else:
            super().do_GET()


def start_server(port=8000):
    """HTTP sunucusunu başlatır."""
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), WatcherHTTPHandler) as httpd:
        print(
            f"[SUNUCU] Canlı izleme adresi: http://localhost:{port}/{os.path.basename(html_path)}"
        )
        # Tarayıcıyı otomatik aç
        webbrowser.open(f"http://localhost:{port}/{os.path.basename(html_path)}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nSunucu kapatılıyor...")
            sys.exit(0)


if __name__ == "__main__":
    print("========================================================")
    print("      REÇETE AĞACI CANLI TAKİP VE İZLEME SERVİSİ")
    print("========================================================")

    # Argüman veya otomatik algılama ile Excel dosyasını bul
    if len(sys.argv) > 1:
        excel_path = os.path.abspath(sys.argv[1])
    else:
        excel_path = find_latest_filtered_excel()

    if not excel_path or not os.path.exists(excel_path):
        print("[HATA] İzlenecek 'Filtered_*.xlsx' dosyası bulunamadı.")
        print("Lütfen önce filtrelemeyi çalıştırın veya dosyayı bu betiğe sürükleyin.")
        input("Çıkmak için Enter'a basın...")
        sys.exit(1)

    # HTML dosyasını bul
    html_path = find_corresponding_html(excel_path)
    if not html_path or not os.path.exists(html_path):
        print(f"[HATA] '{os.path.basename(excel_path)}' için HTML ağaç dosyası bulunamadı.")
        input("Çıkmak için Enter'a basın...")
        sys.exit(1)

    print(f"Hedef Excel: {excel_path}")
    print(f"Hedef HTML : {html_path}")
    print("--------------------------------------------------------")

    # İzleyiciyi arka planda başlat
    watcher_thread = threading.Thread(target=excel_watcher, daemon=True)
    watcher_thread.start()

    # Sunucu dizinini HTML dosyasının bulunduğu yer yap
    os.chdir(os.path.dirname(html_path))

    # Sunucuyu başlat
    start_server(port=8000)
