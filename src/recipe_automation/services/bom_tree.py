import pandas as pd


def build_bom_tree(df: pd.DataFrame, filename: str = "", initial_reqs: dict = None) -> dict:
    """Excel DataFrame'ini Sira sütununu baz alarak içiçe (nested) JSON (dict) yapısına çevirir.
    İki farklı Excel formatını destekler:
    - Format 1 (KR_FZM tipi): Sira | Kod | Ad | Çarpılmış Miktar | Operasyon-1, Operasyon-2...
    - Format 2 (2241 tipi): Sıra No | Kod | Malzeme | Miktar | 1. Operasyon, 2. Operasyon...
    """

    sira_col = None
    kod_col = None
    ad_col = None
    miktar_col = None
    op_cols = []

    for c in df.columns:
        c_norm = str(c).strip().lower()
        # Türkçe karakterleri normalize et
        c_norm = (
            c_norm.replace("ı", "i")
            .replace("ş", "s")
            .replace("ç", "c")
            .replace("ğ", "g")
            .replace("ü", "u")
            .replace("ö", "o")
            .replace("İ", "i")
        )

        if c_norm in ["sira", "sira no"]:
            sira_col = c
        elif c_norm == "kod":
            kod_col = c
        elif c_norm in ["ad", "malzeme", "malzeme adi", "malzeme adı", "tanim", "tanim"]:
            ad_col = c
        elif c_norm in ["carpilmis miktar", "carpılmıs miktar", "miktar"]:
            miktar_col = c
        elif (
            "operasyon" in c_norm
            and "adi" not in c_norm
            and "sira" not in c_norm
            and "no" not in c_norm
        ):
            op_cols.append(c)

    if not sira_col:
        raise ValueError(
            f"Excel dosyasında hiyerarşi oluşturmak için 'Sira' veya 'Sıra No' sütunu bulunamadı.\n"
            f"Bulunan sütunlar: {list(df.columns)}"
        )

    # Kök Node oluştur — dosya adını başlık olarak kullan
    title = filename if filename else "ANA MAKİNE"
    root = {"name": title, "kod": "ROOT", "ad": title, "miktar": "", "ops": [], "children": []}

    nodes = {"0": root}

    for _, row in df.iterrows():
        sira = str(row.get(sira_col, "")).strip()
        if not sira or sira.lower() == "nan":
            continue

        kod = str(row.get(kod_col, "")).strip() if kod_col else ""
        if kod.lower() == "nan":
            kod = ""

        ad = str(row.get(ad_col, "")).strip() if ad_col else ""
        if ad.lower() == "nan":
            ad = ""

        miktar = row.get(miktar_col, "") if miktar_col else ""
        if pd.isna(miktar):
            miktar = ""
        else:
            try:
                # Türkçe ondalık virgülünü noktaya çevir ("1,0000" → "1.0000")
                miktar_str = str(miktar).strip().replace(",", ".")
                val = float(miktar_str)
                # Tam sayıysa ondalık gösterme (1.0 → "1"), değilse kısa yaz (1.5 → "1.5")
                miktar = str(int(val)) if val == int(val) else f"{val:g}"
            except (ValueError, TypeError):
                miktar = str(miktar)

        ops = []
        for c in op_cols:
            val = str(row.get(c, "")).strip().upper()
            if val and val != "NAN":
                ops.append(val)

        # İlk üretim durumu (ihtiyaç miktarları) eşleştirme
        uretilecek_val = 0.0
        uretilen_val = 0.0
        prod_status = "none"

        kod_upper = kod.strip().upper() if kod else ""
        if initial_reqs is not None and kod_upper:
            if kod_upper in initial_reqs:
                uretilecek_val = initial_reqs[kod_upper]
                prod_status = "red" if uretilecek_val > 0 else "none"

        node = {
            "name": kod if kod else ad,
            "kod": kod,
            "ad": ad,
            "miktar": miktar,
            "sira": sira,
            "ops": ops,
            "uretilecek": uretilecek_val,
            "uretilen": uretilen_val,
            "prodStatus": prod_status,
            "children": [],
        }

        # Parent bulma: "1.1.2" nin parentı "1.1" dir, "1" in parentı "0" dır (Root)
        parts = sira.split(".")

        if len(parts) == 1:
            parent_id = "0"
        else:
            parent_id = ".".join(parts[:-1])

        nodes[sira] = node

        if parent_id not in nodes:
            # Pandas kesurat düzeltmesi: "1.010" → float "1.01" olarak okunabilir
            found = False
            if parent_id.count(".") <= 1:
                try:
                    p_val = float(parent_id)
                    for k in nodes.keys():
                        if k.count(".") <= 1:
                            try:
                                if float(k) == p_val:
                                    parent_id = k
                                    found = True
                                    break
                            except ValueError:
                                pass
                except ValueError:
                    pass

            if not found:
                nodes["0"]["children"].append(node)
            else:
                nodes[parent_id]["children"].append(node)
        else:
            nodes[parent_id]["children"].append(node)

    def paginate_children(node, limit=5):
        if not node.get("children"):
            return

        for child in node["children"]:
            paginate_children(child, limit)

        children = node["children"]
        assemblies = []
        leaves = []

        for child in children:
            if child.get("children") or child.get("is_group"):
                assemblies.append(child)
            else:
                leaves.append(child)

        new_children = assemblies

        if len(leaves) > limit:
            group_node = {
                "name": f"\U0001f4e6 [{len(leaves)} Tekil Par\u00e7ay\u0131 G\u00f6ster]",
                "kod": "GROUP",
                "ad": "",
                "miktar": "",
                "sira": "",
                "ops": [],
                "is_group": True,
                "children": leaves,
            }
            new_children.append(group_node)
        else:
            new_children.extend(leaves)

        node["children"] = new_children

    final_root = root["children"][0] if len(root["children"]) == 1 else root
    paginate_children(final_root, limit=5)
    return final_root
