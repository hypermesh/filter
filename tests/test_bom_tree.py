import pandas as pd
from recipe_automation.services.bom_tree import build_bom_tree

def test_build_bom_tree_with_initial_reqs():
    # Basit bir reçete verisi oluştur
    data = {
        "Sira": ["1", "1.1", "1.1.1"],
        "Kod": ["KOD-A", "KOD-B", "KOD-C"],
        "Malzeme": ["M-A", "M-B", "M-C"],
        "Miktar": [1, 2, 3]
    }
    df = pd.DataFrame(data)
    
    initial_reqs = {
        "KOD-B": 5.0,
        "KOD-C": 0.0
    }
    
    tree = build_bom_tree(df, filename="TEST_MACH", initial_reqs=initial_reqs)
    
    # Kök düğüm kontrolü (Eğer tek kök düğüm varsa, build_bom_tree kökü o düğüm yapar)
    assert tree["kod"] == "KOD-A"
    
    # Çocukları kontrol et
    children1 = tree["children"]
    assert len(children1) > 0
    node_b = children1[0]
    assert node_b["kod"] == "KOD-B"
    assert node_b["uretilecek"] == 5.0
    assert node_b["uretilen"] == 0.0
    assert node_b["prodStatus"] == "red"
    
    node_c = node_b["children"][0]
    assert node_c["kod"] == "KOD-C"
    assert node_c["uretilecek"] == 0.0
    assert node_c["uretilen"] == 0.0
    assert node_c["prodStatus"] == "none"
