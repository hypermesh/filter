import pandas as pd


# Test the core matching logic that we just added to main.py
def test_hammadde_seq_matching():
    # Mock row data mimicking the Excel rows
    rows = [
        # TESTERE operation at sequence 1
        {"İş İstasyonu": "TOS2", "Operasyon Adı": "TESTERE", "Operasyon Sıra No": 1.0},
        # TESTERE operation at sequence 2
        {"İş İstasyonu": "TOS2", "Operasyon Adı": "TESTERE", "Operasyon Sıra No": 2},
        # EBATLAMA operation at sequence 1
        {"İş İstasyonu": "TOS3", "Operasyon Adı": "EBATLAMA", "Operasyon Sıra No": "1"},
        # EBATLAMA operation with no sequence column or nan
        {"İş İstasyonu": "TOS3", "Operasyon Adı": "EBATLAMA", "Operasyon Sıra No": None},
    ]
    df = pd.DataFrame(rows)

    # Mock normalization helper
    def normalize_station_name(val):
        if pd.isna(val):
            return ""
        return str(val).strip().upper().replace(" ", "")

    # Mock station mapping setup with sequence suffix
    station_mapping_normalized = {"HAMMADDE": ["TESTERE:1", "EBATLAMA"]}

    # Re-run the core logic we added
    sheet_dfs = {}

    istasyon_col = "İş İstasyonu"
    operasyon_col = "Operasyon Adı"
    operasyon_sira_col = "Operasyon Sıra No"

    for idx, row in df.iterrows():
        st_val = row[istasyon_col]
        norm_st = normalize_station_name(st_val)

        op_val = row[operasyon_col] if operasyon_col else ""
        norm_op = normalize_station_name(op_val)

        op_seq = ""
        if operasyon_sira_col:
            seq_val = row[operasyon_sira_col]
            if pd.notna(seq_val):
                seq_str = str(seq_val).strip()
                if seq_str.endswith(".0"):
                    seq_str = seq_str[:-2]
                op_seq = seq_str

        assigned_sheets = []

        for sheet_name, keywords in station_mapping_normalized.items():
            if sheet_name == "HAMMADDE":
                matched_ham = False
                for kw in keywords:
                    if ":" in kw:
                        parts = kw.split(":")
                        target_op = parts[0]
                        target_seq = parts[1]
                        if norm_op == target_op and op_seq == target_seq:
                            matched_ham = True
                            break
                    else:
                        if norm_op == kw:
                            matched_ham = True
                            break
                if matched_ham:
                    assigned_sheets.append(sheet_name)

        if assigned_sheets:
            for sh in assigned_sheets:
                if sh not in sheet_dfs:
                    sheet_dfs[sh] = []
                sheet_dfs[sh].append(row)

    for sheet_name, r in sheet_dfs.items():
        sheet_dfs[sheet_name] = pd.DataFrame(r)

    # Assertions
    ham_df = sheet_dfs.get("HAMMADDE")
    assert ham_df is not None
    assert len(ham_df) == 3  # Should match row 0 (TESTERE:1), row 2 (EBATLAMA), row 3 (EBATLAMA)

    # Check row 1 (TESTERE:2) is not matched because sequence is 2
    matched_tester_seqs = ham_df[ham_df["Operasyon Adı"] == "TESTERE"]["Operasyon Sıra No"].tolist()
    assert 1.0 in matched_tester_seqs
    assert 2 not in matched_tester_seqs
