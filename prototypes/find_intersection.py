import pandas as pd
from pathlib import Path
from difflib import SequenceMatcher

from format_workbook import format_workbook

def read_client_trial_balance(path: str | Path, sheet_name="Trial Balance") -> pd.DataFrame:
    """
    Read starting where it has debit and credit in the same line to get client
    balances.
    """
    path = Path(path)
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    header_row = None
    for i, row in raw.iterrows():
        row_str = row.astype(str)
        if row_str.str.contains("Debit", case=False, na=False).any() and row_str.str.contains("Credit", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find the header row containing Debit and Credit in the client file.")

    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "account"})
    df = df[~df["account"].isna()].copy()

    df["account"] = df["account"].astype(str).str.strip()
    df["Debit"] = pd.to_numeric(df.get("Debit"), errors="coerce").fillna(0.0)
    df["Credit"] = pd.to_numeric(df.get("Credit"), errors="coerce").fillna(0.0)

    df["balance_client"] = df["Debit"] - df["Credit"]
    df["account_key"] = df["account"].str.lower().str.replace(r"\s+", " ", regex=True)

    return df[["account", "account_key", "Debit", "Credit", "balance_client"]]

def read_import_trial_balance(path: str | Path, sheet_name="Trial Balance") -> pd.DataFrame:
    """
    Look at the spreadsheet, assign header names, and then parse through data
    """
    path = Path(path)
    df = pd.read_excel(path, sheet_name=sheet_name, header=None)

    if df.shape[1] < 5:
        raise ValueError(f"Import file has {df.shape[1]} columns, expected at least 5.")

    df = df.iloc[:, :5].copy()
    df.columns = ["class", "acct_no", "account", "col3", "balance_import"]

    df["account"] = df["account"].astype(str).str.strip()
    df["acct_no"] = pd.to_numeric(df["acct_no"], errors="coerce")
    df["col3"] = pd.to_numeric(df["col3"], errors="coerce").fillna(0.0)
    df["balance_import"] = pd.to_numeric(df["balance_import"], errors="coerce").fillna(0.0)

    df["account_key"] = df["account"].str.lower().str.replace(r"\s+", " ", regex=True)

    return df[["class", "acct_no", "account", "account_key", "col3", "balance_import"]]

def top_matches(source_keys, target_keys, top_n=5, min_ratio=0.70) -> pd.DataFrame:
    target_list = list(target_keys)
    rows = []
    for s in source_keys:
        scored = []
        for t in target_list:
            r = SequenceMatcher(None, s, t).ratio()
            if r >= min_ratio:
                scored.append((r, t))
        scored.sort(reverse=True, key=lambda x: x[0])
        for r, t in scored[:top_n]:
            rows.append({"source_account_key": s, "candidate_account_key": t, "similarity": round(r, 3)})
    return pd.DataFrame(rows)

def compare_trial_balances(
    client_path: str | Path,
    import_path: str | Path,
    out_path: str | Path = "trial_balance_comparison.xlsx",
    sheet_name="Trial Balance",
    tol=0.01
) -> Path:
    client_df = read_client_trial_balance(client_path, sheet_name=sheet_name)
    import_df = read_import_trial_balance(import_path, sheet_name=sheet_name)

    matched = client_df.merge(import_df, on="account_key", how="inner", suffixes=("_client", "_import"))
    matched["balance_diff"] = matched["balance_client"] - matched["balance_import"]

    mismatched = matched[matched["balance_diff"].abs() > tol].copy()

    only_client = client_df[~client_df["account_key"].isin(import_df["account_key"])].copy()
    only_import = import_df[~import_df["account_key"].isin(client_df["account_key"])].copy()

    suggest_client_to_import = top_matches(
        source_keys=only_client["account_key"].unique(),
        target_keys=set(only_import["account_key"].unique()),
        top_n=5,
        min_ratio=0.70
    )

    suggest_import_to_client = top_matches(
        source_keys=only_import["account_key"].unique(),
        target_keys=set(only_client["account_key"].unique()),
        top_n=5,
        min_ratio=0.70
    )

    summary = pd.DataFrame([
        {"metric": "client rows", "value": len(client_df)},
        {"metric": "import rows", "value": len(import_df)},
        {"metric": "exact matches", "value": len(matched)},
        {"metric": "value mismatches", "value": len(mismatched)},
        {"metric": "only in client", "value": len(only_client)},
        {"metric": "only in import", "value": len(only_import)},
    ])

    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="summary")
        matched.sort_values("account_key").to_excel(writer, index=False, sheet_name="matched")
        mismatched.sort_values("account_key").to_excel(writer, index=False, sheet_name="mismatched")
        only_client.sort_values("account_key").to_excel(writer, index=False, sheet_name="only_in_client")
        only_import.sort_values("account_key").to_excel(writer, index=False, sheet_name="only_in_import")
        suggest_client_to_import.to_excel(writer, index=False, sheet_name="suggest_client_to_import")
        suggest_import_to_client.to_excel(writer, index=False, sheet_name="suggest_import_to_client")

    return out_path

if __name__ == "__main__":
    compare_trial_balances(
        client_path="Accounting_Project\\client tb.xlsx",
        import_path="Accounting_Project\\tb to import.xlsx",
        out_path="Accounting_Project\\trial_balance_comparison.xlsx"
    )
    print("Done")

    format_workbook("Accounting_Project\\trial_balance_comparison.xlsx")