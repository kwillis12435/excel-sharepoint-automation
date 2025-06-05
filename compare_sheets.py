import pandas as pd
import os

def compare_sheets(auto_csv_path, manual_excel_path, output_path=None, key_columns=None, value_columns=None):
    """
    Compare the automated CSV and manual Excel sheet (only 'Group Results' sheet).
    Outputs a CSV with differences and a summary.
    key_columns: columns to use as unique identifier (e.g. ['study_code', 'trigger', 'gene_target'])
    value_columns: columns to compare (e.g. ['avg_rel_exp', 'avg_rel_exp_lsd', 'avg_rel_exp_hsd'])
    """
    # Load automated CSV
    auto_df = pd.read_csv(auto_csv_path)
    # Load only the 'Group Results' sheet from manual Excel
    manual_df = pd.read_excel(manual_excel_path, sheet_name='Group Results')

    print('Automated CSV columns:', list(auto_df.columns))
    print('Manual Excel columns:', list(manual_df.columns))

    # Default columns if not provided
    if key_columns is None:
        key_columns = ['study_code', 'trigger', 'gene_target']
    if value_columns is None:
        value_columns = ['avg_rel_exp', 'avg_rel_exp_lsd', 'avg_rel_exp_hsd']

    # Check if all key columns exist in both dataframes
    for col in key_columns:
        if col not in auto_df.columns:
            raise KeyError(f"Key column '{col}' not found in automated CSV.")
        if col not in manual_df.columns:
            raise KeyError(f"Key column '{col}' not found in manual Excel sheet.")

    # Merge on key columns
    merged = pd.merge(
        auto_df, manual_df, 
        on=key_columns, 
        how='outer', 
        suffixes=('_auto', '_manual'),
        indicator=True
    )

    # Find differences in value columns
    diff_rows = []
    for idx, row in merged.iterrows():
        diff = {'_merge': row['_merge']}
        for col in key_columns:
            diff[col] = row[col]
        for col in value_columns:
            auto_val = row.get(f'{col}_auto', None)
            manual_val = row.get(f'{col}_manual', None)
            diff[f'{col}_auto'] = auto_val
            diff[f'{col}_manual'] = manual_val
            diff[f'{col}_match'] = pd.isna(auto_val) and pd.isna(manual_val) or auto_val == manual_val
        diff_rows.append(diff)

    diff_df = pd.DataFrame(diff_rows)

    # Save to output CSV if requested
    if output_path:
        diff_df.to_csv(output_path, index=False)
        print(f'Differences written to {output_path}')
    else:
        print(diff_df.head(20))

    # Print summary
    total = len(diff_df)
    exact_matches = diff_df[[f'{col}_match' for col in value_columns]].all(axis=1).sum()
    print(f'Total rows compared: {total}')
    print(f'Rows with all values matching: {exact_matches}')
    print(f'Rows with any difference: {total - exact_matches}')

if __name__ == "__main__":
    # Example usage - update these paths as needed
    auto_csv = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\study_data_01_20250605.csv"  # Path to your automated CSV
    manual_excel = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\LO_study_data (1).xlsx"      # Path to your manual Excel file
    output_csv = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\output.csv"
    compare_sheets(auto_csv, manual_excel, output_csv)
