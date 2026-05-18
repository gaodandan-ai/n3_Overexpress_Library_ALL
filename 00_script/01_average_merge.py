import pandas as pd
import re
import os

def get_well_info(col_name):
    """Extracts well info (Row, Num) like 'A', 1."""
    match = re.search(r'^([A-H])(\d+)$', str(col_name))
    if match:
        return match.group(1), int(match.group(2))
    return None, None

def add_averages_to_df(df):
    """
    Takes a DataFrame, keeps the first two columns, 
    and adds an average column after every two data columns.
    """
    if len(df.columns) <= 2:
        return df
        
    meta_cols = list(df.columns[:2])
    data_cols = list(df.columns[2:])
    
    new_df = df[meta_cols].copy()
    
    for i in range(0, len(data_cols), 2):
        if i + 1 < len(data_cols):
            c1, c2 = data_cols[i], data_cols[i+1]
            new_df[c1] = df[c1]
            new_df[c2] = df[c2]
            
            avg_name = f"{c1}-{c2}_Avg"
            new_df[avg_name] = (pd.to_numeric(df[c1], errors='coerce') + 
                                pd.to_numeric(df[c2], errors='coerce')) / 2
        else:
            c = data_cols[i]
            new_df[c] = df[c]
            
    return new_df

def rename_columns_standard(df):
    """
    Renames data columns sequentially to A1..A12, B1..B12... etc.
    """
    if len(df.columns) < 2:
        return df
        
    meta_cols = list(df.columns[:2])
    rows = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    new_data_names = []
    for r in rows:
        for n in range(1, 13):
            new_data_names.append(f"{r}{n}")
            
    if len(df.columns) == (len(meta_cols) + len(new_data_names)):
        df.columns = meta_cols + new_data_names
    else:
        print(f"Warning: Sheet column count ({len(df.columns)}) doesn't match standard layout. Renaming skipped.")
    
    return df

def process_pair(df1, df2, s1_nums, s2_nums):
    """
    Processes a pair of sheets: extracts columns, adds averages, and removes from originals.
    """
    rows = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    
    # 1. New Split Sheet
    split_data = df1.iloc[:, 0:2].copy()
    cols_to_move_1 = [c for c in df1.columns if get_well_info(c)[1] in s1_nums]
    cols_to_move_2 = [c for c in df2.columns if get_well_info(c)[1] in s2_nums]

    for r in rows:
        for n in s1_nums:
            col = f"{r}{n}"
            if col in df1.columns: split_data[col] = df1[col]
        for n in s2_nums:
            col = f"{r}{n}"
            if col in df2.columns: split_data[col] = df2[col]
    
    split_data_res = rename_columns_standard(add_averages_to_df(split_data))
            
    # 2. Modify Originals
    df1_mod = df1.drop(columns=cols_to_move_1)
    df2_mod = df2.drop(columns=cols_to_move_2)
    
    df1_res = rename_columns_standard(add_averages_to_df(df1_mod))
    df2_res = rename_columns_standard(add_averages_to_df(df2_mod))
    
    return df1_res, split_data_res, df2_res

def process_excel(input_file, output_dir):
    print(f"Processing file: {input_file}")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    file_name = os.path.basename(input_file)
    output_file = os.path.join(output_dir, file_name)

    xl = pd.ExcelFile(input_file)
    sheet_names = xl.sheet_names
    all_dfs = {name: xl.parse(name) for name in sheet_names}
    
    if len(sheet_names) < 4:
        print(f"Error: {file_name} needs at least 4 sheets.")
        return

    # Process Pair 1
    d1_res, m1_res, d2_res = process_pair(all_dfs[sheet_names[0]], all_dfs[sheet_names[1]], [9,10,11,12], [1,2,3,4])
    
    # Process Pair 2
    d3_res, m2_res, d4_res = process_pair(all_dfs[sheet_names[2]], all_dfs[sheet_names[3]], [9,10,11,12], [1,2,3,4])

    # Construct the final sheet list with NEW NAMES requested by user
    # Order: 1(1-4), 1(5-8), 1(9-12), 2(1-4), 2(5-8), 2(9-12), H
    new_names = ["1(1-4)", "1(5-8)", "1(9-12)", "2(1-4)", "2(5-8)", "2(9-12)", "H"]
    
    final_sheets = [
        (new_names[0], d1_res),
        (new_names[1], m1_res),
        (new_names[2], d2_res),
        (new_names[3], d3_res),
        (new_names[4], m2_res),
        (new_names[5], d4_res)
    ]
    
    # Add Sheet5 (if exists) as "H"
    if len(sheet_names) >= 5:
        final_sheets.append((new_names[6], all_dfs[sheet_names[4]]))
    # Add any other remaining sheets with original names (unlikely based on current structure)
    for i in range(5, len(sheet_names)):
        final_sheets.append((sheet_names[i], all_dfs[sheet_names[i]]))

    print(f"Saving result with new sheet names to: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for name, df in final_sheets:
            df.to_excel(writer, sheet_name=name, index=False)

if __name__ == "__main__":
    base_dir = r'F:\Growth_Profiler_Data\n3_Overexpress_Library_ALL'
    
    # Define raw directories and their corresponding result directories
    tasks = [
        (os.path.join(base_dir, '02_raw_data_od'), os.path.join(base_dir, r'04_result\OD\01_average_merge')),
        (os.path.join(base_dir, '03_raw_data_r'), os.path.join(base_dir, r'04_result\R\01_average_merge'))
    ]
    
    for raw_dir, result_dir in tasks:
        print(f"\nScanning raw data directory: {raw_dir}")
        if not os.path.exists(raw_dir):
            print(f"Directory not found: {raw_dir}")
            continue
            
        # Dynamically list all .xlsx files, skipping temporary Excel files starting with ~$
        excel_files = [f for f in os.listdir(raw_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
        
        if not excel_files:
            print(f"No Excel files found in {raw_dir}")
            continue
            
        print(f"Found {len(excel_files)} file(s) to process.")
        for f in excel_files:
            path = os.path.join(raw_dir, f)
            process_excel(path, result_dir)
            
    print("\nAll tasks completed successfully.")
