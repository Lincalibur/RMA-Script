import os
import pandas as pd

def load_missing_list(missing_list_path, missing_sheet_name):
    print(f"Loading missing list from: {missing_list_path}")
    missing_df = pd.read_excel(missing_list_path, sheet_name=missing_sheet_name)
    missing_df['Tracking No.'] = missing_df['Tracking No.'].astype(str)
    missing_df['Serial No.'] = missing_df['Serial No.'].astype(str)
    print(f"Loaded {len(missing_df)} records from the missing list.")
    print(f"Missing list columns: {missing_df.columns.tolist()}")
    return missing_df

def search_folder_for_numbers(folder_path, search_column, numbers_set, created_at_column):
    found_records = []
    print(f"Searching for numbers in folder: {folder_path}")
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            df[search_column] = df[search_column].astype(str)
            df[created_at_column] = pd.to_datetime(df[created_at_column])  # Ensure the Created At column is datetime
            matched_records = df[df[search_column].isin(numbers_set)]
            if not matched_records.empty:
                # Drop duplicates and keep the latest Created At entry
                matched_records = matched_records.sort_values(by=created_at_column).drop_duplicates(subset=search_column, keep='last')
                found_records.append(matched_records)
    found_df = pd.concat(found_records, ignore_index=True) if found_records else pd.DataFrame()
    print(f"Found {len(found_df)} records in {folder_path}.")
    return found_df

def update_missing_list(missing_df, found_df, match_column):
    initial_count = len(missing_df)
    matched_entries = missing_df[missing_df[match_column].isin(found_df[match_column])]
    found_count = len(matched_entries)
    updated_df = missing_df[~missing_df.index.isin(matched_entries.index)]
    print(f"Updated missing list. {found_count} records were found and removed.")
    print(f"Records before update: {initial_count}, after update: {len(updated_df)}")
    return updated_df, matched_entries

def save_to_excel(df, path, sheet_name):
    with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Saved {len(df)} records to {path} under sheet name {sheet_name}.")

def load_bc_list_data(folder_path):
    data_frames = []
    print(f"Loading data from folder: {folder_path}")
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            data_frames.append(df)
    return pd.concat(data_frames, ignore_index=True) if data_frames else pd.DataFrame()

def main():
    missing_list_path = 'Missing_RMA/Social Mobile Missing RMAs 5-22-24.xlsx'
    tracking_folder_path = 'BC_List_Tracking'
    imei_folder_path = 'BC_List_IMEI'
    
    found_tracking_path = 'Found_Tracking.xlsx'
    found_imei_path = 'Found_IMEI.xlsx'
    not_found_path = 'Not_Found.xlsx'
    
    missing_sheet_name = 'Missing'  # Update this if necessary
    created_at_column = 'Created At'  # Column name for created at timestamp

    # Load Missing List
    missing_df = load_missing_list(missing_list_path, missing_sheet_name)
    initial_missing_count = len(missing_df)
    
    # Search for Tracking Numbers
    tracking_numbers = set(missing_df['Tracking No.'])
    found_tracking_df = search_folder_for_numbers(tracking_folder_path, 'Tracking No.', tracking_numbers, created_at_column)
    
    # Update Missing List for Tracking Numbers
    missing_df, found_tracking_entries = update_missing_list(missing_df, found_tracking_df, 'Tracking No.')
    
    # Save found tracking entries
    save_to_excel(found_tracking_entries, found_tracking_path, 'Found_Tracking')

    # Remove corresponding IMEI numbers for found tracking entries
    found_imei_numbers = set(found_tracking_entries['Serial No.'])
    remaining_imei_numbers = set(missing_df['Serial No.']) - found_imei_numbers
    
    print(f"Remaining IMEI numbers count: {len(remaining_imei_numbers)}")
    
    # Search for IMEI Numbers
    found_imei_df = search_folder_for_numbers(imei_folder_path, 'Serial No.', remaining_imei_numbers, created_at_column)
    
    # Update Missing List for IMEI Numbers
    missing_df, found_imei_entries = update_missing_list(missing_df, found_imei_df, 'Serial No.')
    
    # Save found IMEI entries
    save_to_excel(found_imei_entries, found_imei_path, 'Found_IMEI')
    
    final_missing_count = len(missing_df)
    
    print(f"Initial missing count: {initial_missing_count}")
    print(f"Total found (tracking + IMEI): {len(found_tracking_entries) + len(found_imei_entries)}")
    print(f"Final missing count: {final_missing_count}")
    print(f"Sum (initial - found): {initial_missing_count - (len(found_tracking_entries) + len(found_imei_entries))}")

    # Save Not Found List
    save_to_excel(missing_df, not_found_path, 'Not_Found')

    print("Process complete.")

    # Load BC List IMEI and BC List Tracking Data
    bc_list_tracking_df = load_bc_list_data(tracking_folder_path)
    bc_list_imei_df = load_bc_list_data(imei_folder_path)

    # Convert 'Tracking No.' column to string in both DataFrames
    found_tracking_entries['Tracking No.'] = found_tracking_entries['Tracking No.'].astype(str)
    bc_list_tracking_df['Tracking No.'] = bc_list_tracking_df['Tracking No.'].astype(str)

    # Merge Found Tracking with BC List Tracking
    found_tracking_merged_df = pd.merge(found_tracking_entries, bc_list_tracking_df, on='Tracking No.', how='left')
    save_to_excel(found_tracking_merged_df, 'Found_Tracking_Merged.xlsx', 'Found_Tracking_Merged')

    # Convert 'Serial No.' column to string in both DataFrames
    found_imei_entries['Serial No.'] = found_imei_entries['Serial No.'].astype(str)
    bc_list_imei_df['Serial No.'] = bc_list_imei_df['Serial No.'].astype(str)

    # Merge Found IMEI with BC List IMEI
    found_imei_merged_df = pd.merge(found_imei_entries, bc_list_imei_df, on='Serial No.', how='left')
    save_to_excel(found_imei_merged_df, 'Found_IMEI_Merged.xlsx', 'Found_IMEI_Merged')

    print("Process complete.")

if __name__ == "__main__":
    main()
