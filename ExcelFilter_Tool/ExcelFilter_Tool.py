import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def get_filepath_gui(title):
    """Opens a file dialog window to select an Excel file."""
    root = tk.Tk()
    root.withdraw()  # Hide the small empty window
    filepath = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    return filepath

def get_savepath_gui():
    """Opens a file dialog to specify the save location and name for the output."""
    root = tk.Tk()
    root.withdraw() # Hide the small empty window
    filepath = filedialog.asksaveasfilename(
        title="Save the filtered results as...",
        defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")]
    )
    return filepath

def get_valid_column(df, prompt_message, filename):
    """Continuously asks the user for a column name until a valid one is entered."""
    print(f"\nAvailable columns in '{filename}': {list(df.columns)}")
    while True:
        column_name = input(prompt_message).strip()
        if column_name in df.columns:
            return column_name
        else:
            print(f"--- Error: Column '{column_name}' not found. Please choose from the list above. ---")

def main():
    """Main function to run the interactive Excel filtering process."""
    print("--- Interactive Excel Filtering Tool ---")
    print("This script will help you filter a main Excel file based on lists from other files.")
    print("-" * 50)

    # --- Step 1: Select the main Excel file to filter ---
    print("A window will now open for you to select the main Excel file...")
    main_file_path = get_filepath_gui("Select the Main Excel File to Filter")
    
    if not main_file_path:
        print("--- No file selected. Exiting the program. ---")
        return
        
    try:
        main_df = pd.read_excel(main_file_path)
        print(f"Successfully loaded '{os.path.basename(main_file_path)}' with {len(main_df)} rows.")
    except Exception as e:
        print(f"--- Error: Could not read the Excel file. Reason: {e} ---")
        return

    # --- Step 2: Get the number of columns to filter by ---
    while True:
        try:
            num_filters = int(input("\nHow many columns would you like to apply a filter on? "))
            if num_filters > 0:
                break
            else:
                print("--- Please enter a number greater than zero. ---")
        except ValueError:
            print("--- Invalid input. Please enter a whole number. ---")

    # --- Step 3: Loop through each filter ---
    for i in range(num_filters):
        print(f"\n--- Configuring Filter #{i+1} ---")

        # Get the column to filter in the main dataframe
        target_column = get_valid_column(main_df, f"Enter the name of the column to filter in your main file: ", os.path.basename(main_file_path))

        # Select the Excel file containing the filter list
        print(f"A window will open for you to select the Excel file for Filter #{i+1}...")
        filter_file_path = get_filepath_gui(f"Select Filter List File #{i+1}")

        if not filter_file_path:
            print("--- File selection cancelled. Skipping this filter. ---")
            continue

        try:
            # Read the filter file assuming it has NO header.
            # This treats the first row as data.
            filter_df = pd.read_excel(filter_file_path, header=None)
            # Use the first column (index 0) for the filter list.
            filter_list = filter_df[0].dropna().astype(str).tolist()
            print(f"Using all {len(filter_list)} items from the first column of '{os.path.basename(filter_file_path)}'.")
        except Exception as e:
            print(f"--- Error: Could not read the filter list file. Reason: {e} ---")
            continue
        
        if not filter_list:
            print(f"--- Warning: The filter list from '{os.path.basename(filter_file_path)}' is empty. Skipping this filter. ---")
            continue

        # Ask for the filter method
        while True:
            method = input("Choose the filter method ('exact' or 'contains'): ").strip().lower()
            if method in ['exact', 'contains']:
                break
            else:
                print("--- Invalid method. Please type 'exact' or 'contains'. ---")

        # Apply the filter
        initial_rows = len(main_df)
        if method == 'exact':
            main_df = main_df[main_df[target_column].astype(str).isin(filter_list)]
        elif method == 'contains':
            regex_pattern = '|'.join(filter_list)
            main_df = main_df[main_df[target_column].astype(str).str.contains(regex_pattern, na=False, case=False)]
        
        print(f"Filter applied. Number of rows changed from {initial_rows} to {len(main_df)}.")

    # --- Step 4: Save the final result ---
    print("-" * 50)
    if main_df.empty:
        print("\nNo data matched your filtering criteria. The output file will not be created.")
    else:
        print(f"\nFiltering complete. The final dataset has {len(main_df)} rows.")
        print("A final window will open for you to choose where to save the result...")
        output_filename = get_savepath_gui()

        if not output_filename:
            print("--- Save operation cancelled. The filtered data was not saved. ---")
            return

        try:
            main_df.to_excel(output_filename, index=False)
            print(f"\nSuccess! The filtered data has been saved to '{output_filename}'.")
        except Exception as e:
            print(f"--- Error: Could not save the file. Reason: {e} ---")

if __name__ == "__main__":
    main()
