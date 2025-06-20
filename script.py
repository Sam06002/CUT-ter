import pandas as pd
import os

def split_excel_columns(input_file, output_dir, max_columns=4900):
    # Read the Excel file
    df = pd.read_excel(input_file, engine='openpyxl')

    # Get total number of columns
    total_columns = df.shape[1]
    print(f"Total columns: {total_columns}")

    # Calculate how many files are needed
    num_files = (total_columns + max_columns - 1) // max_columns
    print(f"Total files to create: {num_files}")

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Split columns and save to separate files
    for i in range(num_files):
        start_col = i * max_columns
        end_col = min((i + 1) * max_columns, total_columns)

        sub_df = df.iloc[:, start_col:end_col]
        output_file = os.path.join(output_dir, f"split_part_{i + 1}.xlsx")
        sub_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Saved: {output_file}")

# Example usage
input_excel = 'your_input_file.xlsx'  # Replace with your input file path
output_folder = 'split_excel_output'  # Folder to store the split files
split_excel_columns(input_excel, output_folder)
