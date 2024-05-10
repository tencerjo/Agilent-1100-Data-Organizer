# v2
# Added first sort by sequence name, second seqLine
# Add DAD signal for each?
# Must use short or detail report, full does not have area percent
# Test with other column run data

import os
import xlrd
import pandas as pd
import re

def extract_data_from_file(file_path):
    # Open the Excel file
    wb = xlrd.open_workbook(file_path)
    
    # Extract data from Sheet1 (SampleName)
    sheet1 = wb.sheet_by_name('Sheet1')
    
    # Extract data from IntResults sheet
    int_results_sheet = wb.sheet_by_name('IntResults1')

    # Extract data from Signal sheet
    signal_sheet = wb.sheet_by_name('Signal')
    
    # Check the value of cell B2
    n = int(int_results_sheet.cell_value(rowx=1, colx=1))  # Assuming B2 contains the value of n
    
    # Initialize an empty list to store row data
    rows_data = []
    
    # Loop through cells for each value starting from row 1 up to row n
    for row in range(1, n + 1):
        # Extract Method from Sheet1 cell B6
        method = sheet1.cell_value(rowx=5, colx=1)  # B6
        
        # Extract Sequence from Sheet1 cell B5
        sequence_text = sheet1.cell_value(rowx=4, colx=1)  # B5
        sequence = sequence_text.rsplit("\\", 2)[-2]  # Extract text between two backslashes starting from the right

        # Extract File Name from Signal Sheet H2
        file_name_text = signal_sheet.cell_value(rowx=1, colx=7)  # H2
        file_name = file_name_text.rsplit("\\", 2)[-1]
        
        # Initialize dictionary for current row data
        row_data = {}

        # Add a column named 'SeqLine' with value equal to the sample name
        row_data['SeqLine'] = sheet1.cell_value(rowx=18, colx=1)  # B19
        
        # Add a column named 'Sample' with value equal to the sample name
        row_data['Sample'] = sheet1.cell_value(rowx=25, colx=1)  # B26
        
        # Add a column named 'Peak' with value equal to the current loop number n
        row_data['Peak'] = row

        # Check if column 22 has a value, set area percent accordingly
        try:
            area_percent = int_results_sheet.cell_value(rowx=row, colx=22)
        except IndexError:
            area_percent = ""
        row_data['Area_Percent'] = area_percent
        
        # Add other data columns recursively
        for col, label in zip(range(4, 9), ['RetTime', 'Area', 'Height', 'Width', 'Symmetry']):
            row_data[label] = int_results_sheet.cell_value(rowx=row, colx=col)

        
        # Add Sequence, Method, and File Name to row data
        row_data['Sequence'] = sequence
        row_data['File Name'] = file_name
        row_data['Method'] = method

        # Append the row data to the list
        rows_data.append(row_data)

    # Create DataFrame from the list of row data
    int_results_df = pd.DataFrame(rows_data)

    # Convert SeqLine to string for mixed digit numbers
    int_results_df['SeqLine'] = int_results_df['SeqLine'].astype(str)

    return int_results_df

def main():
    # Prompt user to input the parent directory
    parent_directory = input("Enter the parent directory path: ")

    # Initialize lists to store extracted data
    sample_data_list = []
    int_results_df_list = []

    # Recursively search for .xls files in the parent directory and its subdirectories
    for root, dirs, files in os.walk(parent_directory):
        for file in files:
            if file.endswith(".xls"):
                file_path = os.path.join(root, file)
                
                # Extract data from specific cells
                int_results_df = extract_data_from_file(file_path)
                
                # Append the extracted data to the respective lists
                int_results_df_list.append(int_results_df)

    # Concatenate all DataFrames into a single DataFrame
    int_results_df_combined = pd.concat(int_results_df_list, ignore_index=True)
    
    # Define the sorting order
    sorting_order = ['Sequence', 'SeqLine']
    
    # Convert 'SeqLine' to integer type for sorting
    int_results_df_combined['SeqLine'] = int_results_df_combined['SeqLine'].astype(str)
    int_results_df_combined['SeqLine'] = int_results_df_combined['SeqLine'].apply(extract_numeric_part)
    
    # Sort the combined DataFrame by the specified order
    int_results_df_combined.sort_values(by=sorting_order, ascending=[False, True], inplace=True)
    
    # Print the extracted DataFrame
    # print("\nIntResults DataFrame:")
    # print(int_results_df_combined)
    
    # Save IntResults DataFrame to CSV
    file_name = "HPLC_Summary_" + parent_directory.rsplit("\\", 2)[-1] + ".csv"
    try:
        int_results_csv_path = os.path.join(parent_directory, file_name)
        int_results_df_combined.to_csv(int_results_csv_path, index=False)
    except:
        int_results_csv_path = os.path.join(parent_directory, "HPLC_Summary.csv")
        int_results_df_combined.to_csv(int_results_csv_path, index=False)
    
    print(f"Results saved to: {int_results_csv_path}")

if __name__ == "__main__":
    main()

    
# Define a function to extract the numeric part from the SeqLine string
def extract_numeric_part(seqline):
    # Use regular expression to find the numeric part
    match = re.search(r'\d+', seqline)
    if match:
        return int(match.group())  # Return the numeric part as integer
    else:
        return 0  # If no numeric part found, return 0
