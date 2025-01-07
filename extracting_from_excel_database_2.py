


import pandas as pd
import argparse
import os
import sys
import openpyxl
import math

HELP = """
_______________________________________________________________________________________
                          | Extracting from Excel Database |             
_______________________________________________________________________________________
Description:
  - This script filters and merges an excel file based on the given arguments.
  - It is advisable to run the following command in terminal before running the script:
      pip install -r requirements.txt.pip

Example of usage:
[WINDOWS]
   python extracting_from_excel_database.py --excel_filename db_template.xlsx --out_dir C:\\Users\\work_folder --out_filename filtered_db.xlsx --time_filter min

[MAC]
   python3 extracting_from_excel_database.py --excel_filename db_template.xlsx --out_dir C:\\Users\\work_folder --out_filename filtered_db.xlsx --time_filter min

Required Arguments:
  --excel_filename [string]           Specify the name of the Excel (xlsx) file or directory.
  --out_dir [directory]               Directory to save output file.
  --out_filename [string]             Name of the output file.
  --time_filter [string]              "min" or "max": min for the earliest, max for the latest.

For more information, refer to the documentation or contact the author andretomasdossantos@gmail.com.

"""



def extract_info_per_excel_sheet(sheet_names, excel_filename, take_first_of_column, columns_to_keep, time_filter):

    # Empty dictionary to store all wanted sheets
    all_data = []

    # Iterate over each Excel Sheet
    for sheet in sheet_names:

        # Read individual Excel sheet
        sheet_data = pd.read_excel(excel_filename, sheet_name=sheet)

        # From the wanted columns, collect which ones are present in the current sheet
        columns_from_this_sheet = [col for col in sheet_data.columns if col in columns_to_keep]

        # For the columns we want to filter, if they are present:
        if take_first_of_column[1] in sheet_data.columns:

            # Make sure dates are formated to datetime dtype
            sheet_data[take_first_of_column[1]] = pd.to_datetime(sheet_data[take_first_of_column[1]]).dt.to_period(freq="D")

            # Sort in ascending order the column present in take_first_of_column[1]
            filtered_data = sheet_data.sort_values(by=take_first_of_column[1])

            # If we want the lattest:
            if time_filter == "max":
                # Given take_first_of_column, keep all take_first_of_column[0] based on the latest of the specified in take_first_of_column[1]
                needed_data = filtered_data.groupby(take_first_of_column[0], as_index=False)[take_first_of_column[1]].max()
            else:
                needed_data = filtered_data.groupby(take_first_of_column[0], as_index=False)[
                    take_first_of_column[1]].min()

            needed_data = needed_data[columns_from_this_sheet].copy()

        else:
            # Keep only the columns who are wanted from this sheet
            needed_data = sheet_data[columns_from_this_sheet]

        # Store each dataframe in a list to append afterwards
        all_data.append(needed_data)

    final_data = pd.concat(all_data, axis=1)

    final_data = final_data.loc[:, ~final_data.columns.duplicated()].copy()

    return final_data


def auto_adjust_columns_and_center(final_data, output_directory):

    final_data = final_data.style.set_properties(**{'text-align': 'center'})

    final_data.to_excel(output_directory, index=False, sheet_name="Filtered")

    # Use openpyxl's load_workbook method to open the Excel file and get the worksheet object
    wb = openpyxl.load_workbook(output_directory)
    ws = wb["Filtered"]

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Use the save method to save the modified Excel file
    wb.save(output_directory)


def get_column_options():
    # Read column options
    columns_information = pd.read_excel("column_options.xlsx")

    columns_to_keep = columns_information['Column names to keep'].tolist()

    dup_column = columns_information['Column name with duplicates'].tolist()
    filter_column = columns_information['Column name to filter duplicates'].tolist()

    clean_dup_column = [x for x in dup_column if str(x) != 'nan']
    clean_filter_column = [x for x in filter_column if str(x) != 'nan']

    take_first_of_column = [clean_dup_column[0],
                            clean_filter_column[0]]

    return columns_to_keep, take_first_of_column


def main(excel_filename, out_dir, out_filename, time_filter):
    print("\n\n\tYour options:")
    print("excel_filename:", excel_filename)
    print("out_dir:", out_dir)
    print("out_filename:", out_filename)
    print("time_filter:", time_filter, "\n\n")

    try:
        os.path.exists(out_dir)
    except ValueError:
        print("Invalid output directory, please insert another one")

    # Read the Excel file to extract the sheet names
    excel_file = pd.ExcelFile(excel_filename)

    # Store sheet_names to iterate
    sheet_names = excel_file.sheet_names

    columns_to_keep, take_first_of_column = get_column_options()

    final_data = extract_info_per_excel_sheet(sheet_names, excel_filename, take_first_of_column, columns_to_keep,
                                              time_filter)

    output_directory = os.path.join(out_dir, out_filename)

    auto_adjust_columns_and_center(final_data, output_directory)


def main_main():
    # Check if help was requested before parsing arguments
    if '-h' in sys.argv or '--Help' in sys.argv:
        # Display your custom help messages
        print(HELP)
        sys.exit()

    parser = argparse.ArgumentParser(description='Process HIV sequence data.', add_help=False)

    # Required arguments
    parser.add_argument('--excel_filename', required=True, help='Specify the name of the Excel (xlsx) file or directory.')
    parser.add_argument('--out_dir', required=True, help='Directory to save output file.')
    parser.add_argument('--out_filename', required=True, help='Name of the output file.')
    parser.add_argument('--time_filter', required=True, help='"min" or "max": min for the earliest, max for the latest.')

    # Custom help option
    parser.add_argument('-h', '--Help', action='store_true', help='Show this help message and exit.')

    args = parser.parse_args()

    # Assign arguments to variables
    excel_filename = args.excel_filename
    out_dir = args.out_dir
    out_filename = args.out_filename
    time_filter = args.time_filter


    # Call the main function with the parsed arguments
    main(excel_filename, out_dir, out_filename, time_filter)


if __name__ == "__main__":
    main_main()