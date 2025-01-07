# Extracting from excel database

## Read me

### Dependencies

This project requires the following Python packages:

- `os` (standard library)
- `argparse` (standard library)
- `os` (standard library)
- `sys` (standard library)
- `pandas`
- `openpyxl`

# Highly advised to create a python virtual environment
### 1 - Create a folder for the project (replace the folder by a chosen one)
```bash
python -m venv ./env_for_testing
```
### 2 - Navigate to the folder and enter the following command to activate the virtual environment
```bash
source ./env_for_testing/bin/activate
```

### 3 - Installation

You can install the required packages using `pip`. It is recommended to use a virtual environment to manage your dependencies.
```bash
pip install -r requirements.txt
```
<br />

# Example of usage
#### Windows
```bash
python extracting_from_excel_database.py --excel_filename db_template.xlsx --out_dir C:\Users\work_folder --out_filename filtered_db.xlsx --time_filter min
```
#### MAC
```bash
python3 extracting_from_excel_database.py --excel_filename db_template.xlsx --out_dir C:\Users\work_folder --out_filename filtered_db.xlsx --time_filter min
```
