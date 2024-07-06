### Script Overview

This Python script processes OTDR (Optical Time Domain Reflectometer) data from multiple XLSX files in a specified directory. It calculates cable lengths, average lengths, span losses for different wavelengths, and checks if attenuation values exceed predefined thresholds.

#### Libraries Used
- `pandas` for data manipulation
- `glob` for file path handling
- `csv` for CSV file operations
- `os` and `sys` for system-related operations
- `argparse` for command-line argument parsing
- `re` for regular expressions
- `openpyxl` for reading and writing XLSX files

#### Functions

1. **`main()`**
   - Entry point of the script.
   - Parses command-line arguments using `argparse`.
   - Defaults to the current working directory if no directory path is provided.
   - Processes XLSX files in the specified directory, calculates cable lengths and span losses, and checks attenuation limits.
   - Converts processed data from CSV to XLSX format.
   - Prints "Done" upon completion.

2. **`parse_arguments()`**
   - Uses `argparse.ArgumentParser` to define and parse command-line arguments:
     - `-f, --files`: Optional path to the directory containing XLSX files.
     - `-s, --splices`: Optional argument for specifying number of splices.
     - `-e, --extra`: Optional additional attenuation value.

3. **`change_directory(path)`**
   - Changes the current working directory to the specified `path`.

4. **`read_workbook(file)`**
   - Loads and returns an openpyxl `Workbook` object from the specified `file`.

5. **`process_files(filenames)`**
   - Processes each XLSX file in `filenames`, extracts cable ID and span losses, and writes data to `OTDR.csv`.

6. **`extract_cable_id(sheet)`**
   - Extracts and returns the cable ID from the given `sheet`.

7. **`extract_spans(workbook)`**
   - Extracts cable lengths and span losses (for different wavelengths) from the given `workbook`.

8. **`convert_csv_to_excel(path)`**
   - Converts all CSV files in `path` to XLSX format (`OTDR_Excel.xlsx`).

9. **`check_attenuation_limits(argv)`**
   - Checks attenuation values against predefined limits (`ATTENUATION_LIMIT_1310`, `ATTENUATION_LIMIT_1550`, `ATTENUATION_LIMIT_1625`).
   - Highlights cells in `OTDR_Excel.xlsx` that exceed the limits.

10. **`get_amount_splices(argv, worksheet)`**
    - Retrieves the number of splices from `argv` or worksheet cell.

11. **`calculate_max_spans(cable_length, amount_splices, extra)`**
    - Calculates maximum allowed span losses based on cable length, splices, and additional attenuation.

12. **`highlight_exceeding_cells(worksheet, row, max_span_1310, max_span_1550, max_span_1625)`**
    - Highlights cells in `worksheet` if span losses exceed calculated maximum values.

#### Execution
- The script is executed from the command line.
- It accepts optional arguments (`-f, --files` for directory path, `-s, --splices` for number of splices, `-e, --extra` for additional attenuation).
- If no -f, -s, or -e options are provided, defaults (os.getcwd(), 3, 0.75) will be used, respectively.
- Outputs "Done" upon successful completion.

#### Script Outputs
- After execution, the script outputs:
  - "Done" message indicating successful completion.
- It creates several output files:
  - `OTDR.csv`: Contains processed data including cable ID, lengths, and span losses.
  - `OTDR_Excel.xlsx`: Converted from CSV files containing processed data.
  - `OTDR_Excel_checked.xlsx`: Final version with highlighted cells indicating exceeding attenuation limits.

### Example Usage
```bash
python otdr_check_rd.py -f /path/to/files -s 5 -e 1.0
