### Script Overview

This Python script automates the processing of OTDR (Optical Time Domain Reflectometer) data from multiple XLSX files in a specified directory. It performs two main tasks: calculating cable lengths and checking for excessive attenuation values.

#### Libraries Used:
- `pandas` for data manipulation
- `glob` for file path handling
- `csv` for CSV file operations
- `os` and `sys` for system-related operations
- `openpyxl` for reading XLSX files

#### Functions:

1. **`main()`**
   - Entry point of the script.
   - Retrieves directory path from command line argument or defaults to current directory.
   - Calls `cable_length()` and `attenuation()` functions, prints formatted results.

2. **`cable_length(path)`**
   - Creates or opens `OTDR.csv` to store address and cable length data.
   - Converts data to XLSX format (`OTDR_cable_length.xlsx`).
   - Deletes temporary CSV file.
   - Returns number of addresses processed.

3. **`attenuation(path)`**
   - Checks attenuation values against predefined thresholds (for 1310nm, 1550nm, 1625nm).
   - Writes addresses with invalid values to `OTDR_attenuation.csv`.
   - Returns number of addresses with invalid values.

4. **`print_result(addresses, invalid)`**
   - Formats and returns a string summarizing the results of cable length and attenuation checks.

#### Execution:
- The script is executed from the command line.
- It expects a directory path as an optional argument (`python script.py [directory]`).
  - If no directory is provided, it defaults to the current working directory.
  
#### Script Outputs:
- After execution, the script outputs a formatted message indicating:
  - The number of cable lengths processed.
  - The number of addresses where at least one attenuation value was found to be too high.

### Example Usage:
```bash
python script.py /path/to/files
