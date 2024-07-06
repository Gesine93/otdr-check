# OTDR PROTOCOL CHECKER

## Description
The OTDR Protocol Checker project consists of two Python scripts designed to automate the processing of OTDR (Optical Time Domain Reflectometer) data. Each script handles different types of data: one for raw data and the other for preprocessed data. They read data from Excel files (XLSX format), extract the relevant information, calculate the actual and allowed attenuation, and identify any values exceeding these predetermined thresholds. The scripts generate XLSX-files summarizing the processed data, indicating the number of cable lengths extracted and identifying addresses with potentially problematic attenuation values.

## Project Files

### otdr_check_raw.py

This script processes OTDR Excel files to extract cable lengths, calculate attenuation values, and identify any values exceeding predetermined thresholds. It generates a summary report in an Excel file.

#### Functions

1. **main()**: 
   Orchestrates the entire process, including parsing command-line arguments, changing the directory, processing XLSX files, converting CSV to Excel, and checking attenuation limits.

2. **parse_arguments()**: 
   Parses command-line arguments using `argparse` to specify input directory, number of splices, and additional attenuation.

3. **change_directory(path)**: 
   Changes the current working directory to the specified path.

4. **read_workbook(file)**: 
   Reads and loads an XLSX workbook using `openpyxl.load_workbook()`.

5. **process_files(filenames)**: 
   Iterates through a list of XLSX filenames, reads each workbook, extracts cable IDs, cable lengths, and attenuation values, and writes results to a CSV file.

6. **process_workbook(workbook, OTDR_writer, file)**: 
   Processes a specific workbook to extract cable ID, lengths, and attenuation values for different wavelengths, then writes the summarized data to a CSV file.

7. **extract_cable_id(sheet)**: 
   Extracts the cable ID from a worksheet by checking specific cell values.

8. **extract_spans(workbook)**: 
   Extracts cable lengths and attenuation values for wavelengths 1310nm, 1550nm, and 1625nm from multiple sheets within a workbook.

9. **convert_csv_to_excel(path)**: 
   Converts CSV files in a specified directory to an Excel file.

10. **check_attenuation_limits(argv)**: 
    Loads an Excel file, calculates maximum allowable attenuation values based on cable length, splices, and additional attenuation, then highlights cells exceeding these limits.

11. **get_amount_splices(argv, worksheet)**: 
    Retrieves the number of splices from command-line arguments or defaults to 3 if not specified.

12. **calculate_max_spans(cable_length, amount_splices, extra)**: 
    Calculates maximum allowable attenuation values for wavelengths 1310nm, 1550nm, and 1625nm based on cable length, splices, and additional attenuation.

13. **highlight_exceeding_cells(worksheet, row, max_span_1310, max_span_1550, max_span_1625)**: 
    Highlights cells in an Excel worksheet where attenuation values exceed predefined thresholds.


#### Usage

- Run the script from the command line with the optional directory argument:

```
python otdr_check_rd.py [-f/--files optional_path_to_directory] [-s/--splices number_of_splices] [-e/--extra additional_attenuation]
```

### otdr_check_ppd.py
This script processes preprocessed OTDR data.

#### Functions
**cable_length(path)**
   - Reads the cable lengths from the OTDR Excel files and writes them along with address information to a new Excel file.
   - First, a CSV file with the above information is created. The CSV file is then converted to XLSX format, and the original CSV file is deleted.
   - The cell parameters may need to be adjusted if the length or address information is located in different cells.
   - Returns the number of processed XLSX files.

**attenuation(path)**
   - Reads all the attenuation values from the OTDR Excel files and writes the address information to a CSV file if at least one of the measured attenuations is higher than a threshold value.
   - The cell parameters may need to be adjusted if the attenuation values, the values for calculating the threshold values, or the address information are located in different cells.
   - Returns the number of XLSX files that contain attenuation values higher than the threshold values.

**print_result(addresses, invalid)**
   - Returns a formatted string with the number of addresses and the number of invalid values found.

**main()**
   - This is the entry point of the script.
   - The path to the directory containing XLSX files can be specified in the command line or defaults to the current working directory.
   - Calls `cable_length` and `attenuation` functions and prints the results using `print_result`.

#### Usage
- The path parameter should be set to the directory where all the preprocessed OTDR Excel files are located.
- Run the script from the command line with the optional directory argument:
  ```
  python otdr_check_ppd.py [optional_path_to_directory]
  ```

### requirements.txt
The `requirements.txt` files contain all required Python libraries.

To install the required libraries, run:
```
pip install -r requirements.txt
```

Python 3.6 or higher needs to be installed.

