# OTDR PROTOCOL CHECKER

## Description
The OTDR Protocol Checker project consists of two Python scripts designed to automate the processing of OTDR (Optical Time Domain Reflectometer) data. Each script handles different types of data: one for raw data and the other for preprocessed data. They read data from Excel files (XLSX format), extract the relevant information, calculate the actual and allowed attenuation, and identify any values exceeding these predetermined thresholds. The scripts generate XLSX-files summarizing the processed data, indicating the number of cable lengths extracted and identifying addresses with potentially problematic attenuation values.

## Project Files

### otdr_check_raw.py

This script processes raw OTDR Excel files to extract cable lengths, calculate attenuation values, and identify any values exceeding predetermined thresholds. It generates a summary report in an Excel file.
An example XLSX-file can be found in the processed data directory.

### otdr_check_ppd.py
This script processes preprocessed OTDR data. An example XLSX-file can be found in the raw data directory.

### Requirements
The `requirements.txt` files contain all required Python libraries.

To install the required libraries, run:
```
pip install -r requirements.txt
```

Python 3.6 or higher needs to be installed to execute the scripts.

