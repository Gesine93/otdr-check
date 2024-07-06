import pandas as pd
import glob
from statistics import mean
import csv
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Define constants for attenuation limits
ATTENUATION_LIMIT_1310 = 4.95
ATTENUATION_LIMIT_1550 = 3.93
ATTENUATION_LIMIT_1625 = 4.2

def main():
    try:
        # Define the path to the directory containing XLSX files (default is current working directory)
        path = os.getcwd()
        # Optionally, the directory of the files can be given in the command line
        if len(sys.argv) == 2:
            path = sys.argv[1]
        elif len(sys.argv) > 2:
            raise TypeError("Too many arguments. Command only needs the script name and optionally the directory of the files.")
        
        # Change the current working directory to the specified path
        os.chdir(path)
        filenames = glob.glob(path + "\\*.xlsx")
        with ope("OTDR.csv", mode="w", newline='') as OTDR_file:
            OTDR_writer = csv.writer(OTDR_file, delimiter=",", lineterminator="\r")
            OTDR_writer.writerow(
                [
                    "Address",
                    "Cable number",
                    "Average cable length",
                    "Deviation",
                    "Cable lengths",
                    "Span loss 1310 [dB]",
                    "Span loss 1550",
                    "Span loss 1625",
                    "HA-KVz [m]",
                ]
            )
            for file in filenames:
                print(file)
                workbook = load_workbook(file, data_only=True)
                sheet = workbook.worksheets[0]
                try:
                    if sheet.cell(8, 1).value == "Cable ID":
                        cable_id = sheet.cell(9, 1).value
                    elif sheet.cell(8, 11).value == "Cable ID":
                        cable_id = sheet.cell(9, 11).value
                    else:
                        cable_id = "None"
                except Exception as e:
                    print("An error occurred while reading cable ID: ", str(e))
                
                lengths = []
                spans_1310 = []
                spans_1550 = []
                spans_1625 = []
                for x in range(1, len(workbook.sheetnames)):
                    current_sheet = workbook.worksheets[x]
                    cable_length = current_sheet.cell(25, 4).value
                    span_loss = current_sheet.cell(25, 10).value
                    wavelength = current_sheet.cell(19, 1).value
                    try:
                        cable_length_float = float(cable_length)
                        lengths.append(cable_length_float)
                        if "1310" in str(wavelength):
                            try:
                                spans_1310.append(float(span_loss))
                            except Exception as e:
                                print("An error occurred while reading span loss 1310: ", str(e))
                        elif "1550" in str(wavelength):
                            try:
                                spans_1550.append(float(span_loss))
                            except Exception as e:
                                print("An error occurred while reading span loss 1550: ", str(e))
                        elif "1625" in str(wavelength):
                            try:
                                spans_1625.append(float(span_loss))
                            except Exception as e:
                                print("An error occurred while reading span loss 1625: ", str(e))
                    except Exception as e:
                        print("An error occurred while reading cable length: ", str(e))
                
                if len(spans_1625) == 0:
                    spans_1625.append(0)
                if len(spans_1550) == 0:
                    spans_1550.append(0)
                if len(spans_1310) == 0:
                    spans_1310.append(0)
                
                OTDR_writer.writerow(
                    [
                        str(file[78:-4]),
                        cable_id,
                        round(mean(lengths), 3),
                        round(max(lengths) - min(lengths), 2),
                        lengths,
                        round(mean(spans_1310), 3),
                        round(mean(spans_1550), 3),
                        round(mean(spans_1625), 3),
                    ]
                )
        
        # Convert the CSV file to an Excel file
        csv_files = glob.glob(path + "\\*.csv")
        for csv_file in csv_files:
            read_file = pd.read_csv(csv_file, encoding="latin-1")
            read_file.to_excel("OTDR_Excel.xlsx", index=None, header=True)
    
    except Exception as e:
        print("An error occurred: ", str(e))
    
    try:
        # Load the generated Excel file and check attenuation values
        workbook = load_workbook("OTDR_Excel.xlsx")
        worksheet = workbook.worksheets[0]
        for row in range(2, worksheet.max_row + 1):
            cable_length = worksheet.cell(row=row, column=3).value
            splice_loss = 0.2
            max_span_1310 = (0.36 * cable_length + 0.45 + 0.7 + 0.75) + (3 * splice_loss)
            max_span_1550 = (0.21 * cable_length + 0.45 + 0.7 + 0.75) + (3 * splice_loss)
            max_span_1625 = (0.25 * cable_length + 0.45 + 0.7 + 0.75) + (3 * splice_loss)
            span_1310 = worksheet.cell(row=row, column=6).value
            span_1550 = worksheet.cell(row=row, column=7).value
            span_1625 = worksheet.cell(row=row, column=8).value
            
            # Highlight cells with values exceeding the limits
            if span_1310 > max_span_1310:
                worksheet.cell(row=row, column=6).fill = PatternFill(start_color="00FF0000", fill_type="solid")
            if span_1550 > max_span_1550:
                worksheet.cell(row=row, column=7).fill = PatternFill(start_color="00FF0000", fill_type="solid")
            if span_1625 > max_span_1625:
                worksheet.cell(row=row, column=8).fill = PatternFill(start_color="00FF0000", fill_type="solid")
        
        # Save the checked Excel file
        workbook.save("OTDR_Excel_checked.xlsx")
    except Exception as e:
        print("An error occurred during the checking process: ", str(e))
    
    print("Done")

if __name__ == '__main__':
    main()
