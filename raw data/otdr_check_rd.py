import pandas as pd
import glob
from statistics import mean
import csv
import os
import sys
import argparse
import re  
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Define constants for attenuation limits
ATTENUATION_LIMIT_1310 = 4.95
ATTENUATION_LIMIT_1550 = 3.93
ATTENUATION_LIMIT_1625 = 4.2

def main():
    args = parse_arguments()
    argv = vars(args)
    path = argv["files"] or os.getcwd()

    change_directory(path)
    filenames = glob.glob(path + "\\*.xlsx")
    
    process_files(filenames)
    convert_csv_to_excel(path)
    check_attenuation_limits(argv)

    print("Done")

def parse_arguments():
    parser = argparse.ArgumentParser(
        prog='OTDR Raw Data Checker',
        description='The program reads raw OTDR data in XLSX format and puts out the cable length and attenuation for each address and wavelength. It also checks if the attenuations are below the threshold value. Command line arguments: python otdr_check_rd.py [--files optional_path_to_directory] [--splices number_of_splices (integer or cell reference)] [--extra additional_attenuation]',
        epilog=''
    )
    parser.add_argument('-f', '--files')
    parser.add_argument('-s', '--splices')
    parser.add_argument('-e', '--extra')
    return parser.parse_args()

def change_directory(path):
    try:
        os.chdir(path)
    except Exception as e:
        print("An error occurred while changing directory: ", str(e))
        sys.exit(1)

def read_workbook(file):
    try:
        return load_workbook(file, data_only=True)
    except Exception as e:
        print(f"An error occurred while reading the workbook {file}: ", str(e))
        return None

def process_files(filenames):
    with open("OTDR.csv", mode="w", newline='') as OTDR_file:
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
            workbook = read_workbook(file)
            if workbook:
                process_workbook(workbook, OTDR_writer, file)

def process_workbook(workbook, OTDR_writer, file):
    sheet = workbook.worksheets[0]
    cable_id = extract_cable_id(sheet)
    lengths, spans_1310, spans_1550, spans_1625 = extract_spans(workbook)
    
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

def extract_cable_id(sheet):
    try:
        if sheet.cell(8, 1).value == "Cable ID":
            return sheet.cell(9, 1).value
        elif sheet.cell(8, 11).value == "Cable ID":
            return sheet.cell(9, 11).value
        else:
            return "None"
    except Exception as e:
        print("An error occurred while reading cable ID: ", str(e))
        return "None"

def extract_spans(workbook):
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
                spans_1310.append(float(span_loss))
            elif "1550" in str(wavelength):
                spans_1550.append(float(span_loss))
            elif "1625" in str(wavelength):
                spans_1625.append(float(span_loss))
        except Exception as e:
            print("An error occurred while reading cable length or span loss: ", str(e))
    return lengths, spans_1310, spans_1550, spans_1625

def convert_csv_to_excel(path):
    try:
        csv_files = glob.glob(path + "\\*.csv")
        for csv_file in csv_files:
            read_file = pd.read_csv(csv_file, encoding="latin-1")
            read_file.to_excel("OTDR_Excel.xlsx", index=None, header=True)
    except Exception as e:
        print("An error occurred while converting CSV to Excel: ", str(e))

def check_attenuation_limits(argv):
    try:
        workbook = load_workbook("OTDR_Excel.xlsx")
        worksheet = workbook.worksheets[0]
        for row in range(2, worksheet.max_row + 1):
            cable_length = worksheet.cell(row=row, column=3).value
            splice_loss = 0.2
            amount_splices = get_amount_splices(argv, worksheet)
            extra = float(argv["extra"]) if argv["extra"] else 0.75
            max_span_1310, max_span_1550, max_span_1625 = calculate_max_spans(cable_length, amount_splices, extra)
            highlight_exceeding_cells(worksheet, row, max_span_1310, max_span_1550, max_span_1625)
        workbook.save("OTDR_Excel_checked.xlsx")
    except Exception as e:
        print("An error occurred during the checking process: ", str(e))

def get_amount_splices(argv, worksheet):
    if argv["splices"]:
        if bool(re.compile(r"^[a-zA-Z][0-9]$").match(argv["splices"])):
            return worksheet.cell(argv['splices']).value
        else:
            try:
                return int(argv["splices"])
            except:
                return 3
    else:
        return 3

def calculate_max_spans(cable_length, amount_splices, extra):
    splice_loss = 0.2
    max_span_1310 = (0.36 * cable_length + 0.45 + 0.7 + extra) + (amount_splices * splice_loss)
    max_span_1550 = (0.21 * cable_length + 0.45 + 0.7 + extra) + (amount_splices * splice_loss)
    max_span_1625 = (0.25 * cable_length + 0.45 + 0.7 + extra) + (amount_splices * splice_loss)
    return max_span_1310, max_span_1550, max_span_1625

def highlight_exceeding_cells(worksheet, row, max_span_1310, max_span_1550, max_span_1625):
    span_1310 = worksheet.cell(row=row, column=6).value
    span_1550 = worksheet.cell(row=row, column=7).value
    span_1625 = worksheet.cell(row=row, column=8).value

    if span_1310 > max_span_1310:
        worksheet.cell(row=row, column=6).fill = PatternFill(start_color="00FF0000", fill_type="solid")
    if span_1550 > max_span_1550:
        worksheet.cell(row=row, column=7).fill = PatternFill(start_color="00FF0000", fill_type="solid")
    if span_1625 > max_span_1625:
        worksheet.cell(row=row, column=8).fill = PatternFill(start_color="00FF0000", fill_type="solid")

if __name__ == '__main__':
    main()
