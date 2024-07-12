# Das Programm geht durch alle relevanten Ordner und liest die Rohrnummer und Kabellänge  den Excel-OTDR-Protokollen aus. Die Werte werden anschließend in einer Excel-Datei gespeichert.

import pandas as pd
import glob
from statistics import mean
import csv
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import import argparse

GW_1310 = 4.95
GW_1550 = 3.93
GW_1625 = 4.2

parser = argparse.ArgumentParser(
prog='OTDR Raw Data Checker', description='The program reads raw OTDR data in XLSX format and puts out the cable length and attenuation for each address and wavelength. It also checks if the attenuations are below the threshold value. Command line arguments: python otdr_check_rd.py [--files optional_path_to_directory] [--splices number_of_splices (integer or cell reference)] [--extra additional_attenuation]',epilog='')
parser.add_argument('-f', '--files')
parser.add_argument('-s', '--splices')
parser.add_argument('-e', '--extra')
return parser.parse_args()
argv = vars(args)
try:
    path = argv["files"] or os.getcwd(
    os.chdir(path)
    filenames = glob.glob(path + "\\*.xlsx")
    with open("OTDR.csv", mode="w") as OTDR_file:
        OTDR_writer = csv.writer(OTDR_file, delimiter=",", lineterminator="\r")
        OTDR_writer.writerow(
            [
                "Adresse",
                "Kabelnummer",
                "Mittlere Kabellaenge",
                "Abweichung",
                "Kabellaengen",
                "Spanloss 1310 [dB]",
                "Spanloss 1550",
                "Spanloss 1625",
                "HA-KVz [m]",
            ]
        )
        for file in filenames:
            print(file)
            wb = load_workbook(file, data_only=True)
            sh = wb.worksheets[0]
            try:
                if sh.cell(8, 1).value == "Kabel-ID":
                    rohr = sh.cell(9, 1).value
                elif sh.cell(8, 11).value == "Cable ID":
                    rohr = sh.cell(9, 11).value
                else:
                    rohr = "None"
            except Exception as e:
                print("Ein Fehler ist aufgetreten I: ", str(e))
            laenge = []
            span_1310 = []
            span_1550 = []
            span_1625 = []
            for x in range(1, len(wb.sheetnames)):
                sheet = wb.worksheets[x]
                kabel = sheet.cell(25, 4).value
                span = sheet.cell(25, 10).value
                nm = sheet.cell(19, 1).value
                try:
                    kabel_float = float(kabel)
                    laenge.append(kabel_float)
                    if "1310" in str(nm):
                        try:
                            try:
                                span_1310.append(float(span))
                            except Exception as e:
                                print("Ein Fehler ist aufgetreten A: ", str(e))
                        except Exception as e:
                            print("Ein Fehler ist aufgetreten B: ",str(e))
                    elif "1550" in str(nm):
                        try:
                            try:
                                span_1550.append(float(span))
                            except Exception as e:
                                print("Ein Fehler ist aufgetreten C ", str(e))
                        except Exception as e:
                            print("Ein Fehler ist aufgetreten D ", str(e))
                    elif "1625" in str(nm):
                        try:
                            try:
                                span_1625.append(float(span))
                            except Exception as e:
                                print("Ein Fehler ist aufgetreten E" , str(e))
                        except Exception as e:
                            print("Ein Fehler ist aufgetreten F", str(e))
                except Exception as e:
                    print("Ein Fehler ist aufgetreten G ", str(e))
            if len(span_1625) == 0:
                span_1625.append(0)
            if len(span_1550) == 0:
                span_1550.append(0)
            if len(span_1310) == 0:
                span_1310.append(0)
            OTDR_writer.writerow(
                [
                    #erste Zahl ändern, so dass nur Adresse in Adressspalte steht
                    str(file[78:-4]),
                    rohr,
                    round(mean(laenge), 3),
                    round(max(laenge) - min(laenge), 2),
                    laenge,
                    round(mean(span_1310), 3),
                    round(mean(span_1550), 3),
                    round(mean(span_1625), 3),
                ]
            )
    filenames = glob.glob(path + "\\*.csv")
    for file in filenames:
        read_file = pd.read_csv(file, encoding="latin-1")
        read_file.to_excel("OTDR_Excel.xlsx", index=None, header=True)

except Exception as e:
    print("Ein Fehler ist aufgetreten H ", str(e))

try:
    wb = load_workbook("OTDR_Excel.xlsx")
    ws = wb.worksheets[0]
    for row in range(2, ws.max_row + 1):
        laenge = ws.cell(row=row, column=3).value
        GW_splice = 0.2
        GW_span_1310 = (0.36 * laenge + 0.45 + 0.7 + 0.75) + (3 * GW_splice)
        GW_span_1550 = (0.21 * laenge + 0.45 + 0.7 + 0.75) + (3 * GW_splice)
        GW_span_1625 = (0.25 * laenge + 0.45 + 0.7 + 0.75) + (3 * GW_splice)
        span_1310 = ws.cell(row=row, column=6).value
        span_1550 = ws.cell(row=row, column=7).value
        span_1625 = ws.cell(row=row, column=8).value
        if span_1310 > GW_span_1310:
            ws.cell(row=row, column=6).fill = PatternFill(
                start_color="00FF0000", fill_type="solid"
            )
        if span_1550 > GW_span_1550:
            ws.cell(row=row, column=7).fill = PatternFill(
                start_color="00FF0000", fill_type="solid"
            )
        if span_1625 > GW_span_1625:
            ws.cell(row=row, column=8).fill = PatternFill(
                start_color="00FF0000", fill_type="solid"
            )

    wb.save("OTDR_Excel_geprueft"+ ".xlsx")
except Exception as e:
    print("Ein Fehler ist aufgetreten K ",str(e))

print("Done ")