# OTDR PROTOCOL CHECKER
## Description
The Python script automates the processing of OTDR (Optical Time Domain Reflectometer) data, focusing on cable lengths and attenuation values. \ 
It reads data from Excel files, extracts relevant information, performs calculations based on specified formulas, and identifies any values exceeding predetermined thresholds. \  The script then generates reports summarizing the processed data, indicating the number of cable lengths extracted and identifying addresses with potentially problematic attenuation values.
## Project Files
### project.py
Contains the code. \
Three functions plus a main-function were created.\
The path parameter in the main-function has to be set to the path where all the OTDR excel files are located.
### Functions
The first one (cable_length) reads the cable lenghts from the OTDR excel files and writes them and the adress information to a new excel file.\
At first a csv-file with the information above is created. Afterwards the csv-file is converted to xlsx and lastly the csv-file gets deleted.\
I chose this way because it's easier to create a csv-file first.\
The cell parameters have to be changed, when the length or/and adress information are located in another cell.\
The function returns the number of xlsx-files that were processed.\
The second one (attenuation) reads all the attenuation values from the OTDR excel files and writes the adress information to a csv file if at least one of the measured attenuations is higher than a threshold value.\
The cell parameters have to be changed if the attenuation values or values for calculating the threshold values or the adress information are located in other cells.\
The function returns the number of xlsx-files that contain attenuation values higher than the threshold values.\
The last one (print_result) prints the number of excel files read and the number of excel files, where at least one attenuation value was too high to the console.\
### otdr_check_ppd.py
The file contains three functions to test the three functions of the project.py file.\
Since the virtual VS Code enviroment doesn't display the xlsx-files the correct way and seems to be unable to handle them in general, the functions attenuation() and cable_length() return 0.\
The print_result() functions should return a formatted string including the two input parameters adressen and invalid.
### requirements.txt
The txt-file contains all required python libraries.\
Required are os, sys, openpyxl, glob, csv and pandas.

