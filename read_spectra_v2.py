import os
import pathlib
import re
import glob as gl
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import numpy as np

"""
Created on Thurs Nov 26, 2020 (rzala)
"""

"""
This script is intended to read the spectra files ('.ISD') saved by Specwin Pro
It iterates through a folder of these files
For each file it extracts the source current, and the spectrum data
It adds the spectrum data to a dataframe (i.e. a table)
It then writes this dataframe to an excel file 
(the intent is that the contents of the excel file can be copied into a sheets template)

The dataframe is structured as:
Wavelength | 0 mA | 0.001 mA | ...etc

i.e. it contains one column for wavelength, and each subsequent column contains the measured power (W/nm).  The 
header for each column is the source current to the LED 

--------------------------------------------------------------------------------------------------------------------
Inputs: 
data_folder                 - path to folder where all ".ISD" files for a single sweep are stored (string)
important_current_values    - list of source current values (list of floats, unit is mA)
                              if specified will create a separate sheet where the table only contains those columns
                              if not specified, defaults to taking first, center, and last columns

save_folder                 - path to folder to save the excel file in (string)
save_filename               - name of excel file (string)
---------------------------------------------------------------------------------------------------------------------
"""

# ==================================================================================================================
# INPUTS:
data_folder = os.getcwd() + \
              '\\test_data\\B2-10um_2019-08-11_Green_Prolux_Ref_Reflector\\LED1_P1N2'

important_current_values = [0.001, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5]
important_current_values_scientific = ["{:.2E}".format(value/1000) for value in important_current_values]

save_folder = data_folder
save_filename = [re.split(r'\\', data_folder)][0][-1] + ' Spectra'


# [re.split(r'\\', data_folder)][0][-1] gets the name of innermost data folder
# e.g. data_folder = "...\\B2_100um\\LED1_P1N2", save_filename = "LED1_P1N2 ..."

# ===================================================================================================================
# FUNCTIONS
def find_line(regexp, data_list):
    """
    Returns line # of line containing a search term

    params:
    regexp - regex expression/string to search for
    data_list - list where each entry is a line from the '.ISD' file

    returns: index
    if a single index is found it returns that index
    if nothing is found raises an exception
    if multiple are found, it asks the user
    """
    indices = [i for i, line in enumerate(data_list) if re.search(regexp, line)]

    if len(indices) == 1:
        # found a single matching line
        index = indices[0]
    elif len(indices) == 0:
        # found no matching lines
        raise Exception('Index not found for: ' + regexp)
    else:
        # found multiple matching lines
        raise Exception('Too many matches for: ' + regexp)

        # ask user which one
        # usr_index = input('What line is ' + regexp + ' on? Possible lines: ' + str(indices))
        # try:
        # index = int(usr_index)
        # except ValueError:
        # if they don't know they'll probably just type anything
        # raise Exception('Invalid index.  Exiting.')

    return index


def find_value(regexp, data_list):
    """
    In the .ISD files, single values are written as X=y
    e.g. Width50 [nm]=0
    e.g. RadiometricUnit=W
    Find a line based on the name (X) and return the value (y)

    params:
    regexp - regex expression or string to search for
    data_list - list where each entry is a line from the '.ISD' file

    returns: value (string)
    """
    val_line = data_list[find_line(regexp, data_list)]

    name, value = re.split('=', val_line)
    value = value[:-1]  # remove \n   (1\n -> 1)
    return value


# ===================================================================================================================
# SCRIPT STARTS HERE

# Iterate through ".ISD" files in data folder
# For each file extract the data, and save in dataframe all_spectra_df
files = gl.glob(data_folder + '\\*.ISD')
all_spectra_df = pd.DataFrame()

for file in files:
    # open file, read contents into a list, and close file:
    data_file = open(file, "r")
    data = data_file.readlines()
    data_file.close()

    # find source current value
    source_I = find_value(r'Currentsource/SourceCurrent', data)

    # find where spectrum data starts
    # the file is structured so that spectra data starts one line below the word 'Data'
    idx = find_line(r'Data\n', data)
    # the file states how many data points there are in the spectrum (value is called NumberOfDataX)
    num_lines = int(find_value(r'NumberOfDataX', data))
    # extract the lines corresponding to the spectrum data
    spectrum_data = data[idx + 1:idx + num_lines + 1]

    # extract the numbers ( x,y -> wavelength (nm), power (W/nm) )
    wl_p = []
    for line in spectrum_data:
        x, y = re.split('\t', line)
        wl_p.append([float(x), float(y[:-1])])

    # store spectra data in a table

    df = pd.DataFrame(columns=['Wavelength (nm)', source_I + ' mA'], data=wl_p)

    # for first iteration, the overall dataframe is empty, so you can't merge

    if len(all_spectra_df) == 0:
        all_spectra_df = df

    # add data from this file to the overall dataframe
    all_spectra_df = pd.merge(all_spectra_df, df)

# Often want to plot only certain spectra
# Extract the spectra specified by "important_current_values"

if (important_current_values is None) or (len(important_current_values) == 0):
    # if no columns are specified then just default to taking first, last, and middle columns
    (r, c) = all_spectra_df.shape
    default_cols = [0, 2, round(c / 2), c - 1]  # col 0 is wavelength
    specified_spectra_df = all_spectra_df.iloc[:, default_cols]

else:
    # extract the specified columns
    # important_current_values are given as floats, but column names are strings
    # e.g. important_current_values = [0.1, 0.5], column names are '0.1 mA', '0.5 mA'
    cols = [str(val) + ' mA' for val in important_current_values]  # convert floats to strings matching the column names
    cols.insert(0, 'Wavelength (nm)')
    specified_spectra_df = all_spectra_df.loc[:, cols]  # extract relevant columns by name

# Write data to an excel file
excel_file = pathlib.Path(save_folder + "\\" + save_filename + ".xlsx")
# if the excel file already exists, overwriting it could cause problems, so delete it
if excel_file.exists():
    print("Excel file exists")
    os.remove(excel_file)
    print("Deleted existing file, made new file")
else:
    print("Excel file does not exist")
    print("Made new file")

with pd.ExcelWriter(excel_file) as xls_wr:
    all_spectra_df.to_excel(xls_wr, "All Spectra")
    specified_spectra_df.to_excel(xls_wr, "Chosen Spectra")

# Creating the spectral plot
wavelength = specified_spectra_df.loc[:, cols[0]]  # extract the wavelength column from specified_spectra_df

# Extract the columns corresponding to the desired current values
# Plot those columns, setting the legend label as the column label
j = 0
for i in cols[1:]:
    power = specified_spectra_df.loc[:, i]
    plt.plot(wavelength, power, label=important_current_values_scientific[j]+' A')
    j = j+1

# Plot parameters (title, axis titles, legend, etc.)
plt.xlabel('Wavelength (nm)')
plt.ylabel('Radiant Flux (W/nm)')
plt.title('Emission Spectra')
plt.legend()
plt.show()
