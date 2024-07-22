import pandas as pd
import numpy as np
import regex as re
import re
import os
import sys

# Get the command-line arguments
print(str(sys.argv))
df1_folder = sys.argv[1]
df2_path = sys.argv[2]
df3_path = sys.argv[3]
download_path = sys.argv[4]

# Prompt user to input file paths
# df1_folder = input("Enter the folder path for df1: ")
# df2_path = input("Enter the file path for df2: ")
# df3_path = input("Enter the file path for df3: ")
# download_path = input("Enter the download path for: ")

# Read the Excel files

df2 = pd.read_excel(df2_path, sheet_name='Database (NIC)')
df3 = pd.read_excel(df3_path, sheet_name='Battery Database')

# Function to extract site name from a file name
def extract_site_name(file_name):
    file_name = file_name[0:8]
    if '_' in file_name:
        return file_name.split('_')[0]
    else:
        return file_name.split('-')[0]


# Example folder path where your files are located
#folder_path = r"C:\Users\user\Downloads\testing"

# List files in the folder
file_names = os.listdir(df1_folder)

# Match and perform actions
for file_name in file_names:
    site_name = extract_site_name(file_name)
    if site_name in df2['Site ID'].values:
        # Read the file content
        file_path = os.path.join(df1_folder, file_name)

        df1 = pd.read_excel(file_path, sheet_name='PWR')
        print(f"Matched File: {file_name}, Site Name: {site_name}")

        # read necessary variables for filling up service publish
        exbbrand = df1['C&O'][4]
        exbqty = df1['C&O'][5]
        exbv = str(df1['C&O'][6]) + "V"
        exbc = df1['C&O'][7]
        probqty = df1['C&O'][44]
        load = df1['Unnamed: 19'][9]
        power100 = df1['Unnamed: 19'][60]
        exrmodel = df1['Unnamed: 19'][5]
        nor = df1['Unnamed: 19'][6]
        sore = str(df1['C&O'][43])
        if df1['Unnamed: 27'][76] != 0.0:
            probm = str.upper(df1['Unnamed: 27'][76])
        else:
            probm = 0
        if df1['Unnamed: 14'][76] != 0.0:
            pronor = str.upper(df1['Unnamed: 14'][76])
            cab = str.upper(df1['Unnamed: 14'][76])
        else:
            pronor = 0
            cab = 0

        # back up hours required
        df2['Back up Hours Required'] = np.where(df2['Site ID'].isin(df3['Site id']),
                                                 df2['Site ID'].map(df3.set_index('Site id')['Finalize back up hours']),
                                                 df2['Back up Hours Required'])
        print("Backup Hours Required Done")

        # CSU Load Before Swap (A)
        if load == "":
            load = "EMPTY"
        result = str(load)
        df2['Existing CSU Load Before Swap (A)'] = np.where(df2['Site ID'] == site_name, result,
                                                            df2['Existing CSU Load Before Swap (A)'])
        print("CSU Load Before Swap Done")

        # Power Consumption
        if power100 == "":
            power100 = "EMPTY"
        result = str(power100)
        df2['Power Consumption (W) 100%'] = np.where(df2['Site ID'] == site_name, result,
                                                     df2['Power Consumption (W) 100%'])
        print("Power Consumption Done")

        # Rectifier Module
        if exrmodel == "":
            exrmodel = "EMPTY"

        result = str(nor) + " * " + str(exrmodel)
        df2['Existing Rectifier Module - Survey Info'] = np.where(df2['Site ID'] == site_name, result,
                                                                  df2['Existing Rectifier Module - Survey Info'])
        print("Rectifier Module Done")

        # existing battery
        result = str(exbbrand) + "," + str(exbv) + "," + str(exbqty) + "*" + str(exbc)
        df2['Existing battery Onsite- Survey Info'] = np.where(df2['Site ID'] == site_name, result,
                                                               df2['Existing battery Onsite- Survey Info'])
        print("Existing Battery Done")

        # battery
        if probqty == "":
            probqty = "EMPTY"

        result = str(probqty)
        df2['Battery'] = np.where(df2['Site ID'] == site_name, result, df2['Battery'])
        print("Battery Done")

        # Swap or Expansion
        if "SWAP" in sore:
            sore = "SWAP"
        elif "ADD" in sore:
            sore = "EXPANSION"
        else:
            sore = 0
        result = sore
        df2['Swap/Exp (Battery)'] = np.where(df2['Site ID'] == site_name, result, df2['Swap/Exp (Battery)'])
        print("Swap Expansion Done")

        # Swap Vision Add Solution
        pattern1 = r'SWAP\s*(.*)\s*[A-Z]*'
        pattern2 = r'ADD\s*(.*)\s*[A-Z]*'

        if probm != 0 and "RETAIN" in probm:
            probm = 0
        elif probm != 0 and "*" in probm:
            start_word = "NEW"
            space = 4
            if "SWAP" and "VISION" in probm:
                type1 = "VISION"
            elif "SWAP" and "SOLUTION" in probm:
                type1 = "SOLUTION"
            elif "ADD" and "VISION" in probm:
                type1 = "VISION"
            elif "ADD" and "SOLUTION" in probm:
                type1 = "SOLUTION"

            start_index = probm.find(start_word) + space  # Find the index of "NEW" and add 4 to skip "NEW "
            end_index = probm.find(type1)  # Find the index of "VISION"
            battery_spec = probm[start_index:end_index]

            type2 = "Lithium"
            type3 = battery_spec

            full_specs = (type1 + " , " + type2 + " , " + battery_spec)
        else:
            full_specs = 0
        df2['Battery Bank Model & Type'] = np.where(df2['Site ID'] == site_name, full_specs,
                                                    df2['Battery Bank Model & Type'])
        print("Battery Bank Model Done")

        # Rectifier Module and Model
        pattern = r'ADD\s*(.*)\s*[A-Z]*'
        if pronor != 0 and "RETAIN" in pronor:
            pronor = 0
        elif pronor != 0 and "*" in pronor:
            pronor_num = pronor[15]
            match = re.search(pattern, pronor)
            prormodel = match.group(1)
        else:
            prormodel = 0
            pronor_num = 0
        df2['Rectifier Model'] = np.where(df2['Site ID'] == site_name, prormodel, df2['Rectifier Model'])
        df2['Rectifier Module'] = np.where(df2['Site ID'] == site_name, pronor_num, df2['Rectifier Module'])
        print("Rectifier Model and Module Done")

        # Cab
        if cab != 0 and "SWAP" in cab:
            if pronor != 0 and "CABINET" in pronor:
                cab = "1"
        else:
            cab = "0"
        df2['Cab'] = np.where(df2['Site ID'] == site_name, cab, df2['Cab'])
        print("Cab Done")

df2.to_excel(os.path.join(download_path, 'updated_report.xlsx'))



