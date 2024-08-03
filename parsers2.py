# -*- coding: utf-8 -*-
"""
Created on Tue Jul 23 14:28:56 2024

@author: M67B363
"""

from multiprocessing import Pool, cpu_count
import pandas as pd
from parsers import convert_to_float

def parse_fourd(df):
    results = []
    for index, row in df.iterrows():
        year = row['A']
        if pd.isna(year) or not isinstance(year, (int, float)):
            break
        
        result = {
            "year": year,
            "scenario": row['B'],
            "cohort": row['C'],
            "cwf_op": convert_to_float(row['D']),
            "cwf_pp": convert_to_float(row['E']),
            "total_return": convert_to_float(row['F']),
            "total_return_hon": convert_to_float(row['G'])
        }
        results.append(result)
    print("Tabblad ingelezen")
    return results

def read_and_parse_sheet(args):
    filename, sheetname = args
    df = pd.read_excel(filename, sheet_name=sheetname, skiprows=1)
    df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    return parse_fourd(df)

def parse_fourds(self):
    sheetnames = ['18-27', '28-37', '38-47', '48-57', '58-67', '68-77', '78-87', '88-97', '98-107', '108-117', '118-127']
    filename = self.filename
    with Pool(cpu_count()) as pool:
        results = pool.map(read_and_parse_sheet, [(filename, sheetname) for sheetname in sheetnames])
    print("4ds geschreven")
    print(results[0])
    return results


def parse_twod(df):
    results = []
    
    # Iterate through the DataFrame rows
    for index, row in df.iterrows():
        year = row['A']  # Assuming 'A' is the column name for the 'year'
        
        # Check if year is valid
        if pd.isna(year) or year == "" or not isinstance(year, (int, float)):
            break
        
        result = {
            "year": year,
            "scenario": row['B'],  # Assuming 'B' is the column name for 'scenario'
            "one_year_inflation": convert_to_float(row['C']),  # Assuming 'C' is the column name for 'one_year_inflation'
            "cpi": convert_to_float(row['D']),  # Assuming 'D' is the column name for 'cpi'
            "ff": convert_to_float(row['E']),  # Assuming 'E' is the column name for 'ff'
            "payout_adjustment": convert_to_float(row['F']),  # Assuming 'F' is the column name for 'payout_adjustment'
            "sr_adjustment": convert_to_float(row['G']),  # Assuming 'G' is the column name for 'sr_adjustment'
            "contribution_rate": convert_to_float(row['H'])  # Assuming 'H' is the column name for 'contribution_rate'
        }
        results.append(result)
    
    return results

def read_and_parse_twod(self):
    sheetname = self.twod_sheet_naam
    filename = self.filename
    df = pd.read_excel(filename, sheet_name=sheetname, skiprows=1, header=None)  # Adjust based on your actual sheet
    df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']  # Set column names
    return parse_twod(df)