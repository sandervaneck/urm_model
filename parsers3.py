# -*- coding: utf-8 -*-
"""
Created on Wed Jul 24 10:58:13 2024

@author: M67B363
"""
import pandas as pd
from multiprocessing import Pool, cpu_count

def convert_to_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return None
parameters = [
    "urm state", "savings start", "savings start honorary", "benefit start", "pensionbase",
    "birthdate", "startdate prognosis", "calculationdate", "enddate prognosis",
    "default retirement date", "age fraction", "year fraction retirement year",
    "amount of scenarios", "pp ratio", "start year", "person id"
]

def parse_parameter_chunk(df_chunk):
    # Filter out rows where 'urm state' is not a string
    df_chunk = df_chunk[(df_chunk["urm state"].notna()) & (df_chunk["urm state"].apply(lambda x: isinstance(x, str)))]
    
    # Define date columns
    date_columns = [ "birthdate", 
                    "startdate prognosis", "calculationdate", "enddate prognosis", "default retirement date"]
    
    # Handle non-date columns separately for type conversion
    non_date_columns = {
        "urm state": "string", 
        "savings start": "int",
        "savings start honorary": "int",
        "benefit start": "int",
        "pensionbase": "int", 
        "age fraction": "float",
        "year fraction retirement year": "float", 
        "amount of scenarios": "int",
        "pp ratio": "float", 
        "start year": "int",
        "person id": "string"
    }
    
    # Convert integer date values to actual date values if they are not already datetime
    for col in date_columns:
        if not pd.api.types.is_datetime64_any_dtype(df_chunk[col]):
            df_chunk[col] = pd.to_datetime(df_chunk[col], origin='1899-12-30', unit='D').dt.date

    # Convert other columns to appropriate data types
    df_chunk = df_chunk.astype(non_date_columns)
    
    # Convert the DataFrame to a dictionary of records
    results = df_chunk.to_dict("records")
    
    return results

def parse_fourd_chunk(df_chunk):
    df_chunk = df_chunk[(df_chunk["year"].notna()) & (df_chunk["year"].apply(lambda x: isinstance(x, (int, float))))]
    df_chunk = df_chunk.astype({"year": "int", "scenario": "int", "cohort": "int", "cwf_op": "float", "cwf_pp": "float", "total_return": "float", "total_return_hon": "float"})
    
    results = df_chunk.to_dict("records")
    return results

def read_and_parse_sheet(args):
    filename, sheetname = args
    results = []
    
    # Read the entire sheet into memory
    df = pd.read_excel(filename, sheet_name=sheetname, skiprows=0)
    df.columns = ["year",
    "scenario",
    "cohort",
    "cwf_op",
    "cwf_pp",
    "total_return",
    "total_return_hon"]
    #["A", "B", "C", "D", "E", "F", "G"]
    
    # Process the DataFrame in chunks manually
    chunk_size = 10000
    for start_row in range(0, len(df), chunk_size):
        end_row = min(start_row + chunk_size, len(df))
        df_chunk = df.iloc[start_row:end_row]
        chunk_results = parse_fourd_chunk(df_chunk)
        results.extend(chunk_results)
    
    return results

def parse_fourds(self):
    sheetnames = ["18-27", "28-37", "38-47", "48-57", "58-67", "68-77", "78-87", "88-97", "98-107", "108-117", "118-127"]
    filename = self.filename
    with Pool(cpu_count()) as pool:
        all_results = pool.map(read_and_parse_sheet, [(filename, sheetname) for sheetname in sheetnames])
    
    # Flatten the list of lists into a single list
    flat_results = [item for sublist in all_results for item in sublist]
    
    return flat_results

def vertical_parameters(self):
    results = []
    df = pd.read_excel(self.filename, sheet_name=self.parameters_sheet_naam, skiprows=0)
    df.columns = self.parameters_variables
    chunk_size = 10000
    for start_row in range(0, len(df), chunk_size):
        end_row = min(start_row + chunk_size, len(df))
        df_chunk = df.iloc[start_row:end_row]
        chunk_results = parse_parameter_chunk(df_chunk)
        results.extend(chunk_results)
    
    return results


def parse_parameters(self):
    df = self.parameters_sheet
    parameters_variables = self.parameters_variables
    start_row = 0  # Excel row 2 (0-based index)
    result_list = []
    col_index = 1  # Start from column B (0-based index)

    while col_index < df.shape[1] and not df.iloc[:, col_index].isnull().all():
        result = {}
        for idx, param in enumerate(parameters_variables):
            # Mapping of parameters to specific rows (0-based index)
            
            row_mapping = {
                'urm state': 1,
                'savings start': 2,
                'savings start honorary': 3,
                'benefit start': 4,
                'pensionbase': 5,
                'birthdate': 6,
                'startdate prognosis': 7,
                'calculationdate': 8,
                'enddate prognosis': 9,
                'default retirement date': 10,
                'age fraction': 11,
                'year fraction retirement year': 12,
                'amount of scenarios': 13,
                'pp ratio': 14,
                'start year': 15,
                'person id': 16
            }
            
            row_index = start_row + row_mapping[param] - 1
            
            if row_index < len(df):
                cell_value = df.iloc[row_index, col_index]

                if param in ['birthdate', 'calculation date', 'default retirement date', 'end date prognosis', "start date prognosis"]:
                    if isinstance(cell_value, int):
                        date_value = pd.to_datetime('1899-12-30') + pd.to_timedelta(cell_value, 'D')
                        result[param] = date_value.date()
                    else:
                        result[param] = cell_value
                else:
                    result[param] = cell_value

        result_list.append(result)
        col_index += 1

    return result_list

