# TOOL 1 - GeneratePi

# Version 1: Pi functionality added



import openpyxl
from openpyxl import Workbook
from datetime import datetime
from datetime import timedelta
import os
import pandas as pd
import xlwings as xw
import mikeio
from PI_model_automation_variables import * #import everything

if generate_pi:
    timestep_dict = {}
    timestep_dict["5m"]=300
    
    pi_end_time = datetime.today().replace(minute=0, second=0, microsecond=0) # today's time round down to nearest hour
    
    interval = "5m"
    
    pi_dict = {}
    pi_dict['HGL'] = '"average","time-weighted"'
    pi_dict['Rainfall'] = '"total","event-weighted"'
    pi_dict['Flow'] = '"total","event-weighted"'
    
    df = pd.read_excel(inputsheet_path)
    
    for index, row in df.iterrows():
        
        # Create workbook and worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
    
        name = row["Name"]
        tag = row["Tag"]
        item_type = row["Type"]
        unit = row["Unit"]
        desc = f"{type} ({unit}) in gauge {name}"
    
        dfs0_path = row["DFS0 Path"]
        dfs = mikeio.read(dfs0_path)
        dfs0_end_time = dfs.end_time
        pi_start_time = dfs0_end_time + timedelta(minutes=5)
        
        ws["A1"] = "Start"
        ws["B1"] = pi_start_time.strftime("%Y-%m-%d %H:%M:%S")
        ws["A2"] = "End"
        ws["B2"] = pi_end_time.strftime("%Y-%m-%d %H:%M:%S")
    
        # Duration
        duration =  pi_end_time -  pi_start_time
        duration_seconds = duration.total_seconds()
        timestep_seconds = timestep_dict[interval]
        timestep_number = duration_seconds/timestep_seconds - 1
        last_row = int(timestep_number + 10)
    
        # Interval
        ws["A3"] = "Interval"
        #interval = input("Enter Interval - e.g.5m : ")
        ws["B3"] = interval
    
        # Name (used for Excel file name)
        ws["A5"] = "Name"
        # name = input("Enter Name: ")
        ws["B5"] = name
    
        # Tag
        ws["A6"] = "Tag"
        ws["B6"] = tag
    
        # Desc
        ws["A7"] = "Desc"
        ws["B7"] = desc
    
        # Unit
        ws["A8"] = "Unit"
        ws["B8"] = unit
    
        # Array formula
        ws["A10"] = r'=PIAdvCalcDat(Sheet1!$B$6,Sheet1!$B$1,Sheet1!$B$2,Sheet1!$B$3,' + pi_dict[item_type] + ',0,1,65,"\\gvprdhist01")'
        ws.formula_attributes["A10"] = {'t': 'array', 'ref': f"A10:B{last_row}"} 
        for i in range(10,last_row + 1):
            ws[f'A{i}'].number_format = 'yyyy-mm-dd hh:mm:ss'
        ws.column_dimensions['A'].width = 25
        # Freeze the first column
        ws.freeze_panes = ws['A10']
    
        # Save the file
        file_path = f"{output_folder}\\{name}-{item_type}.xlsx"
        wb.save(file_path)
                    
        # Open, save, and close the workbook
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        wb.save()
        wb.close()
        app.quit()
    
    #     os.startfile(file_name)  
    #     Ctrl + / the above line for debugging only
    
        print(f"Excel file '{file_path}' has been created!!")