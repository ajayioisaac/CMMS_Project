# -*- coding: utf-8 -*-
"""
Created on Mon Nov 13 11:30:57 2023

@author: Planner
"""
import pandas as pd
import numpy as np
from datetime import datetime

def Maint_processor():
    # Define the expected format
    date_format = "%d-%m-%Y"  # For example, "2023-11-15 14:03:09"
    
    # Ask the user for input
    user_lower_limit = input("Enter a lower limit date (DD-MM-YYYY): ")
    user_upper_limit = input("Enter a upper limit date (DD-MM-YYYY): ")
    
    try:
        # Parse the user input using the specified format
        lower_limit = datetime.strptime(user_lower_limit, date_format)
    except ValueError:
        print("Invalid input. Please use the format YYYY-MM-DD HH:MM:SS.")
        
    try:
        # Parse the user input using the specified format
        upper_limit = datetime.strptime(user_upper_limit, date_format)
    except ValueError:
        print("Invalid input. Please use the format YYYY-MM-DD HH:MM:SS.")
    
    save_file = input("File save name: ")
    save_as = "P:/MAINTENANCE/Maint_Planner/Schedule/" + save_file + '.xlsx'
    exportable = [['WO NUMBER', 'PM NO', 'OBJECT ID', 'OBJECT DESCRIPTION', 'WORK DESCRIPTION', 'JOB PROCEDURE', 'OP_STATUS', 'PRIORITY','ACTION', 'WORK TYPE', 'DEPT RESPONSIBLE', 'RESOURCE', 'PLANNED QTY', 'DURATION', 'PLANNED START']]
    
    #This is the file that contain the maintenance plan
    maint_file = "C:/Users/Planner/Documents/Planning/Working Document.xlsx"
    
    #Select the sheet name
    sheet_name_r = "Active PM"
    
    # Read workbook
    read_maint_file = pd.read_excel(maint_file, sheet_name=sheet_name_r)
    
    # Convert to array
    maint_file_array = np.array(read_maint_file)
    
    # Target action columnn
    due_date = maint_file_array[:, 10]
    
    for index, value in enumerate(due_date):
        value_datetime = pd.to_datetime(value)  # Convert numpy datetime64 to Python datetime
        if lower_limit <= value_datetime <= upper_limit:
            exportable.append(('',maint_file_array[index][0], maint_file_array[index][1], maint_file_array[index][2], maint_file_array[index][3], maint_file_array[index][4], maint_file_array[index][13], maint_file_array[index][14], maint_file_array[index][5], maint_file_array[index][11], maint_file_array[index][12], maint_file_array[index][15], maint_file_array[index][16], maint_file_array[index][17], value_datetime ))
        else:
            continue
       
     # Output
    output = pd.DataFrame(exportable)
     
    excel_file_path = save_as
     
    output.to_excel(excel_file_path, index=False)
    #data Cleaning
    import openpyxl
     
    # Load the Excel file
    workbook = openpyxl.load_workbook(save_as)
     
    # Select the worksheet you want to work with (e.g., the first sheet)
    sheet = workbook.active
     
    # Specify the row you want to edit (e.g., row 2)
    row_number = 1
     
    # Delete the top row
    sheet.delete_rows(row_number)
     
    # Save the changes back to the Excel file
    workbook.save(save_as)
     
    # Close the workbook
    workbook.close()

def Overdue_processor():
    current_datetime = datetime.now()
    
    save_as = "P:/MAINTENANCE/Maint_Planner/Schedule/Overdue_PM_Jobs.xlsx"
    
    exportable = exportable = [['WO NUMBER', 'PM NO', 'OBJECT ID', 'OBJECT DESCRIPTION', 'WORK DESCRIPTION', 'JOB PROCEDURE', 'OP_STATUS', 'PRIORITY','ACTION', 'WORK TYPE', 'DEPT RESPONSIBLE', 'RESOURCE', 'PLANNED QTY', 'DURATION', 'PLANNED START']]
    
    #This is the file that contain the maintenance plan
    maint_file = "C:/Users/Planner/Documents/Planning/Working Document.xlsx"
    
    #Select the sheet name
    sheet_name_r = "Active PM"
    
    # Read workbook
    read_maint_file = pd.read_excel(maint_file, sheet_name=sheet_name_r)
    
    # Convert to array
    maint_file_array = np.array(read_maint_file)
    
    # Target action columnn
    due_date = maint_file_array[:, 10]
    
    for index, value in enumerate(due_date):
        value_datetime = pd.to_datetime(value)  # Convert numpy datetime64 to Python datetime
        if value_datetime < current_datetime  :
            exportable.append((maint_file_array[index][0], maint_file_array[index][1], maint_file_array[index][2], maint_file_array[index][3], maint_file_array[index][4], maint_file_array[index][13], maint_file_array[index][14], maint_file_array[index][5], maint_file_array[index][11], maint_file_array[index][12], maint_file_array[index][15], maint_file_array[index][16], maint_file_array[index][17], value_datetime))
        else:
            continue
    
     # Output
    output = pd.DataFrame(exportable)
     
    excel_file_path = save_as
     
    output.to_excel(excel_file_path, index=False)
     
    #data Cleaning
    import openpyxl
     
    # Load the Excel file
    workbook = openpyxl.load_workbook(save_as)
     
    # Select the worksheet you want to work with (e.g., the first sheet)
    sheet = workbook.active
     
    # Specify the row you want to edit (e.g., row 2)
    row_number = 1
     
    # Delete the top row
    sheet.delete_rows(row_number)
     
    # Save the changes back to the Excel file
    workbook.save(save_as)
     
    # Close the workbook
    workbook.close()