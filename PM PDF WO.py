# -*- coding: utf-8 -*-
"""
Created on Tue Dec  5 10:09:02 2023

@author: Planner
"""


import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime
import re

def convert_rows_to_pdf(input_excel, output_folder):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_excel)

    # Set A4 page size with 1-inch margins
    page_width, page_height = letter
    margin = 72  # 1 inch in points (1 inch = 72 points)
    usable_width = page_width - 2 * margin

    # Iterate through rows and generate PDFs
    for index, row in df.iterrows():
        # Get the value from the first cell of the row
        file_name = str(row.iloc[0])
        output_folder1 = str(row.iloc[10])
        #output_folder1 = "MECH"
        

        # Create a PDF file with A4 size and 1-inch margins
        pdf_path = f"{output_folder}/{output_folder1}/{file_name}.pdf"
        c = canvas.Canvas(pdf_path, pagesize=letter)
        c.translate(margin, margin)  # Adjust the origin to account for margins

        # Set font and initial font size
        font_name = "Helvetica"
        font_size = 10
        c.setFont(font_name, font_size)

        # Specify x-coordinates for left, center, and right cells within usable width
        left_x = 0
        center_x = usable_width / 2
        right_x = usable_width

        # Specify initial line height within usable height
        line_height = page_height - 2 * margin

        # Add content to the PDF with adjusted positioning and dynamic font size
        # Left side: "Tamara Elmina"
        c.drawString(left_x, line_height, "CESL Tamara Elmina")
        
        # Center: Column 1
        work_order_number = row.get('WORK ORDER NO', '')  # Update column name
        c.drawCentredString(center_x, line_height, f"WORK ORDER NO: {work_order_number}")

        # Right side: Current date
        current_date = datetime.now().strftime("%Y-%m-%d")
        c.drawRightString(right_x, line_height, f"{current_date}")
        line_height -= 15  # Adjust for the next line
        
        # For the second row
        # I need to add this column in subsequent PM generation and change the column name to ACTION DESCR 
        action = row.get('ACTION', '')
        c.drawCentredString(center_x, line_height, f"Directive: {action}")
        line_height -= 15  # Adjust for the next line
        
        # For the third row
        work_description = row.get('WORK DESCRIPTION', '')
        c.drawCentredString(center_x, line_height, f"Work Description: {work_description}")
        line_height -= 15  # Adjust for the next line
        
        line_height -= 15 # Space
        
        # For the fourth row
        c.drawString(left_x, line_height, "FPSO TAMARA ELMINA")
        line_height -= 10  # Adjust for the next line
        
        # Draw a horizontal line
        start_x, start_y = 0, page_height - 2 * margin - 5 * 13  # Assuming each row is 15 points high
        end_x, end_y = page_width - 2 * margin, page_height - 2 * margin - 5 * 13
        c.line(start_x, start_y, end_x, end_y)
        
        line_height -= 15  # Adjust for the next line
        
        # For the fifth Line
        status = row.get('STATUS', '')
        c.drawString(left_x, line_height, f"WO Status:    {status}" )
        
        # Center: Column 1
        object_id = row.get('OBJECT ID', '')
        c.drawCentredString(center_x, line_height, f"Object ID:   {object_id}")
        line_height -= 10  # Adjust for the next line
        
        object_desc = row.get('OBJECT DESSCRIPTION', '')
        c.drawRightString(right_x, line_height, f"{object_desc}")
        line_height -= 15  # Adjust for the next line
        
        # For the sixth Line
        dept_responsible = row.get('DEPT RESPONSIBLE', '')
        c.drawString(left_x, line_height, f"Maint. Org:    {dept_responsible}" )
        
        # Center: Column 1
        ops_status = row.get('OPERATIONAL STATUS', '')
        c.drawRightString(right_x, line_height, f"Operation Status:    {ops_status}")
        line_height -= 15  # Adjust for the next line
        
        # For the seventh Line
        planned_start = row.get('PLANNED START', '')
        c.drawString(left_x, line_height, f"Planned Start:    {planned_start}" )
        
        # Center: Column 1
        priority = row.get('PRIORITY')
        c.drawCentredString(300, line_height, f"Priority:    {priority}")
        line_height -= 15  # Adjust for the next line
        
        # For the eight Line
        planned_finish = row.get('PLANNED FINISH', '')
        c.drawString(left_x, line_height, f"Planned Finish:    {planned_finish}" )
        
        pm_no = row.get('PM NO', '')
        c.drawCentredString(center_x, line_height, f"PM NO:    {pm_no}")
        line_height -= 15  # Adjust for the next line
        
        # Set font and initial font size
        font_name = "Helvetica"
        font_size = 14
        c.setFont(font_name, font_size)
        
        start_x, start_y = 0, page_height - 2 * margin - 10 * 14.7  # Assuming each row is 15 points high
        end_x, end_y = page_width - 2 * margin, page_height - 2 * margin - 10 * 14.7
        c.line(start_x, start_y, end_x, end_y)
        line_height -= 15
        
        # For the ninth Line
        c.drawCentredString(center_x, line_height, "Work Tasks")
        line_height -= 15  # Adjust for the next line
        
        start_x, start_y = 0, page_height - 2 * margin - 12 * 15  # Assuming each row is 15 points high
        end_x, end_y = page_width - 2 * margin, page_height - 2 * margin - 12 * 15
        c.line(start_x, start_y, end_x, end_y)
        line_height -= 15
        
        # Set font and initial font size
        font_name = "Helvetica"
        font_size = 10
        c.setFont(font_name, font_size)
        
        # For the tenth Line
        job_procedure = row.get('JOB PROCEDURE', '')
        cell_content = str(job_procedure)

        # Split the content based on the presence of numbers
        sentences = re.split(r'\s(\d+\.)', cell_content)

        # Write each sentence on a new line
        for sentence in sentences:
            # Check if the sentence is not None and not an empty string
            if sentence is not None and sentence.strip():
                if len(sentence) > 100:
                    c.drawString(start_x, line_height, sentence[:100]) 
                    line_height -= 12
                    if(len(sentence[100:])) > 100:  
                        c.drawString(start_x, line_height, sentence[100:201])
                        line_height -= 12
                        if(len(sentence[201:])) > 100:  
                            c.drawString(start_x, line_height, sentence[201:301])
                            line_height -= 12
                        else:                            
                            c.drawString(start_x, line_height, sentence[201:])
                            line_height -= 12
                    else:
                        c.drawString(start_x, line_height, sentence[100:])
                    line_height -= 12
                else: 
                    c.drawString(start_x, line_height, sentence) # Adjust x-coordinate as needed
                    line_height -= 15  # Adjust for the next line

        start_x, start_y = 0, page_height - 2 * margin - 45 * 14.7  # Assuming each row is 15 points high
        end_x, end_y = page_width - 2 * margin, page_height - 2 * margin - 45 * 14.7
        c.line(start_x, start_y, end_x, end_y)
        line_height -= 15  

        
        # Save the PDF file
        c.save()


if __name__ == "__main__":
    # Specify the input Excel file and output folder
    #input_excel_file = "C:/Users/Planner/Desktop/file 2.xlsx"
    input_excel_file = "P:/MAINTENANCE/Maint_Planner/Schedule/Maint_Jobs_15th - 28th  Jan. 2024.xlsx"
    output_folder = "P:/MAINTENANCE/Maint_Planner/Schedule/Work Order/"
    
    #Select the sheet name
    sheet_name_r = "PM Jobs"
    
    # Run the conversion function
    convert_rows_to_pdf(input_excel_file, output_folder)
