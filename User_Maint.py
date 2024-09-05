# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 11:04:56 2023

@author: Planner
"""

from Upgraded_Increment_Swap_function import *
from Maintenance_Processor import *


def User_action():
    while True:
        users_input = input('Enter "M" for Maintenance, "O" for Overdue, or "Q" to quit: ').upper()
        if users_input == 'M':
            
            # rename the old week file name
            clean_swap()
            
            # Run the maintenance processor 
            Maint_processor()
            break
        
        elif users_input == 'O':
            # run the date increment, produces a filled called Plan_Incremental_DND
            clean_swap()
            
            # Run the overdue processor
            Overdue_processor()
            break
        
        elif users_input == 'Q':
            print("Exiting the program.")
            break      
    
User_action()