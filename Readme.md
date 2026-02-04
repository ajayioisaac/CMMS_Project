This is a Project to design an end to end maintenance scheduling process, mimicing how a CMMS works.

Problem: 
We had challenge renewing the license of our current CMMS and I as the maintenance planner had to be creative to ensure the maintenance scheduling process is not hampered too much. 

Approach:

So what I envisioned a custom, lightweight CMMS replacement to keep our preventive maintenance operations running despite the expired CMMS license. The Steps include:

1. Weekly Maintenance Sheet (Input)

	An excel file shared on the common drive, updated by technicians with completion data.

	Acts as the real-time input source for the system.

2. Master Data Update Script

	Reads the weekly sheet.

	Updates the preventive maintenance master data with the last completion date for each activity.

3. Increment_New.py

	Swaps “last done date” and calculates the next due date based on frequency and period.

	Keeps the PM schedule rolling accurately.

4. Maintenance Processor

	Reads updated master data.

	Extracts activities due within a specified window, essentially identifying what should be in the next weekly routine.

5. Integration via user.py

	Combines all scripts into a single module.

	Associated with a desktop shortcut for one-click execution, making it easy for me or anyone in my team to run the whole workflow without 	touching individual scripts.

In short, I’ve recreated a fully functional PM workflow with automation for: tracking, scheduling, and planning. It’s a smart, low-cost, and practical solution to bridge the gap until the CMMS license can be renewed.
