import pandas as pd
from datetime import datetime
import openpyxl
import os


def Maint_processor():
    date_format = "%d-%m-%Y"

    user_lower_limit = input("Enter a lower limit date (DD-MM-YYYY): ")
    user_upper_limit = input("Enter a upper limit date (DD-MM-YYYY): ")

    try:
        lower_limit = datetime.strptime(user_lower_limit, date_format)
        upper_limit = datetime.strptime(user_upper_limit, date_format)
    except ValueError:
        print("Invalid input. Please use the format DD-MM-YYYY.")
        return

    save_file = input("File save name: ")

    output_dir = r"D:/CESL/CMMS Issues/Refined/Maintenance_Schedule"
    os.makedirs(output_dir, exist_ok=True)
    save_as = os.path.join(output_dir, save_file + ".xlsx")

    headers = [
        'PM NO', 'WO NUMBER', 'OBJECT ID', 'OBJECT DESCRIPTION',
        'WORK DESCRIPTION', 'WORK TYPE', 'DEPT RESPONSIBLE', 'FREQUENCY', 'PERIOD',
        'PLANNED DUE DATE', 'PROCEDURE'
    ]

    maint_file = r"D:/CESL/CMMS Issues/Refined/Preventive_Mainenance_data_2026-01-31T08_29_13.606Z.xlsx"
    sheet_name_r = "Preventive_Mainenance_data_2026"

    df = pd.read_excel(maint_file, sheet_name=sheet_name_r)

    # Ensure date column is datetime
    df["nextDueDate"] = pd.to_datetime(
    df["nextDueDate"],
    utc=True,
    errors="coerce"
).dt.tz_localize(None)

    df = df.dropna(subset=["nextDueDate"])
    exportable = []

    today_str = datetime.today().strftime("%Y-%m-%d")
    wo_counter = 1

    for _, row in df.iterrows():
        if lower_limit <= row["nextDueDate"] <= upper_limit:
            wo_number = f"WO-{today_str}-{wo_counter}"
            wo_counter += 1

            exportable.append([
                row["PM NO"],
                wo_number,
                row["machineId.equipmentNo"],
                row["machineId.name"],
                row["title"],
                row["criticality"],
                row["department"],
                row["frequency"],                     
                row["period"],                
                row["nextDueDate"],
                row["procedures.0.title"]
            ])

    output_df = pd.DataFrame(exportable, columns=headers)
    output_df.to_excel(save_as, index=False)

    print(f"File successfully generated: {save_as}")

    return save_as


def Overdue_processor():
    current_datetime = datetime.now()

    save_as = r"D:/CESL/CMMS Issues/Refined/Maintenance_Schedule/Overdue_PM_Jobs.xlsx"

    headers = [
        'PM NO', 'OBJECT ID', 'OBJECT DESCRIPTION',
        'WORK DESCRIPTION', 'WORK TYPE', 'DEPT RESPONSIBLE', 'FREQUENCY', 'PERIOD',
        'PLANNED DUE DATE', 'PROCEDURE'
    ]

    maint_file = r"D:/CESL/CMMS Issues/Refined/Preventive_Mainenance_data_2026-01-31T08_29_13.606Z.xlsx"
    sheet_name_r = "Preventive_Mainenance_data_2026"

    df = pd.read_excel(maint_file, sheet_name=sheet_name_r)
    df["nextDueDate"] = pd.to_datetime(
    df["nextDueDate"],
    utc=True,
    errors="coerce"
).dt.tz_localize(None)

    df = df.dropna(subset=["nextDueDate"])

    overdue_df = df[df["nextDueDate"] < current_datetime]

    export_df = overdue_df[
        [
            "PM NO",
            "machineId.equipmentNo",
            "machineId.name",
            "title",
            "criticality",
            "department",
            "frequency",
            "period",
            "nextDueDate",
            "procedures.0.title"
        ]
    ]

    export_df.columns = headers
    export_df.to_excel(save_as, index=False)

    print("Overdue PM Jobs file generated successfully.")
