from Increment_New import *
from Maint_Processor import *
from Work_Pack_Generator import generate_work_packs


def User_action():
    while True:
        users_input = input(
            'Enter "M" for Maintenance, "O" for Overdue, or "Q" to quit: '
        ).upper()

        if users_input == 'M':

            # 1️⃣ Update dates
            clean_swap()

            # 2️⃣ Generate maintenance schedule
            maintenance_file = Maint_processor()

            # 3️⃣ Generate work packs automatically
            if maintenance_file:
                generate_work_packs(maintenance_file)

            break

        elif users_input == 'O':
            clean_swap()
            Overdue_processor()
            break

        elif users_input == 'Q':
            print("Exiting the program.")
            break


User_action()
