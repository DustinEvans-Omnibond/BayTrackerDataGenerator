#-------------------------------------------------------------------------------
# Name:        excel_generator.py
# Purpose:
#
# Author:      Dustin Evans
#
# Created:     26/04/2016
# Copyright:   (c) Dustin Evans 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter
from openpyxl.drawing.image import Image
from datetime import datetime, timedelta

def write_data(data_dict, dest_filename):
    wb = Workbook()

    sheet = wb.active
    sheet.title = 'BayTracker Data'

    # Create column headers
    headers = ['Bay', 'Start Time', 'Service Time', 'Wait Time', 'Total Customer Time', 'Snapshot']
    for col in range(1, len(headers)+1):
        _ = sheet.cell(column=col, row=1, value=headers[col-1])


    # Write data to sheet
    vehicles = data_dict['vehicles']
    current_row = 2
    for obj in vehicles:
        bay = obj['bay']
        start_time = datetime.fromtimestamp(obj['t_enter']).strftime('%I:%M %p')
        service_time = abs(obj['t_leave'] - obj['t_enter'])
        wait_time = abs(obj['t_enter'] - obj['t_queue_enter'])
        total_customer_time = service_time + wait_time
        snapshot = obj['snapshot']
        col_entries = [bay, start_time, str(timedelta(seconds=service_time)), str(timedelta(seconds=wait_time)), str(timedelta(seconds=total_customer_time)), snapshot]
        for col in xrange(1, len(col_entries)+1):
            if col == len(col_entries):
                # Snapshot
                img = Image(col_entries[col-1])
                sheet.add_image(img, get_column_letter(col))
            else:
                _ = sheet.cell(column=col, row=current_row, value=col_entries[col-1])

        current_row += 1


    # Write Excell workbook to destination file
    wb.save(filename=dest_filename)