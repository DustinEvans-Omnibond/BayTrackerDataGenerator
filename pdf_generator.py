#-------------------------------------------------------------------------------
# Name:        pdf_generator.py
# Purpose:
#
# Author:      Dustin Evans
#
# Created:     28/04/2016
# Copyright:   (c) Dustin Evans 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from pytz import timezone
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Image
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from spreadsheettable import SpreadsheetTable

styleSheet = getSampleStyleSheet()
MARGIN_SIZE = 5 * mm
PAGE_SIZE = A4

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def write_data(data_dict, dest_filename, default_timezone):
    table_style = [
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1),'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('FONTNAME', (0,0), (-1,0), 'Times-Bold'),
    ]

    # Get meta info headers
    if 'headers' in data_dict:
        meta_headers = data_dict['headers']

    data = [
        ['Bay', 'Start Time', 'Service Time', 'Wait Time', 'Total Customer Time', 'Snapshot']
    ]

    vehicles = data_dict['vehicles']
    for obj in vehicles:
        bay = obj['bay']
        tz = timezone(default_timezone)
        if 'timezone' in obj:
            tz = timezone(obj['timezone'])
        start_time = datetime.fromtimestamp(obj['t_enter'], tz).strftime('%Y/%m/%d %I:%M %p')
        service_time = abs(obj['t_leave'] - obj['t_enter'])
        wait_time = abs(obj['t_enter'] - obj['t_queue_enter'])
        total_customer_time = service_time + wait_time
        snapshot = Image(obj['snapshot'], 40*mm, 30*mm)

        col_entries = [bay, start_time, str(timedelta(seconds=service_time)), str(timedelta(seconds=wait_time)), str(timedelta(seconds=total_customer_time)), snapshot]
        data.append(col_entries)

    story = []
    spreadsheet_table = SpreadsheetTable(data, repeatRows = 1)
    spreadsheet_table.setStyle(table_style)
    story.append(spreadsheet_table)

    create_pdfdoc(dest_filename, story)



#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def create_pdfdoc(dest_filename, story):
    pdf_doc = BaseDocTemplate(dest_filename, pagesize=PAGE_SIZE,
        leftMargin=MARGIN_SIZE, rightMargin=MARGIN_SIZE,
        topMargin=MARGIN_SIZE, bottomMargin=MARGIN_SIZE)
    main_frame = Frame(MARGIN_SIZE, MARGIN_SIZE,
        PAGE_SIZE[0] - 2 * MARGIN_SIZE, PAGE_SIZE[1] - 2 * MARGIN_SIZE,
        leftPadding = 0, rightPadding = 0, bottomPadding = 0,
        topPadding = 0, id = 'main_frame')
    main_template = PageTemplate(id = 'main_template', frames = [main_frame])
    pdf_doc.addPageTemplates([main_template])

    pdf_doc.build(story)