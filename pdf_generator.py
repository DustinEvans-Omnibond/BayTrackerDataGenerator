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

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def write_data(data_dict, dest_filename):
    canvas = canvas.Canvas(dest_filename, pagesize=letter)
    width, height = letter



#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def draw_header(source_canvas, letter_width, letter_height):
    headers = ['Bay', 'Start Time', 'Service Time', 'Wait Time', 'Total Customer Time', 'Snapshot']
    for i in xrange(0, len(headers)):
        x = i * 100
        if i != len(headers)-1:
            # Text header
            pass
        else:
            # Snapshot header
            pass
