#-------------------------------------------------------------------------------
# Name:        generate_report.py
# Purpose:
#
# Author:      Dustin Evans
#
# Created:     03/05/2016
# Copyright:   (c) Dustin Evans 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import json, sys, getopt, os
import excel_generator
#import pdf_generator

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def main(argv):
    # Get and parse command-line arguments
    try:
        opts, args = getopt.getopt(argv, "i:o:t:")
    except getopt.GetoptError:
        print("Usage: generate_report.py -i </vehicles JSON file> -o <output .pdf/.xlsx file> -t <optional timezone>")
        sys.exit(2)

    input_filename = None
    output_filename = None
    timezone = "America/New_York"
    for opt, arg in opts:
        if opt == "-i":
            input_filename = arg
        elif opt == "-o":
            output_filename = arg
        elif opt == "-t":
            timezone = arg

    # Must have input_filename
    if input_filename != None and output_filename != None:
        # Send /vehicles request to localhost
        vehicles_data = read_json_file(input_filename)

        if vehicles_data != None:
            # Write report base on output_filename extension
            filepath, file_extension = os.path.splitext(output_filename)
            excel_generator.write_data(vehicles_data, output_filename, timezone)
            #if file_extension == '.pdf':
            #    pdf_generator.write_data(vehicles_data, output_filename, timezone)
            #else:
            #    excel_generator.write_data(vehicles_data, output_filename, timezone)

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def read_json_file(filename):
    try:
        fp = open(filename, 'r')
        data = json.load(fp)
        fp.close()
        return data
    except IOError as e:
        print('ERROR: Could not read ' + filename)
        return None

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
if __name__ == '__main__':
    main(sys.argv[1:])
