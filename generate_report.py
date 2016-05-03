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
import pdf_generator

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
            # Check response and convert response for write_data
            data_dict = convert_response(vehicles_data, timezone)

            # Write report base on output_filename extension
            filepath, file_extension = os.path.splitext(output_filename)
            if file_extension == '.pdf':
                pdf_generator.write_data(data_dict, output_filename)
            else:
                excel_generator.write_data(data_dict, output_filename)

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
def convert_response(response, timezone):
    data_dict = {'vehicles': []}
    for bay in response:
        vehicle_list = response[bay]['vehicles']
        for vehicle in vehicle_list:
            vehicle['bay'] = bay
            vehicle['timezone'] = timezone
            vehicle['snapshot'] = 'C:/baytracker/webroot' + vehicle['snapshot']
            data_dict['vehicles'].append(vehicle)

    return data_dict

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
if __name__ == '__main__':
    main(sys.argv[1:])
