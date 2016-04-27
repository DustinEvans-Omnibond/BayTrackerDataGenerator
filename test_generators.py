#-------------------------------------------------------------------------------
# Name:        test_generators.py
# Purpose:
#
# Author:      Dustin Evans
#
# Created:     27/04/2016
# Copyright:   (c) Dustin Evans 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import json, sys, getopt
import excel_generator

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
def main(argv):
    # Get and parse command-line arguments
    try:
        opts, args = getopt.getopt(argv, "f:")
    except getopt.GetoptError:
        print("Usage: test_generators.py -f </vehicles JSON file>")
        sys.exit(2)

    source_filename = None
    for opt, arg in opts:
        if opt == "-f":
            source_filename = arg

    # Must have source_filename
    if source_filename != None:

        # Send /vehicles request to localhost
        vehicles_data = read_json_file(source_filename)

        if vehicles_data != None:
            # Check response and convert response for write_data
            data_dict = convert_response(vehicles_data)

            # Write data to Excel and PDF
            excel_generator.write_data(data_dict, 'output.xlsx')
            #pdf_generator.write_data(data_dict, 'output.pdf')

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
def convert_response(response):
    data_dict = {'vehicles': []}
    for bay in response:
        vehicle_list = response[bay]['vehicles']
        for vehicle in vehicle_list:
            #obj = {}
            #obj['bay'] = bay
            #obj['t_enter'] = vehicle['t_enter']
            #obj['t_leave'] = vehicle['t_leave']
            #obj['t_queue_enter'] = vehicle['t_queue_enter']
            #obj['snapshot'] = 'C:/baytracker/webroot' + vehicle['snapshot']
            #data_dict['vehicles'].append(obj)
            vehicle['bay'] = bay
            vehicle['snapshot'] = 'C:/baytracker/webroot' + vehicle['snapshot']
            data_dict['vehicles'].append(vehicle)

    return data_dict

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
if __name__ == '__main__':
    main(sys.argv[1:])
