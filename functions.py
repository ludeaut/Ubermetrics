# Functions.py
# Contains the functions using the Google APIs and some utilities functions

from apiclient.discovery import build
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import json
import argparse
import sys
from datetime import date, timedelta
import os
import re

def get_service(api_name, api_version, scope, client_secrets_path):
    """Get a service that communicates to a Google API.

    Args:
    api_name: string The name of the api to connect to.
    api_version: string The api version to connect to.
    scope: A list of strings representing the auth scopes to authorize for the
    connection.
    client_secrets_path: string A path to a valid client secrets file.

    Returns:
    A service that is connected to the specified API.
    """
    # Parse command-line arguments.
    parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter, parents=[tools.argparser])
    flags = parser.parse_args([])

    # Set up a Flow object to be used if we need to authenticate.
    flow = client.flow_from_clientsecrets(client_secrets_path, scope=scope, message=tools.message_if_missing(client_secrets_path))

    # Prepare credentials, and authorize HTTP object with them.
    # If the credentials don't exist or are invalid run through the native client
    # flow. The Storage object will ensure that if successful the good
    # credentials will get written back to a file.
    dirpath = '/'.join(client_secrets_path.split('/')[:-1]) + '/'
    storage = file.Storage(dirpath + api_name + '.dat')
    credentials = storage.get()
    if credentials is None or credentials.invalid:
        credentials = tools.run_flow(flow, storage, flags)
    http = credentials.authorize(http=httplib2.Http())

    # Build the service object.
    service = build(api_name, api_version, http=http)

    return service

def get_value(service, view_id, start_date, end_date, metric, dimensions='', sort='', filters='', max_results=''):
    # Use the Analytics Service Object to query the Core Reporting API
    # for the given metric over the given data range and for the given view
    options = {};
    if dimensions != '':
        options['dimensions'] = 'ga:' + dimensions.replace(',', ', ga:')
    if sort != '':
        sort = sort.split('_');
        if sort[1] == 'desc':
            options['sort'] =  '-'
            options['sort'] += 'ga:' + sort[0]
        if filters != '':
            options['filters'] = 'ga:' + filters.replace(',', ',ga:').replace(';', ';ga:')
        if max_results != '':
            options['max_results'] = max_results
    metrics = 'ga:' + metric.replace(',', ', ga:')
    dimensionRows = service.data().ga().get(
                                            ids='ga:' + str(view_id),
                                            start_date=start_date,
                                            end_date=end_date,
                                            metrics=metrics,
                                            **options).execute().get('rows')
    if dimensionRows != None:
        return dimensionRows
    return 0.0

def fill_spreadsheet(service, spreadsheetId, sheetName, range, values):
    # Range to update in the Google Spreadsheet
    SAMPLE_RANGE_TO_UPDATE = sheetName + '!' + range

    value_input_option = 'USER_ENTERED'

    value_range_body = {
        'range':SAMPLE_RANGE_TO_UPDATE, 'majorDimension':'ROWS', 'values':values
    }

    service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range=SAMPLE_RANGE_TO_UPDATE, valueInputOption=value_input_option, body=value_range_body).execute()

def str_from_list(list):
    return str(list).replace('],', ']\n')[1:-1]

def get_data_from_spreadsheet(columns, service, spreadsheetId, sheetName):
    data_from_spreadsheet = []
    for column in columns:
        if is_A1_notation_range(str(column)):
            data = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=sheetName + '!' + str(column)).execute().get('values')
            data_from_spreadsheet.append(','.join([item[0] for item in data]).replace(' ', ''))
        else:
            data_from_spreadsheet.append(column)
    return data_from_spreadsheet

def is_A1_notation_range(str):
    pattern = "([A-Z]?[1-9][0-9]*)|([A-Z])"
    str = str.split(':')
    for string in str:
        if re.search(pattern, string) == None:
            return False
    if re.search('[A-Z]([1-9][0-9]*)', str[0]) != None and re.search('[A-Z]([1-9][0-9]*)', str[0]).group() == str[0]:
        return True
    elif re.search('[A-Z]', str[0]) != None and re.search('[A-Z]', str[0]).group() == str[0]:
        if len(str) == 1:
            return False
        return re.search("[A-Z]([1-9][0-9]*)?", str[1]) != None and re.search("[A-Z]([1-9][0-9]*)?", str[1]).group() == str[1]
    else:
        if len(str) == 1:
            return False
        return re.search("[A-Z]?[1-9][0-9]*", str[1]) != None and re.search("[A-Z]?[1-9][0-9]*", str[1]).group() == str[1]
