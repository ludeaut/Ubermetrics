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

from Cocoa import *
from Foundation import NSObject

class UbermetricsController(NSWindowController):
    valueTextField = objc.IBOutlet()
    startDatePicker = objc.IBOutlet()
    endDatePicker = objc.IBOutlet()
    viewIdTextField = objc.IBOutlet()
    metricsTextField = objc.IBOutlet()
    dimensionsTextField = objc.IBOutlet()
    sortTextField = objc.IBOutlet()
    filtersTextField = objc.IBOutlet()
    maxResultsTextField = objc.IBOutlet()

    spreadsheetIdTextField = objc.IBOutlet()
    sheetNameTextField = objc.IBOutlet()
    rangeTextField = objc.IBOutlet()

    checkBox = objc.IBOutlet()

    def windowDidLoad(self):
        NSWindowController.windowDidLoad(self)

        # Path to the files
        self.dirpath = str(os.getcwd()) + '/../../../../'

        # Define the auth scopes to request.
        scopeSheets = ['https://www.googleapis.com/auth/spreadsheets']
        scopeAnalytics = ['https://www.googleapis.com/auth/analytics.readonly']

        # Authenticate and construct services.
        self.serviceSheets = get_service('sheets', 'v4', scopeSheets, self.dirpath + 'credentials.json')
        self.serviceAnalytics = get_service('analytics', 'v3', scopeAnalytics, self.dirpath + 'client_secrets.json')

        # Value to display
        self.value = ''

        # Get previous entered values
        with open(self.dirpath + 'temp.json', 'r+') as f:
    	       self.temp = json.load(f)

        # Arguments for getValue function
        self.viewId = self.temp["viewId"]
        if self.viewId != 0:
            self.viewIdTextField.setStringValue_(self.viewId)

        startDateStr = self.temp["startDate"].split('-')
        self.startDate = date(int(startDateStr[0]), int(startDateStr[1]), int(startDateStr[2]))
        endDateStr = self.temp["endDate"].split('-')
        self.endDate = date(int(endDateStr[0]), int(startDateStr[1]), int(endDateStr[2]))
        self.startDatePicker.setDateValue_(self.startDate)
        self.endDatePicker.setDateValue_(self.endDate)

        self.metrics = self.temp["metrics"]
        if self.metrics != '':
            self.metricsTextField.setStringValue_(self.metrics)

        self.dimensions = self.temp["dimensions"]
        if self.dimensions != '':
            self.dimensionsTextField.setStringValue_(self.dimensions)
        self.sort = self.temp["sort"]
        if self.sort != '':
            self.sortTextField.setStringValue_(self.sort)
        self.filters = self.temp["filters"]
        if self.filters != '':
            self.filtersTextField.setStringValue_(self.filters)
        self.maxResults = self.temp["maxResults"]
        if self.maxResults != '':
            self.maxResultsTextField.setStringValue_(self.maxResults)

        # Arguments for fillSpreadsheet function
        self.spreadsheetId = self.temp["spreadsheetId"]
        if self.spreadsheetId != '':
            self.spreadsheetIdTextField.setStringValue_(self.spreadsheetId)
        self.sheetName = self.temp["sheetName"]
        if self.sheetName != '':
            self.sheetNameTextField.setStringValue_(self.sheetName)
        self.range = self.temp["range"]
        if self.range != '':
            self.rangeTextField.setStringValue_(self.range)

        # Checkbox to choose to display only numeric values or no
        self.checkBox.setState_(int(self.temp["checkBox"]))

        # self.startDatePicker.setTimeZone_(NSTimeZone.localTimeZone())
        # NSLog(str(self.startDatePicker.timeZone()))

        # Column values for all clients
        self.values = []

    @objc.IBAction
    def setStartDate_(self, sender):
        if self.startDatePicker.dateValue() <= self.endDatePicker.dateValue():
            startDatePickerValue = str(self.startDatePicker.dateValue())[:10].split('-')
            self.startDate = date(int(startDatePickerValue[0]), int(startDatePickerValue[1]), int(startDatePickerValue[2]))
            self.startDate += timedelta(days=1) # Timezone issue to deal with
            self.temp["startDate"] = str(self.startDate)
        else:
            self.value = 'Choose a start date before the end date'
            self.updateDisplay()

    @objc.IBAction
    def setEndDate_(self, sender):
        if self.startDatePicker.dateValue() <= self.endDatePicker.dateValue():
            endDatePickerValue = str(self.endDatePicker.dateValue())[:10].split('-')
            self.endDate = date(int(endDatePickerValue[0]), int(endDatePickerValue[1]), int(endDatePickerValue[2]))
            self.endDate += timedelta(days=1) # Timezone issue to deal with
            self.temp["endDate"] = str(self.endDate)
        else:
            self.value = 'Choose a end date after the start date'
            self.updateDisplay()

    @objc.IBAction
    def enterViewID_(self, sender): # Must have 7 to 9 figures to write
        if self.viewIdTextField.stringValue().isnumeric():
            self.viewId = self.viewIdTextField.stringValue()
            self.temp["viewId"] = str(self.viewId)
        else:
            self.value = 'Enter a number please'
            self.updateDisplay()

    @objc.IBAction
    def enterMetrics_(self, sender):
        if self.metricsTextField.stringValue():
            self.metrics = self.metricsTextField.stringValue()
            self.temp["metrics"] = str(self.metrics)
        else:
            self.value = 'This metric does not exist'
            self.updateDisplay()

    @objc.IBAction
    def enterDimensions_(self, sender):
        if self.dimensionsTextField.stringValue():
            self.dimensions = self.dimensionsTextField.stringValue()
            self.temp["dimensions"] = str(self.dimensions)
        else:
            self.value = 'This dimension does not exist'
            self.updateDisplay()

    @objc.IBAction
    def enterSort_(self, sender):
        if self.sortTextField.stringValue():
            self.sort = self.sortTextField.stringValue()
            self.temp["sort"] = str(self.sort)
        else:
            self.value = 'This metric does not exist'
            self.updateDisplay()

    @objc.IBAction
    def enterFilters_(self, sender):
        if self.filtersTextField.stringValue():
            self.filters = self.filtersTextField.stringValue()
            self.temp["filters"] = str(self.filters)
        else:
            self.value = 'Enter a  correct filter'
            self.updateDisplay()

    @objc.IBAction
    def setMaxResults_(self, sender):
        if int(self.maxResultsTextField.stringValue()) > 0 and int(self.maxResultsTextField.stringValue()) < 1000:
            self.maxResults = int(self.maxResultsTextField.stringValue())
            self.temp["maxResults"] = str(self.maxResults)
        else:
            self.value = 'Choose a number between 1 and 1000'
            self.updateDisplay()

    @objc.IBAction
    def setSpreadsheetID_(self, sender):
        if self.spreadsheetIdTextField.stringValue():
            self.spreadsheetId = self.spreadsheetIdTextField.stringValue()
            self.temp["spreadsheetId"] = str(self.spreadsheetId)
        else:
            self.value = 'Enter a correct Spreadsheet Id'
            self.updateDisplay()

    @objc.IBAction
    def setSheetName_(self, sender):
        if self.sheetNameTextField.stringValue():
            self.sheetName = self.sheetNameTextField.stringValue()
            self.temp["sheetName"] = str(self.sheetName)
        else:
            self.value = 'This sheet does not exist'
            self.updateDisplay()

    @objc.IBAction
    def setRange_(self, sender):
        if self.rangeTextField.stringValue():
            self.range = self.rangeTextField.stringValue()
            self.temp["range"] = str(self.range)
        else:
            self.value = 'Enter a correct range'
            self.updateDisplay()

    @objc.IBAction
    def displayValue_(self, sender):
        if self.viewId != 0 and self.startDate <= self.endDate and self.metrics != '':
            value = get_value(self.serviceAnalytics, self.viewId, str(self.startDate), str(self.endDate), self.metrics, self.dimensions, self.sort, self.filters, self.maxResults)
            if value == 0.0:
                self.value = str_from_list([[value]])
                self.values += [['', value]]
            else:
                self.value = str_from_list(value)
                self.values += value
            with open(self.dirpath + 'temp.json', 'w') as f:
                json.dump(self.temp, f, indent=4)
        else:
            if self.viewId == 0:
                self.value = 'You must enter a view ID'
            elif self.startDate > self.endDate:
                self.value = 'Choose a possible date range'
            else:
                self.value = 'You must enter a metric'
        self.updateDisplay()

    @objc.IBAction
    def fillSpreadsheet_(self, sender):
        if self.values != []:
            self.temp["checkBox"] = self.checkBox.state()
            if str(self.checkBox.state()) == '1':
                fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, [[value[1]] for value in self.values])
            else:
                fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, self.values)
            with open(self.dirpath + 'temp.json', 'w') as f:
                json.dump(self.temp, f, indent=4)
        else:
            self.value = 'You did not add values for the moment'
        self.updateDisplay()

    @objc.IBAction
    def emptyValues_(self, sender):
        self.values = []
        self.value = 'Previous values have been erased'
        self.updateDisplay()

    def updateDisplay(self):
        self.valueTextField.setStringValue_(self.value)

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

def main():

    # Define the auth scopes to request.
    scopeSheets = ['https://www.googleapis.com/auth/spreadsheets']
    scopeAnalytics = ['https://www.googleapis.com/auth/analytics.readonly']

    # Authenticate and construct service.
    # Will create a file token.pickle.
    serviceSheets = get_service('sheets', 'v4', scopeSheets, '/Users/thomas/Documents/Ubermetrics/credentials.json')
    serviceAnalytics = get_service('analytics', 'v3', scopeAnalytics, '/Users/thomas/Documents/Ubermetrics/client_secrets.json')

    # The ID and range of the Google spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1BMrOh20xeuiMS_6t-Axxr_R09UYLREaHPIdr-N-7tvs'
    SAMPLE_RANGE_NAME = 'Sheet1!A1:T'

    # Data already on the spreadsheet
    allClients = serviceSheets.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute().get('values')

    print(get_value(serviceAnalytics, "48213793", "2019-06-01", "2019-06-30", "sessions,users", "channelGrouping", "sessions_desc", "sessions>100,users>1000;users<10000", 5))

    print('Done!')


if __name__ == '__main__':

    app = NSApplication.sharedApplication()

    # Initiate the contrller with a XIB
    viewController = UbermetricsController.alloc().initWithWindowNibName_("Ubermetrics")

    # Show the window
    viewController.showWindow_(viewController)

    # Bring app to top
    NSApp.activateIgnoringOtherApps_(True)

    from PyObjCTools import AppHelper
    AppHelper.runEventLoop()
