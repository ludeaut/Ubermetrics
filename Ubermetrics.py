# Ubermetrics.py
# Contains the class corresponding to the application

from apiclient.discovery import build
from googleapiclient.errors import HttpError
import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import json
import argparse
import sys
from datetime import date, timedelta
import os

import functions # Python script containing the following functions: get_service, get_value,
# fill_spreadsheet and str_from_list

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

    checkBoxNumericValues = objc.IBOutlet()

    spreadsheetIdStackField = objc.IBOutlet()
    sheetNameStackField = objc.IBOutlet()
    checkBoxSpreadsheetData = objc.IBOutlet()

    def windowDidLoad(self):
        NSWindowController.windowDidLoad(self)

        # Path to the files
        self.dirpath = str(os.getcwd()) + '/../../../../'

        # Define the auth scopes to request.
        scopeSheets = ['https://www.googleapis.com/auth/spreadsheets']
        scopeAnalytics = ['https://www.googleapis.com/auth/analytics.readonly']

        # Authenticate and construct services.
        self.serviceAnalytics = functions.get_service('analytics', 'v3', scopeAnalytics, self.dirpath + 'credentials.json')
        self.serviceSheets = functions.get_service('sheets', 'v4', scopeSheets, self.dirpath + 'credentials.json')

        # Value to display
        self.value = ''

        # Get previous entered values
        with open(self.dirpath + 'temp.json', 'r+') as f:
    	       self.temp = json.load(f)

        # Get data already in the first sheet
        self.checkBoxSpreadsheetData.setState_(int(self.temp["checkBoxSpreadsheetData"]))
        self.spreadsheetIdStack = self.temp["spreadsheetIdStack"]
        if self.spreadsheetIdStack != '':
            self.spreadsheetIdStackField.setStringValue_(self.spreadsheetIdStack)
        elif int(self.checkBoxSpreadsheetData.state()) == 1:
            self.spreadsheetIdStackField.setStringValue_('Spreadsheet ID - Required -')
        self.sheetNameStack = self.temp["sheetNameStack"]
        if self.sheetNameStack != '':
            self.sheetNameStackField.setStringValue_(self.sheetNameStack)
        elif int(self.checkBoxSpreadsheetData.state()) == 1:
            self.sheetNameStackField.setStringValue_('Sheet name - Required -')

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
        if self.maxResults != 0:
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
        self.checkBoxNumericValues.setState_(int(self.temp["checkBoxNumericValues"]))

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
        if self.viewIdTextField.stringValue(): #.isnumeric()
            self.viewId = self.viewIdTextField.stringValue()
            self.temp["viewId"] = str(self.viewId)
            # self.value = 'Enter a number please'
        else:
            self.viewId = 0
            self.temp["viewId"] = str(self.viewId)
            self.updateDisplay()

    @objc.IBAction
    def enterMetrics_(self, sender):
        if self.metricsTextField.stringValue():
            self.metrics = self.metricsTextField.stringValue()
            self.temp["metrics"] = str(self.metrics)
            # self.value = 'This metric does not exist'
        else:
            self.metrics = ""
            self.temp["metrics"] = str(self.metrics)
            self.updateDisplay()

    @objc.IBAction
    def enterDimensions_(self, sender):
        if self.dimensionsTextField.stringValue():
            self.dimensions = self.dimensionsTextField.stringValue()
            self.temp["dimensions"] = str(self.dimensions)
            # self.value = 'This dimension does not exist'
        else:
            self.dimensions = ""
            self.temp["dimensions"] = str(self.dimensions)
            self.updateDisplay()

    @objc.IBAction
    def enterSort_(self, sender):
        if self.sortTextField.stringValue():
            self.sort = self.sortTextField.stringValue()
            self.temp["sort"] = str(self.sort)
            # self.value = 'This metric does not exist'
        else:
            self.sort = ""
            self.temp["sort"] = str(self.sort)
            self.updateDisplay()

    @objc.IBAction
    def enterFilters_(self, sender):
        if self.filtersTextField.stringValue():
            self.filters = self.filtersTextField.stringValue()
            self.temp["filters"] = str(self.filters)
            # self.value = 'Enter a  correct filter'
        else:
            self.filters = ""
            self.temp["filters"] = str(self.filters)
            self.updateDisplay()

    @objc.IBAction
    def setMaxResults_(self, sender):
        if self.maxResultsTextField.stringValue():
            if int(self.maxResultsTextField.stringValue()) > 0 and int(self.maxResultsTextField.stringValue()) < 1000:
                self.maxResults = int(self.maxResultsTextField.stringValue())
                self.temp["maxResults"] = str(self.maxResults)
            else:
                self.value = 'Choose a number between 1 and 1000'
                self.updateDisplay()
        else:
            self.maxResults = 0
            self.temp["maxResults"] = str(self.maxResults)
            self.updateDisplay()

    @objc.IBAction
    def setSpreadsheetID_(self, sender): # RegEx ([a-zA-Z0-9-_]+)
        if self.spreadsheetIdTextField.stringValue():
            self.spreadsheetId = self.spreadsheetIdTextField.stringValue()
            self.temp["spreadsheetId"] = str(self.spreadsheetId)
            # self.value = 'Enter a correct Spreadsheet Id'
        else:
            self.spreadsheetId = ""
            self.temp["spreadsheetId"] = str(self.spreadsheetId)
            self.updateDisplay()

    @objc.IBAction
    def setSheetName_(self, sender):
        if self.sheetNameTextField.stringValue():
            self.sheetName = self.sheetNameTextField.stringValue()
            self.temp["sheetName"] = str(self.sheetName)
            # self.value = 'This sheet does not exist'
        else:
            self.sheetName = ""
            self.temp["sheetName"] = str(self.sheetName)
            self.updateDisplay()

    @objc.IBAction
    def setRange_(self, sender):
        if self.rangeTextField.stringValue():
            self.range = self.rangeTextField.stringValue()
            self.temp["range"] = str(self.range)
            # self.value = 'Enter a correct range'
        else:
            self.range = ""
            self.temp["range"] = str(self.range)
            self.updateDisplay()

    @objc.IBAction
    def setSpreadsheetIDStack_(self, sender): # RegEx ([a-zA-Z0-9-_]+)
        if self.spreadsheetIdStackField.stringValue():
            self.spreadsheetIdStack = self.spreadsheetIdStackField.stringValue()
            self.temp["spreadsheetIdStack"] = str(self.spreadsheetIdStack)
            # self.value = 'Enter a correct Spreadsheet Id'
        else:
            self.spreadsheetIdStack = ""
            self.temp["spreadsheetIdStack"] = str(self.spreadsheetIdStack)
            self.updateDisplay()
        self.refreshArguments()

    @objc.IBAction
    def setSheetNameStack_(self, sender):
        if self.sheetNameStackField.stringValue():
            self.sheetNameStack = self.sheetNameStackField.stringValue()
            self.temp["sheetNameStack"] = str(self.sheetNameStack)
            # self.value = 'This sheet does not exist'
        else:
            self.sheetNameStack = ""
            self.temp["sheetNameStack"] = str(self.sheetNameStack)
            self.updateDisplay()
        self.refreshArguments()

    @objc.IBAction
    def displayValue_(self, sender):
        if self.viewId != 0 and self.startDate <= self.endDate and self.metrics != '':
            try:
                if int(self.checkBoxSpreadsheetData.state()) == 1:
                    self.viewId, self.metrics, self.dimensions, self.sort, self.filters, self.maxResults \
                    = functions.get_data_from_spreadsheet([self.viewId, self.metrics, self.dimensions, self.sort, \
                    self.filters, self.maxResults], self.serviceSheets, self.spreadsheetIdStack, self.sheetNameStack)
                self.value = ''
                viewIds = self.viewId.split(',')
                for viewId in viewIds:
                    value = functions.get_value(self.serviceAnalytics, viewId, str(self.startDate), str(self.endDate), self.metrics, self.dimensions, self.sort, self.filters, self.maxResults)
                    self.value += functions.str_from_list(value) + '\n'
                    if len(value[0]) == 1:
                        value[0].append('')
                        value[0].reverse()
                    self.values += value
                with open(self.dirpath + 'temp.json', 'w') as f:
                    json.dump(self.temp, f, indent=4)
            except (HttpError) as e:
                self.value = json.loads(e.content.decode('utf-8'))['error']['message']
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
        if self.values != [] and self.spreadsheetId != '' and self.sheetName != '' and self.range != '':
            self.temp["checkBoxNumericValues"] = self.checkBoxNumericValues.state()
            if int(self.checkBoxNumericValues.state()) == 1:
                functions.fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, [[value[1]] for value in self.values])
            else:
                functions.fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, self.values)
            with open(self.dirpath + 'temp.json', 'w') as f:
                json.dump(self.temp, f, indent=4)
            self.value = "Done!"
        else:
            if self.values == []:
                self.value = 'You did not add values for the moment'
            elif self.spreadsheetId == '':
                self.spreadsheetIdTextField.setStringValue_('Spreadsheet ID - Required -')
                self.value = 'You must enter a spreadsheet ID'
            elif self.sheetName == '':
                self.sheetNameTextField.setStringValue_('Sheet name - Required -')
                self.value = 'You must enter a sheet name'
            else:
                self.rangeTextField.setStringValue_('Range (e.g. A1:B2) - Required -')
                self.value = 'You must enter a range'
        self.updateDisplay()

    @objc.IBAction
    def emptyValues_(self, sender):
        self.values = []
        self.value = 'Previous values have been erased'
        self.updateDisplay()

    def updateDisplay(self):
        if self.value.count('\n') > 64:
            self.valueTextField.setStringValue_('\n'.join(self.value.split('\n')[:64])+'\n and ' + str(self.value.count('\n')-64) + ' others values...')
        else:
            self.valueTextField.setStringValue_(self.value)

        if self.spreadsheetIdStack == '':
            if int(self.checkBoxSpreadsheetData.state()) == 1:
                self.spreadsheetIdStackField.setStringValue_('Spreadsheet ID - Required -')
            else:
                self.spreadsheetIdStackField.setStringValue_('Spreadsheet ID')
        if self.sheetNameStack == '':
            if int(self.checkBoxSpreadsheetData.state()) == 1:
                self.sheetNameStackField.setStringValue_('Sheet name - Required -')
            else:
                self.sheetNameStackField.setStringValue_('Sheet name')
        if self.viewId == 0:
            self.viewIdTextField.setStringValue_('View ID  - Required -')
        if self.metrics == '':
            self.metricsTextField.setStringValue_('Metric(s)  - Required -')
        if self.dimensions == '':
            self.dimensionsTextField.setStringValue_('Dimension(s)')
        if self.sort == '':
            self.sortTextField.setStringValue_('Sort choices')
        if self.filters == '':
            self.filtersTextField.setStringValue_('Filters')
        if self.maxResults == 0:
            self.maxResultsTextField.setStringValue_('Number max of results')
        if self.spreadsheetId == '':
            self.spreadsheetIdTextField.setStringValue_('Spreadsheet ID')
        if self.sheetName == '':
            self.sheetNameTextField.setStringValue_('Sheet name')
        if self.range == '':
            self.rangeTextField.setStringValue_('Range (e.g. A1:B2)')

    def refreshArguments(self):
        self.viewId = self.temp["viewId"]
        self.metrics = self.temp["metrics"]
        self.dimensions = self.temp["dimensions"]
        self.sort = self.temp["sort"]
        self.filters = self.temp["filters"]
        self.maxResults = str(self.temp["maxResults"])

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
