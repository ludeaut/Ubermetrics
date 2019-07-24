# Ubermetrics.py
# Contains the class corresponding to the application

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
        self.spreadsheetIdStack = self.temp["spreadsheetIdStack"]
        if self.spreadsheetIdStack != '':
            self.spreadsheetIdStackField.setStringValue_(self.spreadsheetIdStack)
        self.sheetNameStack = self.temp["sheetNameStack"]
        if self.sheetNameStack != '':
            self.sheetNameStackField.setStringValue_(self.sheetNameStack)
        self.checkBoxSpreadsheetData.setState_(int(self.temp["checkBoxSpreadsheetData"]))
        if self.spreadsheetIdStack != '' and self.sheetNameStack != '' and int(self.checkBoxSpreadsheetData.state()) == 1:
            self.data = self.serviceSheets.spreadsheets().values().get(spreadsheetId=self.spreadsheetIdStack, range=self.sheetNameStack).execute().get('values')
        else:
            self.data = ''

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
            self.temp["viewId"] = self.viewIdTextField.stringValue()
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
    def setSpreadsheetIDStack_(self, sender):
        if self.spreadsheetIdStackField.stringValue():
            self.spreadsheetIdStack = self.spreadsheetIdStackField.stringValue()
            self.temp["spreadsheetIdStack"] = str(self.spreadsheetIdStack)
        else:
            self.value = 'Enter a correct Spreadsheet Id'
            self.updateDisplay()

    @objc.IBAction
    def setSheetNameStack_(self, sender):
        if self.sheetNameStackField.stringValue():
            self.sheetNameStack = self.sheetNameStackField.stringValue()
            self.temp["sheetNameStack"] = str(self.sheetNameStack)
        else:
            self.value = 'This sheet does not exist'
            self.updateDisplay()

    @objc.IBAction
    def displayValue_(self, sender):
        if self.viewId != 0 and self.startDate <= self.endDate and self.metrics != '':
            if len(self.viewIdTextField.stringValue()) == 1 and int(self.checkBoxSpreadsheetData.state()) == 1:
                self.viewId = ','.join([client[ord(self.viewIdTextField.stringValue().lower())-97] for client in self.data[1:]]).replace(' ', '')
            self.value = ''
            viewIds = self.viewId.split(',')
            for viewId in viewIds:
                value = functions.get_value(self.serviceAnalytics, viewId, str(self.startDate), str(self.endDate), self.metrics, self.dimensions, self.sort, self.filters, self.maxResults)
                if value == 0.0:
                    self.value += functions.str_from_list([[value]])
                    self.values += [['', value]]
                else:
                    self.value += functions.str_from_list(value)
                    self.values += value
                with open(self.dirpath + 'temp.json', 'w') as f:
                    json.dump(self.temp, f, indent=4)
                self.value += '\n'
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
            self.temp["checkBoxNumericValues"] = self.checkBoxNumericValues.state()
            if int(self.checkBoxNumericValues.state()) == 1:
                functions.fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, [[value[1]] for value in self.values])
            else:
                functions.fill_spreadsheet(self.serviceSheets, self.spreadsheetId, self.sheetName, self.range, self.values)
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
        if self.value.count('\n') > 64:
            self.valueTextField.setStringValue_('\n'.join(self.value.split('\n')[:64])+'\n and ' + str(self.value.count('\n')-64) + ' others values...')
        else:
            self.valueTextField.setStringValue_(self.value)
        if self.spreadsheetIdStack != '' and self.sheetNameStack != '' and int(self.checkBoxSpreadsheetData.state()) == 1:
            self.data = self.serviceSheets.spreadsheets().values().get(spreadsheetId=self.spreadsheetIdStack, range=self.sheetNameStack).execute().get('values')

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
