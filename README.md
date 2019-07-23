# Ubermetrics

Ubermetrics is a little software written in Python and with a Cocoa interface built using the pyobjc library.
Using Google Analytics API and Google Spreadsheets API, it allows you to get Google Analytics Data from your accounts
and eventually to paste it in one of your Google Spreadsheets.

# What is in there?
* Ubermetrics.py the python script describing the software
* Ubermetrics.xib the Cocoa interface
* Ubermetrics.xcodeproj the xcode project used to build the Cocoa interface
* setup.py the python script building the application
* temp.json a temporary json file who records the previous entered value

# Before launching
* The application is written in Python3 so be sure to have it installed.
* You may also need to install some libraries. I don't remember every one I installed so ask or tell me if I forgot one.
Here are the commands to run:
  * sudo pip3 install --upgrade google-api-python-client
  * sudo pip3 install py2app
* In order to use the different API, you will have to start a Google project and to create credentials.
Everything you need to know to do it is here: https://console.developers.google.com/apis/credentials.
Don't forget to add the JSON files in the working directory.
* Still in the Google project, you will need to enable Google Analytics Reporting and Google Sheets in the APIs.

# How to build the application
1. python setup.py py2app -A # To run once at the beginning and then only if you modify the Cocoa interface.
2. ./dist/Ubermetrics.app/Contents/MacOS/Ubermetrics # To launch the application.

# Todos (without specific order)
* Add documentation
* Add requirement indication on the UI
* Add segments filtering
* Add possibility to fill fields from a spreadsheet
* Add some conditions on input

# Note(s)
* If all the required fields are filled but no value is displayed, there are multiple possibilities but the most common are:
 * You made a typing mistake,
 * You entered at least one value which doesn't exist
 * You entered dimensions and/or metrics which can't be used together
 * You added filters and/or sort conditions on metrics/dimensions unused


# Documentation
* All dimensions and metrics: https://developers.google.com/analytics/devguides/reporting/core/dimsmets