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
* Still in the Google project, you will need to enable Google Analytics Reporting and Google Sheets in the APIs.

# How to build the application
1. python setup.py py2app -A # To run once at the beginning and then only if you modify the Cocoa interface.
2. ./dist/Ubermetrics.app/Contents/MacOS/Ubermetrics # To launch the application.

# Todos (without specific order)
* Add documentation
* Add segments filtering
* Add possibility to get only numeric results
* Add some conditions on input

