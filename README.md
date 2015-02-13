# Watson Dashboard

## Overview

The Watson Dashboard Generator was created to assist Binghamton University staff by parsing through XML files of arbitrary size and outputting a useful "dashboard" spreadsheet. The program utilizes the [Apache POI](http://poi.apache.org/) and [dom4j](http://dom4j.sourceforge.net/) libraries (for parsing XML / Microsoft Excel documents), along with [WindowBuilder](https://eclipse.org/windowbuilder/) (for GUI designer).

## How to Use
1. [Click here](http://matthewdiana.com/downloads/DashboardGenerator.jar) to download the .jar file.
2. Open DashboardGenerator.jar, click "Browse" and select your input data spreadsheet (NOTE: you can use [WatsonDecember2014_revised](https://github.com/MatthewDiana/WatsonDashboard/raw/master/WatsonDashboard/WatsonDecember2014_revised.xlsx) as sample input).
3.	Click "Create Dashboard" and a the dashboard spreadsheet should automatically open in your default .xlsx viewer.

## Screenshots

### Simple Main GUI
![main interface](http://i.imgur.com/cJNAioK.png)

### Generated Faculty Sheet
![html report](http://i.imgur.com/CU7exA0.png)
