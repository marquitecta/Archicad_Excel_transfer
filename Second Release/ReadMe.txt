This program is an experimental script intended to demo the capabilities of the Archicad Python API to transfer properties from an Excel spreadsheet to an Archicad Plan. Its purpose is to expedite the use of ‘office standards for selected Objects and Elements. It should be considered a pre-Alpha version and assumed to contain ‘bugs’ and various anomalies. As such it should be ‘tested’ and used only with test or backup plans.
Notable deficiencies include:
1.0 It does not, in any way, conform to PEP 8.
2.0 Variables and functions naming are not descriptive and have been chosen for convenience and speed.
3.0 Comments are limited.
4.0 It has not been extensively tested.
5.0 Spelling errors are not caught and may result in unpredicted behavior.
6.0 Try/Except blocks are not used – But must be added.
7.0 Deviation from the expected protocol will cause unexpected and uncaught results.

This script uses Python 3.7 + and the add-ins of the Archicad Python API, Tkinter and xlrd. They must be pre-installed before running this script. See usage of PIP. There are several instructional videos on Youtube for using PIP. The test Excel SS has been included for your convenience.

Testers are invited to comment, revise and extend this script and to report “bugs”. The most notable revisions will be included in a future release.

As a test Plan, I have used the Archicad ‘Sample house Plan – version 23’ and the Excel SS references the objects in that plan. Unfortunately, that plan is too large to include in GitHub but I have included the xlsx files to load the used schedules and property groups. Remember to verify that all properties and groups have been properly classified. The script does not catch those errors. 

Revision: 8/9/2020
Sheetsheets now use data types to define properties -- see video #3 for description.