# AeroAPI-Request
Python code snippet used to request Airline schedule data from the AeroAPI
Used for collecting data for Aircraft Leasing and Finance Specialist Diploma Fleet Planning Module

The code requests the user for a number of calls to AeroAPI 2.0 (default max is set to 50)
Initial lines fo code set dates at which the schedule data will be pulled from and which Airline to request info on
Example shown is All Nippon Airways (ANA)

Data is then converted to JSON format for posterit
Further filtering was used to reduce data to ANA Wings flights only
This information was then inserted into a .xlsx file for further evaluation

This code is rough and ready to allow fast and curated data collection

Shawn McCormack 2021
