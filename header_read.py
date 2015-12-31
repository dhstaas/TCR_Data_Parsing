import os
import csv

#headerFile = raw_input("Please enter the full directory and name of the Station Header file used by TCE: ")
headerFile = r"S:\gis_data\Transportation\Projects\Traffic_Monitoring\Config\2015\Station_Header.csv"
header = open(headerFile)

dictionary = dict((rows[7],rows[1]) for rows in header)


print dictionary

