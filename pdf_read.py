import pdfquery
import os
import xlsxwriter
import time #needed to audit

#############################################################################
#                                                                           #
#         Traffic Count PDF Parser Created by David Staas UCTC              #
#                                                                           #
#############################################################################

###################################
# Establish a working environment #
###################################

countDirectory = raw_input("Enter the directory where pdf versions of the Traffic Count Hourly Reports are located: ")
#countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol" #can set static directory for testing
os.chdir(countDirectory)
pdfFileList=[fn for fn in os.listdir(countDirectory) if fn.endswith('.pdf')] #creates a list of pdf files in the directory
workbookName = raw_input("Please enter the name of the Excel workbook to be generated: ") #establises output excel file

start_time = time.time() #start audit timer


countData = [] # Global list to store all the station information 


################################
#  checks the report type and  #
#  then passes the pdf to be   #
#  processed if it is the      #
#  desired report type         #
#                              #
# Returns countData which has  #
# multiple stations data in it #
################################

def stationDataScrape(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    if reportType(countPdf) == 2:
        countData.append((processCount(countPdf)))
    return countData 
 
#################################
#   Checks each pdf to see      #
#   what kind report it is      #
#                               #
#   Currently rudimentary as    #
#   it only checks # of pages   #
#################################
def reportType(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pageNum = pdf.doc.catalog['Pages'].resolve()['Count']
    if pageNum == 3:
        countType = 3 #"NYSDOT 3 Page Volume"
    elif pageNum == 2:
        countType = 2 #"NYSDOT 2 Page Volume"
    elif pageNum == 1:
        countType = 1 #"Class or Speed Count"
    else:
        countType = 4 #"Unknown Count Type"
    return  countType


########################
#  aggregates all of   #
#  the fields needed   #
#  for a station into  #
#  a list stationData  #
########################
def processCount(countPdf):
    stationData =[] #list where we are storing the count data for each station
    stationData.extend([(getStation(countPdf)),"Date", "Road Name", "From", "To", "Municipality", "Year", "Northing", "Easting",
                        (getAADT(countPdf)[0]), (getAADT(countPdf)[1]), (getPMPeak(countPdf)[0]), (getPMPeak(countPdf)[1]),
                        "Sp85_1", "Sp85_2", (getDirection(countPdf)[0]),(getDirection(countPdf)[1])]) 

    return stationData


###############
# Excel Setup #
###############
def stationToExcel(countData):
# Creates workbook, headings and format
    workbook = xlsxwriter.Workbook((str(os.curdir)[:-1]) + workbookName + ".xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.write('A1', 'Station', bold)
    worksheet.write('B1', 'Date', bold)
    worksheet.write('C1', 'Road_Name', bold)
    worksheet.write('D1', 'From', bold)
    worksheet.write('E1', 'To', bold)
    worksheet.write('F1', 'Municipality', bold)
    worksheet.write('G1', 'Year', bold)
    worksheet.write('H1', 'Northing', bold)
    worksheet.write('I1', 'Easting', bold)
    worksheet.write('J1', 'AADT_1', bold)
    worksheet.write('K1', 'AADT_2', bold)
    worksheet.write('L1', 'PM_45_1', bold)
    worksheet.write('M1', 'PM_45_2', bold)
    worksheet.write('N1', 'Sp85_1', bold)
    worksheet.write('O1', 'Sp85_2', bold)
    worksheet.write('P1', 'Dir_1', bold)
    worksheet.write('Q1', 'Dir_2', bold)

    row =1
    col = 0

    #iterates through each station stored in countData and adds it to the workbook 
    for station, date, RoadName, From, To, Municipality, Year, Northing, Easting, AADT_1, AADT_2, PM_45_1, PM_45_2, Sp_85_1, Sp_85_2, Dir_1, Dir_2 in (countData):
        worksheet.write(row, col,   station)
        worksheet.write(row, col + 1,   date)
        worksheet.write(row, col + 2,   RoadName)
        worksheet.write(row, col + 3,   From)
        worksheet.write(row, col + 4,   To)
        worksheet.write(row, col + 5,   Municipality)
        worksheet.write(row, col + 6,   Year)
        worksheet.write(row, col + 7,   Northing)
        worksheet.write(row, col + 8,   Easting)
        worksheet.write(row, col + 9,   AADT_1)
        worksheet.write(row, col + 10,   AADT_2)
        worksheet.write(row, col + 11,   PM_45_1)
        worksheet.write(row, col + 12,   PM_45_2)
        worksheet.write(row, col + 13,   Sp_85_1)
        worksheet.write(row, col + 14,   Sp_85_2)
        worksheet.write(row, col + 15,   Dir_1)
        worksheet.write(row, col + 16,   Dir_2)
        row += 1
        
        

###########
# Station #
###########

def getStation(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(1) #only need to load one page as it is the same on both (saves some time)
    station = pdf.pq('LTTextLineHorizontal:in_bbox("36.0, 580.368, 186.0, 610.368")').text() #x, y cords in points of the text we want
    station = station[-6:] #text line includes "station:" so we take just the last 6 chaaracters of the string
    return station
    
    
#############
# Direction #
#####################################
#                                   #
# Could reduce load here by         #
# only loading one page and then    #
# putting in the opposite           #
# direction                         #
#####################################


def getDirection(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    directionList = []
    pdf=pdfquery.PDFQuery(countPdf)
    for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get direction
        pdf.load(page)
        directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
        directionTMPList = directionTMP.split()
        directionList.extend((directionTMPList[-1:]))
    return directionList


##############
#    AADT    #
##############
''' Needs to be cleaned up, still searching for AADT rather than going to the specific cords of the AADT'''

def getAADT(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load()
    AADT_label = pdf.pq('LTTextLineHorizontal:contains("AADT")')
    left_corner = float(AADT_label.attr('x0'))
    bottom_corner = float(AADT_label.attr('y0'))
    #print left_corner, bottom_corner
    AADT = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -30, left_corner+150, bottom_corner)).text()
    #print AADT
    volumeList = AADT.split()
    return volumeList


#########
#PM Peak#
#########

def getPMPeak(countPdf):
    #pdf.load(1) #Do one page at a time as this returns all the values in the 4 to 5 column (days and the final avg)
    pmPeakList = []
    pdf=pdfquery.PDFQuery(countPdf)
    for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get PM peak
        pdf.load(page)
        pmPeak4_5 = pdf.pq('LTTextLineHorizontal:in_bbox("475.0, 137.0, 496.0, 504.0")').text()
        pmPeakListTMP = pmPeak4_5.split()
        pmPeakList.extend((pmPeakListTMP[-1:]))
    return pmPeakList

    

#############################################
#                    Tests                  #
#############################################
#print getAADT("0005.pdf")
#print processCount("0005.pdf")

#stationToExcel((processCount("0005.pdf")))
#stationToExcel(['860005', 'Date', 'Road Name', 'From', 'To', 'Municipality', 'Year', 'Northing', 'Easting', '1097', '1087', '80', '114', 'Sp85_1', 'Sp85_2', 'Northbound', 'Southbound'])

#############
# Main loop #
#############

for countPdf in pdfFileList:
    print "Reading " + countPdf
    count_time = time.time()
    stationDataScrape(countPdf)
    print "Finished " + countPdf + " in " + str(time.time() - count_time)
stationToExcel(countData) #sends countData to excel for format
os.startfile((str(os.curdir)[:-1]) + workbookName + ".xlsx") #opens output file

print "Completed in", time.time() - start_time
