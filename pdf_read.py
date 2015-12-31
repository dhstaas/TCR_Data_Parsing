import pdfquery
import os
import xlsxwriter
import time #needed to audit

#############################################################################
#                                                                           #
#                       Traffic Count PDF Parser                            #
#                              Created by                                   # 
#                             David Staas                                   #
#                                 UCTC                                      #
#                                                                           #
#############################################################################

###################################
# Establish a working environment #
###################################

#countDirectory = raw_input("Enter the directory where pdf versions of the Traffic Count Hourly Reports are located: ")
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol" #can set static directory for testing
os.chdir(countDirectory)
pdfFileList=[fn for fn in os.listdir(countDirectory) if fn.endswith('.pdf')] #creates a list of pdf files in the directory
peak_start = int(raw_input("Enter desired peak hour starting time (0 - 24 eg. enter 16 for 4PM):" ))
peak_end = int(raw_input("Enter desired peak hour ending time (0 - 24 eg. enter 17 for 5PM):" ))
workbookName = raw_input("Please enter the name of the Excel workbook to be generated: ") #establises output excel file

start_time = time.time() #start audit timer

'''directionList = []
volumeList = []
pmPeakList = []
station = ""
date = ""
roadname = "" '''


countData = [] # Global list to store all the station information 

####################################
# Multi Hour Peak Range Validation #
####################################
def peakRangeValid(peak_start, peak_end):
    global startMeridiem
    global endMeridiem
    global peakStartLabel
    global peakEndLabel
    global peakLabel
    validInput = True
    
    if peak_start > 12 and peak_start <= 24:
        peakStartLabel = peak_start - 12
        startMeridiem = "PM"
    elif peak_start >= 0:
        peakStartLabel = peak_start
        startMeridiem = "AM"
    else:
        validInput = False
        

    if peak_end > 12 and peak_start <=24:
        peakEndLabel = peak_end - 12
        endMeridiem =  "PM"
    elif peak_start >= 0:
        peakEndLabel = peak_end
        endMeridiem = "AM"
    else:
        validInput = False

    if startMeridiem == endMeridiem:
        peakLabel = startMeridiem + "_" + str(peakStartLabel) + "to" + str(peakEndLabel)
    else:
        peakLabel = startMeridiem + "_" + str(peakStartLabel) + "to" + endMeridiem + "_" + str(peakEndLabel)

    return validInput
    

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
        print "NYSDOT 3 page report not supported"
    elif pageNum == 2:
        countType = 2 #"NYSDOT 2 Page Volume"
    elif pageNum == 1:
        countType = 1 #"Class or Speed Count"
        print "Class and Speed Count not suppoted"
    else:
        countType = 4 #"Unknown Count Type"
        print "Unknown pdf. Is this a Count?"
    return  countType


########################
#  aggregates all of   #
#  the fields needed   #
#  for a station into  #
#  a list stationData  #
########################
def processCount(countPdf):
    stationData =[] #list where we are storing the count data for each station
    ###############################
    #populate the global variables#
    ###############################
    #individually#
    '''volumeList = getAADT(countPdf)
    pmPeakList = getPMPeak(countPdf)
    directionList = getDirection(countPdf)
    station = getStation(countPdf)'''
    #by page load type#
    '''getSinglePageData(countPdf)
    getMultiPageData(countPdf)'''

    getAllCountData(countPdf)
    
    stationData.extend([(station),(date), (roadName), (fromName), (toName), (municipality), (year), "Northing", "Easting",
                        (volumeList[0]), (volumeList[1]), (totalPeakList[0]), (totalPeakList[1]),
                        "Sp85_1", "Sp85_2", (directionList[0]),(directionList[1])]) 

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
    worksheet.write('L1', peakLabel + '_1', bold)
    worksheet.write('M1', peakLabel + '_2', bold)
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
        
#####################################
#        DATA EXTRACTION            #
#####################################

def getAllCountData(countPdf):
    global station
    global date
    global year
    global roadName
    global fromName
    global toName
    global municipality
    global directionList
    global volumeList
    global totalPeakList  

    directionList = []
    volumeList = []    
    totalPeakList = []
    
    pdf=pdfquery.PDFQuery(countPdf)
    for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get required data        
        pdf=pdfquery.PDFQuery(countPdf)
        pdf.load(page)
        if page == 0: #pulls single page data
            #############
            #  Station  #
            #############
            station = pdf.pq('LTTextLineHorizontal:in_bbox("36.0, 580.368, 186.0, 610.368")').text() #x, y cords in points of the text we want
            station = station[-6:] #text line includes "station:" so we take just the last 6 chaaracters of the string

            #############
            # Date/year #
            #############
            date = pdf.pq('LTTextLineHorizontal:in_bbox("35.999, 517.736, 160.999, 529.736")').text()
            date = date[-10:] #takes last 10 characters to create the date string
            year = date[-4:]

            #############
            # Road Name #
            #############
            roadName = pdf.pq('LTTextLineHorizontal:in_bbox("152.999, 547.496, 302.999, 559.496")').text()
            roadName = roadName[11:] #removes first 11 characters making up the label and returns the rest

            ############
            #   From   #
            ############
            fromName = pdf.pq('LTTextLineHorizontal:in_bbox("296.999, 547.496, 440.999, 559.496")').text()
            fromName = fromName[6:]

            ############
            #    To    #
            ############
            toName = pdf.pq('LTTextLineHorizontal:in_bbox("486, 547.496, 686, 559.496")').text()
            toName = toName[4:-7]

            ################
            # Municipality #
            ################
            municipality = pdf.pq('LTTextLineHorizontal:in_bbox("614, 537, 815, 549")').text()
            municipality = municipality.replace(":", " OF")

            #############
            # Driection #
            #############
            directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
            directionTMPList = directionTMP.split()
            direction1 = directionTMPList[-1:]
            if direction1[0] == "Northbound":
                direction2 = "Southbound"
            elif direction1[0] == "Eastbound":
                direction2 = "Westbound"
            else:
                direction1 = "Check PDF"
                direction2 = "Check PDF"
            directionList.extend((direction1[0], direction2))
            
            ########
            # Peak #
            ########

            peakTotal = 0
            left_corner = 98.0
            bottom_corner = 120.0
            right_corner = 121.5
            top_corner =  440.0
            columnWidth = 23.5

            for hour in range(peak_start, peak_end):

                peak = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % ((left_corner + (23.5 * peak_start)), bottom_corner, (right_corner+(23.5 * (peak_start))), top_corner)).text()
                peakList = peak.split()
                peakTotal += int(peakList.pop(-1))
                left_corner += columnWidth
                right_corner += columnWidth

            totalPeakList.append(peakTotal)    
            
            ##############
            #    AADT    #
            ##############
            AADT = pdf.pq('LTTextLineHorizontal:in_bbox("658.38, 67.428, 808, 97.428")').text()#no need to split as with pmPeak and others as only the value we want is supplied
            volumeList.append(AADT) 
        else:
            ########
            # Peak #
            ########
           
            peakTotal = 0
            #starting position of the hourly columns
            left_corner = 98.0
            bottom_corner = 120.0
            right_corner = 121.5
            top_corner =  440.0
            columnWidth = 23.5 # spacing between the hourly columns
            
            for hour in range(peak_start, peak_end):#takes the user defined input range and adds the hourl averages together

                peak = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % ((left_corner + (23.5 * peak_start)), bottom_corner, (right_corner+(23.5 * (peak_start))), top_corner)).text()
                peakList = peak.split()
                peakTotal += int(peakList.pop(-1))
                left_corner += columnWidth
                right_corner += columnWidth

            totalPeakList.append(peakTotal)
            '''#########
            #PM Peak#
            #########
            pmPeak4_5 = pdf.pq('LTTextLineHorizontal:in_bbox("475.0, 137.0, 496.0, 504.0")').text() #this has all values in the column (daily counts and the avg)
            pmPeakListTMP = pmPeak4_5.split()
            pmPeakList.extend((pmPeakListTMP[-1:])) #only adds the final number from the column which is the avg, ignoring the daily counts'''

            ##############
            #    AADT    #
            ##############
            AADT = pdf.pq('LTTextLineHorizontal:in_bbox("658.38, 67.428, 808, 97.428")').text()#no need to split as with pmPeak and others as only the value we want is supplied
            volumeList.append(AADT)


#########################
''' Single Page Loads'''
#########################

def getSinglePageData(countPdf, page):
    global station
    global directionList
    directionList = []
    ###########
    # Station #
    ###########
    
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(page) #only need to load one page as it is the same on both (saves some time)
    station = pdf.pq('LTTextLineHorizontal:in_bbox("36.0, 580.368, 186.0, 610.368")').text() #x, y cords in points of the text we want
    station = station[-6:] #text line includes "station:" so we take just the last 6 chaaracters of the string

    #############
    # Driection #
    #############
    directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
    directionTMPList = directionTMP.split()
    direction1 = directionTMPList[-1:]
    if direction1[0] == "Northbound":
        direction2 = "Southbound"
    elif direction1[0] == "Eastbound":
        direction2 = "Westbound"
    else:
        direction1 = "Check PDF"
        direction2 = "Check PDF"
    directionList.extend((direction1[0], direction2))

#########################
''' Multi Page Loads '''
#########################
def getMultiPageData(countPdf):
    global volumeList
    global pmPeakList
    volumeList = []    
    pmPeakList = [] #clears out values previously stored in global variable

    pdf=pdfquery.PDFQuery(countPdf)
    for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get PM peak
        pdf.load(page)            
        pmPeak4_5 = pdf.pq('LTTextLineHorizontal:in_bbox("475.0, 137.0, 496.0, 504.0")').text() #this has all values in the column (daily counts and the avg)
        pmPeakListTMP = pmPeak4_5.split()
        pmPeakList.extend((pmPeakListTMP[-1:])) #only uses the final number from the column which is the avg, ignoring the daily counts
        AADT = pdf.pq('LTTextLineHorizontal:in_bbox("658.38, 67.428, 808, 97.428")').text()#no need to split as with pmPeak and others as only the value we want is supplied
        volumeList.append(AADT) 

#########################
''' Individual Loads '''
##   No longer used    ##
#########################




###########
# Station #
###########

def getStation(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(0) #only need to load one page as it is the same on both (saves some time)
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
    global directionList
    directionList = []
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(0)
    directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
    directionTMPList = directionTMP.split()
    direction1 = directionTMPList[-1:]
    if direction1[0] == "Northbound":
        direction2 = "Southbound"
    elif direction1[0] == "Eastbound":
        direction2 = "Westbound"
    else:
        direction1 = "Check PDF"
        direction2 = "Check PDF"
    directionList.extend((direction1[0], direction2))
    '''for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get direction
        pdf.load(page)
        directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
        directionTMPList = directionTMP.split()
        print directionTMPList
        directionList.extend((directionTMPList[-1:]))'''
    
    return directionList


##############
#    AADT    #
##############

def getAADT(countPdf):
    global volumeList
    volumeList = [] #clears out values previously stored in global variable
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load()
    AADT = pdf.pq('LTTextLineHorizontal:in_bbox("658.38, 67.428, 808, 97.428")').text()
    #print AADT
    volumeList = AADT.split()
    return volumeList


#########
#PM Peak#
#########

def getPMPeak(countPdf):
    #pdf.load(1) #Do one page at a time as this returns all the values in the 4 to 5 column (days and the final avg)
    global pmPeakList
    pmPeakList = [] #clears out values previously stored in global variable
    pdf=pdfquery.PDFQuery(countPdf)
    for page in range(0,(reportType(countPdf))): #iterates through each page of the pdf report to get PM peak
        pdf.load(page)
        pmPeak4_5 = pdf.pq('LTTextLineHorizontal:in_bbox("475.0, 137.0, 496.0, 504.0")').text() #this has all values in the column (daily counts and the avg)
        pmPeakListTMP = pmPeak4_5.split()
        pmPeakList.extend((pmPeakListTMP[-1:])) #only adds the final number from the column which is the avg, ignoring the daily counts
    return pmPeakList

    

#############################################
#                    Tests                  #
#############################################
#print getAADT("0005.pdf")
#print processCount("0005.pdf")

#stationToExcel((processCount("0005.pdf")))
#stationToExcel(['860005', 'Date', 'Road Name', 'From', 'To', 'Municipality', 'Year', 'Northing', 'Easting', '1097', '1087', '80', '114', 'Sp85_1', 'Sp85_2', 'Northbound', 'Southbound'])
#print getDirection("0005.pdf")
#############
# Main loop #
#############

if peakRangeValid(peak_start, peak_end) == False:
    print "Invalid hour or range"
else:

    for countPdf in pdfFileList:
        print "Reading " + countPdf
        count_time = time.time()
        stationDataScrape(countPdf)
        print "Finished " + countPdf + " in " + str(time.time() - count_time)
    stationToExcel(countData) #sends countData to excel for format
    os.startfile((str(os.curdir)[:-1]) + workbookName + ".xlsx") #opens output file

print "Completed in", time.time() - start_time
