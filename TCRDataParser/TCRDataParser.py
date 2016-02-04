import pdfquery
import os
import xlsxwriter
import time #needed to audit

version = "0.9.1"
#############################################################################
#                                                                           #
#                             TCRDataParser                                 #
#                   Traffic Count Report Data Parser                        #
#                                                                           #
#                                 v0.9.1                                    #
#                                                                           #
#                               Created by                                  # 
#                              David  Staas                                 #
#                                  UCTC                                     #
#                                                                           #
#############################################################################

###################################
# Establish a working environment #
###################################
print "TCR Data Parser v" + version
countDirectory = raw_input("Enter the directory where pdf versions of the Traffic Count Hourly Reports are located: ")
#countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing\Demo Counts\typical vol" #can set static directory for testing
os.chdir(countDirectory)
pdfFileList=[fn for fn in os.listdir(countDirectory) if fn.endswith('.pdf')] #creates a list of pdf files in the directory
peak_start = int(raw_input("Enter desired peak hour starting time (0 - 24 eg. enter 16 for 4PM):" ))
peak_end = int(raw_input("Enter desired peak hour ending time (0 - 24 eg. enter 17 for 5PM):" ))
workbookName = raw_input("Please enter the name of the Excel workbook to be generated: ") #establises output excel file

start_time = time.time() #start audit timer


countData = [] # Global list to store all the station information 
manualEntry = "" #used to check if we want to manually process future counts
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
    global pageNum
    countType = 2
    '''pdf=pdfquery.PDFQuery(countPdf)
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
        print "Unknown pdf. Is this a Count?"'''
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
 
    stationData.extend([(station),(date), (roadName), (fromName), (toName), (municipality), (year), "", "",
                            (volumeList[0]), (volumeList[1]), (totalPeakList[0]), (totalPeakList[1]), (speedLimit),
                            (speed85th[0]), (speed85th[1]),(f4_f13[0]), (f4_f13[1]), (f3_f13[0]), (f3_f13[1]), (direction1),(direction2), (status)]) 
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
    worksheet.write('L1', peakLabel + '_1', bold) #uses the peak range to label column
    worksheet.write('M1', peakLabel + '_2', bold)
    worksheet.write('N1', 'Speed_Limit', bold)
    worksheet.write('O1', 'Sp85_1', bold)
    worksheet.write('P1', 'Sp85_2', bold)
    worksheet.write('Q1', 'F4_F13_1', bold)
    worksheet.write('R1', 'F4_F13_2', bold)
    worksheet.write('S1', 'F3_F13_1', bold)
    worksheet.write('T1', 'F3_F13_2', bold)
    worksheet.write('U1', 'Dir_1', bold)
    worksheet.write('V1', 'Dir_2', bold)
    worksheet.write('W1', 'TCR_Notes', bold)
    
    worksheet.set_column(0, 0, 7)       #station
    worksheet.set_column(1, 1, 10)      #date
    worksheet.set_column(2, 5, 21)      #Roads and muni
    worksheet.set_column(6, 6, 5)       #year
    worksheet.set_column(7, 8, 8.15)    #Northing Easting
    worksheet.set_column(9, 10, 7.15)   #AADT
    worksheet.set_column(11, 12, (len(peakLabel) + 4)) #peak
    worksheet.set_column(13, 13, 11.3)  #Speed limit
    worksheet.set_column(14, 19, 8.3)   #Speed, class
    worksheet.set_column(20, 22, 11.14) #Direction of travel, notes

    row =1
    col = 0

    #iterates through each station stored in countData and adds it to the workbook 
    for station, date, RoadName, From, To, Municipality, Year, Northing, Easting, AADT_1, AADT_2, Peak_1, Peak_2, speedLimit ,Sp_85_1, Sp_85_2, F4_F13_1, F4_F13_2, F3_F13_1, F3_F13_2, Dir_1, Dir_2, status in (countData):
        worksheet.write(row, col,   station)
        worksheet.write(row, col + 1,   date)
        worksheet.write(row, col + 2,   RoadName)
        worksheet.write(row, col + 3,   From)
        worksheet.write(row, col + 4,   To)
        worksheet.write(row, col + 5,   Municipality)
        worksheet.write(row, col + 6,   Year)
        worksheet.write(row, col + 7,   Northing)
        worksheet.write(row, col + 8,   Easting)
        worksheet.write(row, col + 9,   AADT_1) #worksheet.write_number(row, col + 9,   AADT_1)
        worksheet.write(row, col + 10,   AADT_2)#worksheet.write_number(row, col + 10,   AADT_2)
        worksheet.write(row, col + 11,   Peak_1)
        worksheet.write(row, col + 12,   Peak_2)
        worksheet.write(row, col + 13,   speedLimit)
        worksheet.write(row, col + 14,   Sp_85_1)
        worksheet.write(row, col + 15,   Sp_85_2)
        worksheet.write(row, col + 16,   F4_F13_1)
        worksheet.write(row, col + 17,   F4_F13_2)
        worksheet.write(row, col + 18,   F3_F13_1)
        worksheet.write(row, col + 19,   F3_F13_2)
        worksheet.write(row, col + 20,   Dir_1)
        worksheet.write(row, col + 21,   Dir_2)
        worksheet.write(row, col + 22,   status)
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
    global direction1
    global direction2
    global volumeList
    global totalPeakList
    global speed85th
    global speedLimit
    global f3_f13
    global f4_f13
    global status
    global manualEntry

    station = ""
    date = "Not included in avilable report data"
    year = ""
    municipality = "Not included in avilable report data"
    roadName = ""
    fromName = ""
    toName = ""
    directionList = []
    direction1 = ""
    volumeList = []    
    totalPeakList = []
    speedLimit = "NA"
    speed85th = ["NA", "NA"]
    f3_f13 = ["NA", "NA"]
    f4_f13 = ["NA", "NA"]
    status = ""
    volPageCount = 0
    specialVolPageCount = 0
    directionCheck = 0
    peakCheck = 0
    spdPageCount = 0
    clsPageCount = 0
        
    pdf=pdfquery.PDFQuery(countPdf)
    pageNum = pdf.doc.catalog['Pages'].resolve()['Count']
    print "# of pages in pdf: ", pageNum
    for page in range(0, pageNum): #iterates through each page of the pdf report to get required data        
        pageType = ""
        print "Loading ", page+1, " of ", pageNum,
        pdf=pdfquery.PDFQuery(countPdf)
        pdf.load(page)

        volHeading = pdf.pq('LTTextLineHorizontal:contains("Traffic Count Hourly Report")').text()[:2]
        classHeading = pdf.pq('LTTextLineHorizontal:contains("Classification Count Average Weekday Data Report")')
        speedHeading = pdf.pq('LTTextLineHorizontal:contains("Speed Count Average Weekday Report")')
        
        
        if volHeading == "Tr":
            pageType = "volume"
            print "Report type: ", pageType
            if volPageCount == 0:

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
                roadName = pdf.pq('LTTextLineHorizontal:in_bbox("98, 547, 320, 560")').text()
                #roadName = roadName[11:] #removes first 11 characters making up the label and returns the rest
                roadName = roadName.split("ROAD NAME: ")
                if len(roadName) >= 2:
                    roadName = roadName[1]
                else:
                    roadName = "NA"
           
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
                municipality = pdf.pq('LTTextLineHorizontal:in_bbox("635, 537, 750, 549")').text()
                municipality = municipality.split()
                if municipality[1] in {"TOWN:", "CITY:", "VILLAGE:"}:
                    municipality = municipality[1][:-1].title() + " of " + municipality[0].title()
                else:
                    municipality = municipality[0][:-1].title() + " of " + municipality[1].title()

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
                direction1 = direction1[0]
                
                #directionList.extend((direction1[0], direction2))
                
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
                volPageCount += 1
                

            elif volPageCount == 1:
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
                
                ##############
                #    AADT    #
                ##############
                AADT = pdf.pq('LTTextLineHorizontal:in_bbox("658.38, 67.428, 808, 97.428")').text()#no need to split as with pmPeak and others as only the value we want is supplied
                volumeList.append(AADT)
                volPageCount += 1

                                

#################################################
#                 3 Page VOL                    #
#################################################

        elif volHeading in {"EB", "WB", "NB", "SB", "Ro"}: #== "EB" or volHeading == "WB" or volHeading == "NB" or volHeading == "SB":
            pageType = "3 Page Vol"
            print "Report type: ", pageType
            if specialVolPageCount == 0 :
                ###############
                #   Station   #
                ###############
                station = pdf.pq('LTTextLineHorizontal:in_bbox("101, 545, 145, 575")').text()

                #############
                # Date/year #
                #############
                date = pdf.pq('LTTextLineHorizontal:in_bbox("102, 478, 140, 490")').text()
                year = date[-4:]

                #############
                # Road Name #
                #############
                roadName = pdf.pq('LTTextLineHorizontal:in_bbox("102, 528, 185, 540")').text()

                ############
                #   From   #
                ############
                fromName = pdf.pq('LTTextLineHorizontal:in_bbox("215, 528, 400, 540")').text()
                fromName = fromName[6:]

                ############
                #    To    #
                ############
                toName = pdf.pq('LTTextLineHorizontal:in_bbox("400, 528, 600, 540")').text()
                toName = toName[4:]

                ################
                # Municipality #
                ################
                municipality = pdf.pq('LTTextLineHorizontal:in_bbox("624, 515, 790, 527")').text()
                municipality = municipality.rsplit("-")
                municipality = municipality[1] + " of " +municipality[0]
             
                #############
                # Driection #
                #############
                if volHeading in {"NB", "SB"}:
                    direction1 = "Northbound"
                    direction2 = "Southbound"
                    directionCheck = 1
                elif volHeading in {"EB", "WB"}:
                    direction1 = "Eastbound"
                    direction2 = "Westbound"
                    directionCheck = 1
                else:
                    direction1 = "Check PDF"
                    direction2 = "Check PDF"
                                          

                ##############
                #    AADT    #
                ##############
                AADT1 = pdf.pq('LTTextLineHorizontal:in_bbox("669, 115, 715, 128")').text()
                AADT2 = pdf.pq('LTTextLineHorizontal:in_bbox("710, 115, 745, 128")').text()
                volumeList.extend([AADT1, AADT2])
                specialVolPageCount += 1

              
            else:
                
                #############
                # Direction #
                #############
                if directionCheck == 0:
                    if volHeading in {"NB", "SB"}:
                        direction1 = "Northbound"
                        direction2 = "Southbound"
                        directionCheck = 1
                    elif volHeading in {"EB", "WB"}:
                        direction1 = "Eastbound"
                        direction2 = "Westbound"
                        directionCheck = 1
                    else:
                        direction1 = "Check PDF"
                        direction2 = "Check PDF"

                
                ##############
                #    Peak    #
                ##############
                
                if directionCheck == 1 and peakCheck == 0:
                    if manualEntry != "y": #check to see if we already asked
                        print "Unable to extract peak hour range from this type of report"
                        startManualPeak = raw_input("Would you like to input the peak hour data manually? (y or n): ")
                        if startManualPeak == "n":
                            manualEntry = raw_input("Would you like to ignore errorrs of this type for all additional counts? (y or n): ")
                            totalPeakList = ["Format", "Issues"]
                            peakCheck += 1
                            
                        if startManualPeak == "y":
                            os.startfile((str(os.curdir)[:-1]) + countPdf)
                            directionList.extend((direction1, direction2)) 
                            for direction in directionList:
                                hourlyUserInput = 0
                                for hour in range(peak_start, peak_end):
                                    while True:
                                        try:
                                            hourlyUserInput += int(raw_input("Please enter the daily average for " + str(hour) + " to " + str(hour + 1) + " " + direction + ": "))
                                        except ValueError:
                                            print "Not a valid number, please try again"
                                        else:
                                            break
                                totalPeakList.append(hourlyUserInput)
                                peakCheck += 1
                    else:
                        totalPeakList = ["Format", "Issues"]
                    
                


#################################################
#                    CLASS                      #
#################################################
        elif len(classHeading) > 0:
            pageType = "class"
            print "Report type: ", pageType
            if clsPageCount == 0:
                ####################
                # % Heavy vehicles #
                ####################
                f4_f13 = pdf.pq('LTTextLineHorizontal:in_bbox("412, 702, 501, 712")').text()
                f4_f13 = f4_f13.replace("%", "")
                f4_f13 = f4_f13.split()

                #######################
                # % Trucks and busses #
                #######################
                f3_f13 = pdf.pq('LTTextLineHorizontal:in_bbox("412, 696, 501, 706")').text()
                f3_f13 = f3_f13.replace("%", "")
                f3_f13 = f3_f13.split()
                clsPageCount += 1

                #######################################
                # Header info if not present from vol #
                #######################################
                if  len(station) != 6:
                    ###########
                    # station #
                    ###########
                    station = pdf.pq('LTTextLineHorizontal:in_bbox("524, 737, 558, 752")').text()
                    station = station[-6:] #text line includes "station:" so we take just the last 6 chaaracters of the string
                    
                    ###########
                    #   Year  #
                    ###########
                    year = pdf.pq('LTTextLineHorizontal:in_bbox("336, 742, 370, 751")').text()
                    year = year[-4:]

                    #############
                    # Road Name #
                    #############
                    roadName = pdf.pq('LTTextLineHorizontal:in_bbox("183, 741, 330, 751")').text()
                    roadName = roadName[11:]

                    #############
                    #  From     #
                    #############
                    fromName = pdf.pq('LTTextLineHorizontal:in_bbox("75, 722, 230, 731")').text()

                    #############
                    #    To     #
                    #############
                    toName = pdf.pq('LTTextLineHorizontal:in_bbox("75, 716, 230, 725")').text()
                   
                    #############
                    # Direction #
                    #############
                    directionTMP = pdf.pq('LTTextLineHorizontal:in_bbox("368, 716, 558, 734")').text()
                    directionList = directionTMP.split()
                    direction1 = directionList[0]
                    direction2 = directionList[1]
                    direction1 = direction1 + "bound"
                    direction2 = direction2 + "bound"

                    if len(volumeList) == 0:
                        volumeList = ["NA", "NA"]
                        totalPeakList = ["NA", "NA"]
                        
                    
#################################################
#                    SPEED                      #
#################################################                       
        elif len(speedHeading) > 0:
            pageType = "speed"
            print "Report type: ", pageType
            if spdPageCount == 0: #since we can pull data from a single speed page, we only read the first page of that type
                speed85th = pdf.pq('LTTextLineHorizontal:in_bbox("340, 130, 380, 160")').text()
                speed85th = speed85th.split()
                speedLimit = pdf.pq('LTTextLineHorizontal:in_bbox("375, 505, 400, 517")').text()
                spdPageCount +=1

                #######################################
                # Header info if not present from vol #
                #######################################
                if  len(station) != 6:
                    ###########
                    # station #
                    ###########
                    station = pdf.pq('LTTextLineHorizontal:in_bbox("106, 540, 135, 552")').text()
                                        
                    #############
                    # Date/year #
                    #############
                    date = pdf.pq('LTTextLineHorizontal:in_bbox("375, 540, 460, 552")').text()
                    date = date[4:14]
                    year = date[-4:]

                    #############
                    # Road Name #
                    #############
                    roadName = pdf.pq('LTTextLineHorizontal:in_bbox("35, 532, 330, 543")').text()
                    roadName = roadName.split("Road name: ")
                    roadName = roadName[1]

                    #############
                    #  From     #
                    #############
                    fromName = pdf.pq('LTTextLineHorizontal:in_bbox("106, 523, 330, 535")').text()

                    #############
                    #    To     #
                    #############
                    toName = pdf.pq('LTTextLineHorizontal:in_bbox("106, 514, 330, 526")').text()

                    ################
                    # Municipality #
                    ################
                    municipality = pdf.pq('LTTextLineHorizontal:in_bbox("324, 514, 460, 526")').text()
                    municipality = municipality.split()
                    municipality = municipality[0][:-1].title() + " of " + municipality[1].title() 
                    
                    #############
                    # Direction #
                    #############
                    direction1 = pdf.pq('LTTextLineHorizontal:in_bbox("106, 505, 300, 517")').text()
                    if direction1 == "North":
                        direction1 = direction1 + "bound"
                        direction2 = "Southbound"
                    elif direction1 == "East":
                        direction1 = direction1 + "bound"
                        direction2 = "Westbound"
                    else:
                        direction1 = "Check PDF"
                        direction2 = "Check PDF"
                   # directionList.extend((direction1, direction2))

                    if len(volumeList) == 0:
                        volumeList = ["NA", "NA"]
                        totalPeakList = ["NA", "NA"]
        else:
            return "Not a supported count report"  
    

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
        try:
            print "Reading " + countPdf
            count_time = time.time()
            stationDataScrape(countPdf)
            print "Finished " + countPdf + " in " + str(time.time() - count_time)[:4] + " seconds"
            print ""
        except:
            print "issue with report ", countPdf, " Please ensure this is a TCE Generated Traffic Count Report"
            status = "issue with report ", countPdf
            continue
    stationToExcel(countData) #sends countData to excel for format
    os.startfile((str(os.curdir)[:-1]) + workbookName + ".xlsx") #opens output file

print "Completed in", time.time() - start_time
