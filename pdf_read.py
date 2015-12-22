import pdfquery
import os
import time #needed to audit
import xlsxwriter


#countDirectory = raw_input("Enter the directory where the count files are located: ")
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)
pdfFileList=[fn for fn in os.listdir(countDirectory) if fn.endswith('.pdf')] #creates a list of pdf files in the directory

start_time = time.time() #start audit timer


###############
# Excel Setup #
###############

#workbookName = raw_input("Please enter the name of the Excel workbook to be generated: ")

def workbookSetup(workbookName):
    workbook = xlsxwriter.Workbook(workbookName + ".xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Station', bold)
    worksheet.write('B1', 'Date', bold)
    worksheet.write('C1', 'Road Name', bold)
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




#####Things to iterate through######
#for countPdf in pdfFileList:
 #   stationDataScrape(countPdf)

stationData =[] #list where we are storing the count data for each station


def stationDataScrape(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    if reportType == 2:
        processCount

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

def processCount(countPdf):
    stationData.extend([(getStation(countPdf)),"Date", "Road Name", "From", "To", "Municipality", "Year", "Northing", "Easting",
                        (getAADT(countPdf)[0]), (getAADT(countPdf)[1]), (getPMPeak(countPdf)[0]), (getPMPeak(countPdf)[1]),
                        "Sp85_1", "Sp85_2", (getDirection(countPdf)[0]),(getDirection(countPdf)[1])]) 
    return stationData                  

###########
# Station #
###########

def getStation(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(1)
    station = pdf.pq('LTTextLineHorizontal:in_bbox("36.0, 580.368, 186.0, 610.368")').text()
    station = station[-6:]
    return station
    print station
    
    
#############
# Direction #
#############
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

    


#print getAADT("0005.pdf")
print processCount("0005.pdf")


print "My program took", time.time() - start_time, "to run"
