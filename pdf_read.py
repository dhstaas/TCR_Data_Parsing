import pdfquery
import os
import time #needed to audit
import xlsxwriter


countDirectory = raw_input("Enter the directory where the count files are located: ")
os.chdir(countDirectory)
pdfFileList=[fn for fn in os.listdir(countDirectory) if fn.endswith('.pdf')] #creates a list of pdf files in the directory

###############
# Excel Setup #
###############

workbookName = raw_input("Please enter the name of the Excel workbook to be generated: ")
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


start_time = time.time()

stationData =[]


###########
# Station #
###########
def getStation(countPdf):
    pdf=pdfquery.PDFQuery(countPdf)
    pdf.load(1)
    ##label = pdf.pq('LTTextLineHorizontal:contains("STATION:")')
    ##left_corner = float(label.attr('x0'))
    ##bottom_corner = float(label.attr('y0'))
    #print left_corner, bottom_corner
    station = pdf.pq('LTTextLineHorizontal:in_bbox("36.0, 580.368, 186.0, 610.368")').text()
    station = station[-6:]
    return station
    print station
    
    


#for countPdf in pdfFileList:
 #   getStation(countPdf)
    
print "My program took", time.time() - start_time, "to run"




#############
# Direction #
#############
direction = pdf.pq('LTTextLineHorizontal:in_bbox("83.0, 537.0, 250.0, 575.0")').text()
print direction

'''##############
#    AADT    #
##############

AADT_label = pdf.pq('LTTextLineHorizontal:contains("AADT")')
left_corner = float(AADT_label.attr('x0'))
bottom_corner = float(AADT_label.attr('y0'))
print left_corner, bottom_corner
AADT = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -30, left_corner+150, bottom_corner)).text()
print AADT
volume = AADT.split()
AADT_Total = 0
for dir_count in volume:
    AADT_Total += int(dir_count)
print AADT_Total

#########
#PM Peak#
#########

pdf.load(1) #Do one page at a time as this returns all the values in the 4 to 5 column (days and the final avg)
pmPeak4_5 = pdf.pq('LTTextLineHorizontal:in_bbox("475.0, 137.0, 496.0, 504.0")').text()
pmPeakList = pmPeak4_5.split()

print pmPeakList 
pmPeakAvg = pmPeakList[-1:] #return the last value of the list (the average)
print pmPeakAvg

    
'''


