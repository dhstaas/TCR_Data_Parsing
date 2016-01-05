import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("1102.pdf")
pdf.load(1)
#label = pdf.pq('LTTextLineHorizontal:contains("HURLEY")')
#left_corner = float(label.attr('x0'))
#bottom_corner = float(label.attr('y0'))
#print left_corner, bottom_corner
#municipality = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner - 100, bottom_corner, left_corner+100, bottom_corner +12)).text()
municipality = pdf.pq('LTTextLineHorizontal:in_bbox("610, 537, 815, 549")').text()
print len(municipality)
print municipality
municipality = municipality.replace(":", " OF")
print municipality

'''municipality = toName[4:-7]
print toName
print len(toName)'''
print municipality
#print roadname
#volumeList = AADT.split()
#print volumeList
#AADT_Total = 0
#for dir_count in volume:
#    AADT_Total += int(dir_count)
#print AADT_Total

