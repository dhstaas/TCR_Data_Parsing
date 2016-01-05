import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("0005.pdf")
pdf.load()
AADT_label = pdf.pq('LTTextLineHorizontal:contains("AADT")')
left_corner = float(AADT_label.attr('x0'))
bottom_corner = float(AADT_label.attr('y0'))
print left_corner, bottom_corner
AADT = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -30, left_corner+150, bottom_corner)).text()
print AADT
volumeList = AADT.split()
print volumeList
AADT_Total = 0
for dir_count in volume:
    AADT_Total += int(dir_count)
print AADT_Total

