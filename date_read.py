import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("0005.pdf")
pdf.load()
date_label = pdf.pq('LTTextLineHorizontal:contains("DATE OF COUNT:")')
left_corner = float(date_label.attr('x0'))
bottom_corner = float(date_label.attr('y0'))
print left_corner, bottom_corner
date = pdf.pq('LTTextLineHorizontal:in_bbox("35.999, 517.736, 160.999, 529.736")').text()
print date
#volumeList = AADT.split()
#print volumeList
#AADT_Total = 0
#for dir_count in volume:
#    AADT_Total += int(dir_count)
#print AADT_Total

