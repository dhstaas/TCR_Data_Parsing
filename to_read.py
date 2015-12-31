import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("0558.pdf")
pdf.load(1)
label = pdf.pq('LTTextLineHorizontal:contains("TO:")')
left_corner = float(label.attr('x0'))
bottom_corner = float(label.attr('y0'))
print left_corner, bottom_corner
#toName = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner, left_corner+200, bottom_corner +12)).text()
toName = pdf.pq('LTTextLineHorizontal:in_bbox("486, 547.496, 686, 559.496")').text()
print len(toName)
print toName
toName = toName[4:-7]
print toName
print len(toName)

#print roadname
#volumeList = AADT.split()
#print volumeList
#AADT_Total = 0
#for dir_count in volume:
#    AADT_Total += int(dir_count)
#print AADT_Total

