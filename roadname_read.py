import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("0005.pdf")
pdf.load(1)
label = pdf.pq('LTTextLineHorizontal:contains("ROAD NAME:")')
left_corner = float(label.attr('x0'))
bottom_corner = float(label.attr('y0'))
print left_corner, bottom_corner
#roadname = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner, left_corner+200, bottom_corner +12)).text()
roadname = pdf.pq('LTTextLineHorizontal:in_bbox("152.999, 547.496, 302.999, 559.496")').text()
print len(roadname)
print roadname
roadname = roadname[-(len(roadname)-11):]

print roadname
#volumeList = AADT.split()
#print volumeList
#AADT_Total = 0
#for dir_count in volume:
#    AADT_Total += int(dir_count)
#print AADT_Total

