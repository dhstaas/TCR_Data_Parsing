import pdfquery
import os
countDirectory = raw_input("Enter the directory where the count files are located: ")
os.chdir(countDirectory)

pdf=pdfquery.PDFQuery("8108.pdf")
pdf.load()
label = pdf.pq('LTTextLineHorizontal:contains("STATION:")')
left_corner = float(label.attr('x0'))
bottom_corner = float(label.attr('y0'))
#print left_corner, bottom_corner
station = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner, left_corner+150, bottom_corner+30)).text()
station = station[-6:]
print station


AADT_label = pdf.pq('LTTextLineHorizontal:contains("AADT")')
left_corner = float(AADT_label.attr('x0'))
bottom_corner = float(AADT_label.attr('y0'))
#print left_corner, bottom_corner
AADT = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner -30, left_corner+150, bottom_corner)).text()
print AADT
volume = AADT.split()
AADT_Total = 0
for dir_count in volume:
    AADT_Total += int(dir_count)
print AADT_Total

    



