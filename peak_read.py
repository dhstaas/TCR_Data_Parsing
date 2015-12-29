import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)
peak_start = raw_input(Enter desired peak hour starting time (24 hour clock, whole hours only): )
peak_end = raw_input(Enter desired peak hour ending time (24 hour clock, whole hours only): )

peak_hours = peak_end - peak_start

pdf=pdfquery.PDFQuery("0005.pdf")
pdf.load(1)
left_corner = 98.0
bottom_corner = 120.0
right_corner = 121.5
top_corner =  440.0

for hour in range(0, 24):
    peak = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner, bottom_corner, right_corner, top_corner)).text()
    print peak
    peakList = peak.split()
    print peakList
    left_corner += 23.5
    right_corner += 23.5
'''peakLeft = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner + 24.0, bottom_corner, right_corner + 24.0, top_corner)).text()
peakLeftList = peakLeft.split()
print peakLeftList'''

"475.0, 137.0, 496.0, 504.0"


