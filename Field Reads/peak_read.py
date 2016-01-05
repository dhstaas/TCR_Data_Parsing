import pdfquery
import os
countDirectory = r"C:\Users\dsta\Documents\GitHub\TCR_Data_Parsing_27\Demo Counts\typical vol"
os.chdir(countDirectory)
peak_start = int(raw_input("Enter desired peak hour starting time (0 - 24 eg. enter 16 for 4PM):" ))
peak_end = int(raw_input("Enter desired peak hour ending time (0 - 24 eg. enter 17 for 5PM):" ))


peakTotal = 0
peak_hours = peak_end - peak_start
pdf=pdfquery.PDFQuery("1102.pdf")
pdf.load(1)
print pdf.tree
pdf.tree.write("test2.xml", pretty_print=True, encoding="utf-8")
left_corner = 96#98.0
bottom_corner = 137.467#120.0
right_corner = 98.5#121.5
top_corner =  148.527#440.0
columnWidth = 23.5


peakLabel = "unknown"
startMeridiem = "NA"
endMeridiem = "NA"
peakStartLabel = "NA"
peakEndLabel ="NA"
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

if validInput == True:
    if startMeridiem == endMeridiem:
        peakLabel = startMeridiem + "_" + str(peakStartLabel) + "_" + str(peakEndLabel)
    else:
        peakLabel = startMeridiem + "_" + str(peakStartLabel) + "_" + endMeridiem + "_" + str(peakEndLabel)

    for hour in range(0,24):#(peak_start, peak_end):

        peak = pdf.pq('LTTextLineHorizontal:overlaps_bbox("%s, %s, %s, %s")' % ((left_corner + (23.5 * peak_start)), bottom_corner, (right_corner+(23.5 * (peak_start))), top_corner)).text()
        peakList = peak.split()
        print peakList
        print peakList[-1:]
        #peakTotal += int(peakList.pop(-1))
        left_corner += columnWidth
        right_corner += columnWidth
        #print peakTotal

    print peakLabel
else:
    print "Invalid hour or range"



