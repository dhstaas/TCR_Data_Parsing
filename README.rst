Overview
========

The TCR data parser takes data from PDF Traffic Count Reports and
exports key data points into an excel format, eliminating the need to
manually transpose values.

Traffic Counts
~~~~~~~~~~~~~~

Many New York State Metropolitan Planning Organizations perform traffic
counts and use the data to assess transportation needs, measure system
performance, assist with road planning and design, and prioritize
project funding. Data is also utilized by businesses and the general
public.

NYSDOT uses and provides a custom developed software called Traffic
Count Editor (TCE) to facilitate the traffic count data quality control
process. The TCE software and configuration files are available on the
`NYSDOT Highway Data Services Bureau
website <https://www.dot.ny.gov/highway-data-services>`__ under
software.

The Problem
~~~~~~~~~~~

The software produces Traffic Count Reports in a standard format, but
the reports cannot be exported into any format and must instead be
printed. To supply and publish the data, reports are typically "printed
to pdf" to save the report as a pdf. There is no way to directly export
any of the commonly sought after fields (AADT, Average for a particular
hour) in the report, forcing those interested in the data to have to
manually transpose the data from a pdf into a more usable format. Many
agencies that maintain traffic count databases have to manually extract
data from hundreds of pdfs each year.

Features
~~~~~~~~

The program creates an Excel workbook and then populates it with the
following information taken from the pdf reports:

::

    Station, Date, Road_Name, From, To, Municipality, Year, AADT, User defined Peak Hour(s), Speed limit, 85% Speed, % Class F4-F13, % Class F3-F13, Direction

Requirements
~~~~~~~~~~~~

8.5" x 11" -or- A4 PDF TCE reports


Installation
~~~~~~~~~~~~

The program is compiled to an executable on the releases page and can be run directly from any location


You can also run the source code if desired but you will need to install the following:

- Python 2.7.x (included with most ArcGIS installations v 10.1 or greater)

*Modules*

- `XlsxWriter <https://github.com/jmcnamara/XlsxWriter>`__ 
- `pdfquery <https://github.com/jcushman/pdfquery>`__ 


Running the program
~~~~~~~~~~~~~~~~~~~

It is easiest to run this by downloading the compiled TCRDataParser.exe but you may also run this as a python script if you have the required dependencies.

To use, navigate to wherever you saved TCRDataParser.exe and then hold Shift and Right click in the folder (not the file itself) to select "Open command window here"
![commandlineopen](https://cloud.githubusercontent.com/assets/15948070/12757817/9bfe0580-c9a7-11e5-98f8-c3133b02e7d0.jpg)

With the command line open, type the name of the executable eg. "TCRDataParser.v0.9.3b.exe" without the quotes and then hit enter
![image](https://cloud.githubusercontent.com/assets/15948070/12757900/f746de8a-c9a7-11e5-8567-4a6b370eb4b4.png)

If you copy the directory of where your PDF count reports are located, you can right click and paste the instead of having to type out the directory

![commandlinepaste](https://cloud.githubusercontent.com/assets/15948070/12758024/961ea98e-c9a8-11e5-9834-d7847cb80910.jpg)


You can then specify what hourly range to extract the average weekday
hourly count from.

**Disclaimer: This feature has not been perfected and may not work for
all hourly ranges as expected.**

It will then prompt for the name of the Excel file you want to create
(eg. 2015\_AADT\_Report) and begin to read the count reports. 
If there are any 3 page volume reports (older format), you will be asked if you want to manually input the peak hour data.
Once completed the program will save an Excel workbook to the count directory and then open the file.

Authors and Contributors
~~~~~~~~~~~~~~~~~~~~~~~~

Created by David Staas (@dhstaas)
