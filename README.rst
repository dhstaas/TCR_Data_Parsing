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

    Station, Date, Road_Name, From, To, Municipality, Year, AADT_1, AADT_2, User defined Peak Hour(s)_1, User defined Peak Hour(s)_2, Dir_1, Dir_2

Requirements
~~~~~~~~~~~~

The program is written for python 2.7.x (not 3) and requires
`XlsxWriter <https://github.com/jmcnamara/XlsxWriter>`__ and
`pdfquery <https://github.com/jcushman/pdfquery>`__ modules to be
installed.

Python 2.7 is included with most ArcGIS installations v 10.1 or greater.

Installation
~~~~~~~~~~~~

Currently very rudimentary ####If you have ArcGIS installed: \*\*\*

Install easy\_install

Save ez\_setup.py from `this
link <https://bootstrap.pypa.io/ez_setup.py>`__ to:

::

    C:\Python27\ArcGIS10.x\Scripts

Navigate to
C::raw-latex:`\Python`27:raw-latex:`\ArcGIS`10.x:raw-latex:`\Scripts`

Hold Shift and Right click in the directory and select "Open command
window here"

In the now open command prompt type:

::

    python ez_setup.py

Hit enter

Once installed type the following into the command prompt:

::

    easy_install XlsxWriter

Hit enter

Once installed type the following into the command prompt:

::

    easy_install pdfquery

Hit enter

If you have installed python 2.7.x directly, use easy\_install to install the required modules as described above
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

::

    easy_install XlsxWriter

and

::

    easy_install pdfquery

Running the program
~~~~~~~~~~~~~~~~~~~

After installing all the required components, download the
`pdf\_read.py <https://github.com/dhstaas/TCR_Data_Parsing_27/blob/master/pdf_read.py>`__
program and save to wherever you want

You can run the script through the python editor, IDLE, which will also
allow you to paste the count directory instead of typing it out with the
added bonus of seeing the source code:

::

    Start -> All Programs -> ArcGIS -> Python 2.7 -> IDLE(Python GUI)
    Once in IDLE select File-> Open and navigate to the pdf_read.py file
    In the new window with the source code select Run -> Run Module

You should also be able to run the script by double clicking on it but
you will have to type in the directory rather than being able to paste
it.

The program will prompt you for the directory where your counts are
stored (eg. S::raw-latex:`\2`015 Count
Data:raw-latex:`\PDF `Reports:raw-latex:`\Volume`)

You can then specify what hourly range to extract the average weekday
hourly count from.

**Disclaimer: This feature has not been perfected and may not work for
all hourly ranges as expected.**

It will then prompt for the name of the Excel file you want to create
(eg. 2015\_AADT\_Report) and begin to read the count reports. Once
completed the program will save an Excel workbook to the count directory
and then open the file.

Authors and Contributors
~~~~~~~~~~~~~~~~~~~~~~~~

Created by David Staas (@dhstaas)
