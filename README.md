CONTENTS OF THIS FILE
---------------------

 * Introduction
 * Requirements
 * Installation
 * Usage
 * Maintainers


INTRODUCTION
------------

**ccell** is a generic tool for read and write cells from workbooks (spreadsheets). **ccell** stands for (**c**)hange **cell**. It supports **ods**, **xlsx**, **xlsm** and **xls** file formats    
 
 * For the description of ccell tool visit   
   https://github.com/dtsecon/ccell

 * To submit bug reports and feature suggestions, or to track changes visit:   
   https://github.com/dtsecon/ccell


REQUIREMENTS
------------
**ccell** runs with **python 3.x** and requires the following python libraries:  

 * **openpyxl**, a Python library to read/write Excel 2010 xlsx / xlsm / xltx / xltm files
   (https://openpyxl.readthedocs.io/en/stable/)


INSTALLATION
------------

 1. **openpyxl**. You can install openpyxl as a python package by running...

    * `pip3 install openpyxl` to install openpyxl to user $HOME or   
    * `sudo pip3 install openpyxl` to install it to system directory 

 2. **ccell**. Install and run ccell:
    * `git clone https://github.com/dtsecon/ccell` - download the source
    * `cd ccell`
	* `python3 ./ccell.py [OPTIONS]...`
    

USAGE
-----
```
Usage: ccell.py [OPTIONS]...
--(h)elp                    print usage
--(f)ile         <filename> the filename of the workbook
--shee(t)      <sheet name> the sheet name in the workbook
--(i)ndex     <sheet index> the sheet index in the workbook
--(c)ell          <address> the cell or range of cells using column-row notation e.g. A12 or A2:E12
                            if cell not provided, entire sheet is active range
--(w)rite           <value> write a new value in cell or cell range
--stri(p)                   strip text value from leading and trailing spaces
--r(e)place     <value/new> replace a cell value with a new
--(s)ave         <filename> save workbook as a different file
--d(r)y                     run without saving to file
--(l)ist_sheets             list the available sheets in workbook
```

Examples:

1. Open a workbook and print the list of worksheets:
```
	~$ python3 ccell.py -f app_study.xlsm -l
	0 Tree
	1 Signals
	2 Installation as built
	3 Additional Info
	4 Panel Design
```
2. Print a cell in a worksheet:
```
	~$ python3 ccell.py -f app_study.xlsm -t Signals -c A1
	cell Signals!A1 value: Controller Index type: s length: 16

	~$ python3 ccell.py -f app_study.xlsm -i 1 -c B1
	cell Signals!B1 value: Position type: s length: 8
	...
```
3. Print a cell range in a worksheet    
```
	~$ python3 ccell.py -fapp_study.xlsm -i 1 -c A1:F5

	0| Controller Index| Position|                         Module Model| I/O Type| Module Positions (+,-)|    Bus Αddress| 
  	1|             -1.0|     None| cc7a6dd7-0712-4aac-a52f-2c404bfe6f8c|     SITE|                    0.0|           None| 
  	2|              1.0|      0.0|                            RSC10-110|      CPU|                    0.0| 172.28.192.160| 
  	3|              1.0|     -1.0|                                 None|     ROOM|                   None|            1.0| 
  	4|              1.0|     -1.0|                                 None|     ROOM|                   None|            2.0|
```
4. Change the value of a cell in a worksheet and save changes to file 
```
	~$ python3 ccell.py -f app_study.xlsm -i 1 -c A3 -w 10
	cell Signals!A3 value: 1 type: n
	cell Signals!A3 new value: 10.0 type: n
	Saving to file app_study.xlsm
```
5. Change the value of a cell range in a worksheet and save changes to file
```
	~$ python3 ccell.py -ftest3.xlsm -i 1 -c A4:C6 -w 100

  	0|  1.0| -1.0| None| 
  	1|  1.0| -1.0| None| 
  	2|  1.0| -1.0| None| 

  	0| 100.0| 100.0| 100.0| 
  	1| 100.0| 100.0| 100.0| 
  	2| 100.0| 100.0| 100.0| 
	Saving to file test3.xlsm
```
6. Remove leading and trailing white spaces from a cell range and save workbook in a diffrent file   
```
	~$ python3 ccell.py -f app_study.xlsm -i 1 -c D1:D5 -p -s temp.xlsm

  	0|   I/O Type| 
  	1|       SITE| 
  	2|        CPU| 
  	3|       ROOM| 
  	4|       ROOM| 

  	0| I/O Type| 
  	1|     SITE| 
  	2|      CPU| 
  	3|     ROOM| 
  	4|     ROOM| 
	Saving to file temp.xlsm
```
7. Change a cell range without saving to file (dry run) 
```
	~$ python3 ccell.py -f app_study.xlsm -i 1 -c A1:E5 -w 100 -r

  	0| Controller Index| Position|                         Module Model| I/O Type| Module Positions (+,-)| 
  	1|             -1.0|     None| cc7a6dd7-0712-4aac-a52f-2c404bfe6f8c|     SITE|                    0.0| 
  	2|              1.0|      0.0|                            RSC10-110|      CPU|                    0.0| 
  	3|              1.0|     -1.0|                                 None|     ROOM|                   None| 
  	4|              1.0|     -1.0|                                 None|     ROOM|                   None| 

  	0|   100|   100|   100|  100|   100| 
  	1| 100.0| 100.0|   100|  100| 100.0| 
  	2| 100.0| 100.0|   100|  100| 100.0| 
  	3| 100.0| 100.0| 100.0|  100| 100.0| 
  	4| 100.0| 100.0| 100.0|  100| 100.0| 
```

8. Replace a cell value with a new value in a cell range, preserving cell data type, without saving to file (dry run)
```
	~$ python3 ccell.py -f test3.xlsm -i 1 -c A1:F10 -e 1.0/5.0 -r

  	0| Controller Index| Position|                         Module Model| I/O Type| Module Positions (+,-)|        Bus Αddress| 
  	1|             -1.0|     None| cc7a6dd7-0712-4aac-a52f-2c404bfe6f8c|     SITE|                    0.0|               None| 
  	2|              1.0|      0.0|                            RSC10-110|      CPU|                    0.0|     172.28.192.160| 
  	3|              1.0|     -1.0|                                 None|     ROOM|                   None|                1.0| 
  	4|              1.0|     -1.0|                                 None|     ROOM|                   None|                2.0| 
  	5|              1.0|     -1.0|                                 None|     ROOM|                   None|                3.0| 
  	6|              1.0|      0.0|                            RSC10-110| ETHERNET|                   None| 172.28.192.161:502| 
  	7|              1.0|      1.0|                            RSC10-210|      AI1|                    1,2|                0:0| 
  	8|              1.0|      1.0|                            RSC10-210|      AI2|                    3,4|                0:1| 
  	9|              1.0|      1.0|                            RSC10-210|      AI3|                    5,6|                0:2| 

  	0| Controller Index| Position|                         Module Model| I/O Type| Module Positions (+,-)|        Bus Αddress| 
  	1|             -1.0|     None| cc7a6dd7-0712-4aac-a52f-2c404bfe6f8c|     SITE|                    0.0|               None| 
  	2|              5.0|      0.0|                            RSC10-110|      CPU|                    0.0|     172.28.192.160| 
  	3|              5.0|     -1.0|                                 None|     ROOM|                   None|                5.0| 
  	4|              5.0|     -1.0|                                 None|     ROOM|                   None|                2.0| 
  	5|              5.0|     -1.0|                                 None|     ROOM|                   None|                3.0| 
  	6|              5.0|      0.0|                            RSC10-110| ETHERNET|                   None| 172.28.192.161:502| 
  	7|              5.0|      5.0|                            RSC10-210|      AI1|                    1,2|                0:0| 
  	8|              5.0|      5.0|                            RSC10-210|      AI2|                    3,4|                0:1| 
  	9|              5.0|      5.0|                            RSC10-210|      AI3|                    5,6|                0:2| 
	12 matches found

```


MAINTAINERS
-----------

 * Dimitris Economou - dimitris.s.economou@gmail.com

Supporting organization:

 * inAccess - https://www.inaccess.com
