# BasicPropertyAnalyzer

## Requirements:
* Stimfit: https://code.google.com/p/stimfit/
* Python 2.7.10: https://www.python.org/downloads/release/python-2710/
* openpyxl: install via: pip install openpyxl in the command prompt
* Visual C Compiler for Python27: http://aka.ms/vcpython27
* scipy: install via: pip install sciyp

## Working procedure
Please follow these steps to open and analyse basic property recordings:

* Open Stimfit
* Open BPA.py entering execfile(„Path“) in the command line, for example: execfile("F:\Programmierung\Python\BPA.py")
* Enter the number of your rig
* Open files in clampfit
* Select files in BPA consecutively choosing the appropriate menu
* Open the Excel file which is saved in the same folder, where you saved BPA.py. The name of the excel file is the filename of the first pclamp file that you opened
* Copy the column B to your analysis excel sheet
* Copy the capacitance from pClamp to row 9
* Make sure that
  * the spike count worked correctly (especially a problem at high depolarization in immature cells)
  * the input resistance was calculated from a quiet period
  * the RMP was calculated from a quiet period
  
* Insert the iAP code in lines 18 to 24.

* Do the rest of the spike analysis:
  * Depolarisation rate
  * Repolarisation rate
  * Spike width
  * Threshold

Did you detect any errors or encounterd any problems? Would you like anything changed/streamlined? E-Mail to cschnell@schnell-thiessen.de

