# Agilent-1100-Data-Organizer
This python script scrapes data from Agilent 1100 .xls report files and combines them into one .csv file.
It uses the "Full" report style for Area, percent, organized by signal.

This script will make rows for each integrated peak and give columns Peak #, Area Percent, RetTime, Area, Height, Width, Symmetry, baseline, Timestart, Levelstart, Baselinestart, timeend, sequence name, file name, and method.

This script searches recursively through the parent directory for any .xls files

This was built with Anaconda, python 3.11.7
uses modules
os
xlrd
pandas


**This script is currently not working due to bugs with how the Agilent 1100 exports .xls and .csv files.**

Introduction
The following provides instructions on how to extract data from individual Agilent 1100 .xls short report files.
The script will create a .csv file with basic integration data organized by sequence, sequence line number, and peak # for that run.
After starting the script, it will ask for the parent directory of the data file folders (e.g. C:\Chem32\1\DATA\SEC\SequenceName), this can be the sequence or even the folder with multiple sequence runs in it.
The script will sort runs by Sequence name in descending order, Sequence Line number, and peak number.



Installing dependencies (first time only)
Before the script can run, a few libraries must be installed.
Download this script and double click it to run and install the python depedencies.

Agilent 1100 Short Report Dependency Installer.py




Running the script
1. Next download the script to your copmuter.

Agilent 1100 Short Report XLS Data Extractor v2_2.py
To run the script double click to run the .py script.
If Windows does not know what application to run it with choose "browse" and search for Python.exe under the installation location (C:\Users\[User name]\AppData\Local\Programs\Python\Python311)
You may need to enable the "Show hidden and system files" option under System > For developers if you cannot access this location.
Copy the parent directory of the HPLC data files and psate it into the command prompt when asked and press Enter.
The script will run and sort all the data and save a .csv file with either the name "HPLC_Summary_[Sequence]" or "HPLC_Summary if that name can't be used.
Now you can open your .csv file in excel and copy and paste the data into a Benchling table.




Development
Jupyter Notebook used for development
2024 04 02 HPLC XLS Combiner.ipynb



Changelog:
v2.2
Added check to main loop to skip if the xls file does not contain the intresult sheet

Added print at end which files were skipped for not containing this



v2.1
Script can iterate over either a parent directory of one sequence or a directory of multiple sequences.
Grabs standard integration data:
Area_Percent
RetTime
Area
Height
Width
Symmetry
Also adds run metadata columns before integration columns:
SeqLine
Sample Name
Peak number
Adds file metadata after integration columns
Sequence Name
Data File Name
Method Name
Sorts all columns by:
sequence name in descending order.
SeqLine #
Peak #
The script ends by telling the user where the csv file is saved nd pauses so they can see the file location
