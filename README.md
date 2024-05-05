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
