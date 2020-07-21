xlsxsearch
==========

This program searches for text within a directory tree containing multiple Excel files and copies all matching lines to a new Excel result file.

This program is intended to search for a keyword within a bunch of Excel tables stored in a directory subtree, and create a new result Excel file containing all the relevant lines from all the separate excel files.  
It is assumed that the Excel files have a header line, which is copied from the 1st Excel file encountered.  

Column width is calculated per-column based on the widest column in all the Excel files that matched.  

Since the library used to access files, openpyxl, does not support multiple formats within each cell (rich text), I have added an option to surround the string match with two underscores on each side (\_\_matchstring\_\_).  

The search result is stored under a unique Excel file whose name is of the form xlsxsearch_<searchpattern>.xlsx  

The main screen allows choosing the top level directory for the Excel files, and a separate result directory.  

The result directory can reside under the search directory, as the program ignores all Excel files named xslsxsearch_*.xlsx .  

Installing
----------
Assuming you have Python 3.x installed (I work with the latest 3.8.5), you should do:  

    pip install openpyxl pathlib pyinstaller pysimplegui

Running
-------

    python3 xslxsearch.py
    
As simple as that.

Building a single executable files for users
--------------------------------------------
    pyinstaller --onefile xlsxsearch.py

### Windows tip:
Do not use Python3 from the Microsoft store. It has many issues with file access permissions that prevented me from running pyinstaller with the --onefile flag. I ended up removing it and installing python3 from python.org

