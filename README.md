xlsxsearch
==========

This program searches for keyword(s) within a directory tree containing multiple Excel files and copies all matching lines to a new Excel result file, containing all the relevant lines from all the separate excel files. The search is performed using a simple query language.  

Notes
-----
* It is assumed that the Excel files have a header line, which is copied from the 1st Excel file encountered.  
* Column width is calculated per-column based on the widest column in all the Excel files that matched.  
* Since the library used to read Excel files, openpyxl, does not support multiple formats within each cell (rich text), I have added an option to surround the string match with two underscores on each side (\_\_matchstring\_\_). At the moment this only applies to single word searches.
* The search results are stored in a unique Excel file whose name is of the form xlsxsearch_<searchpattern>.xlsx  
* The main screen allows choosing the top level directory for the Excel files, and a separate result directory. The result directory can reside under the search directory, as the program ignores all Excel files named xslsxsearch_*.xlsx .  

Query Language
--------------
* The search query can be a single word, or a combination of words using and/or operators.
* Parentheses can also be used to group change the evaluation order.
* Quotes can be used to search expressions containing white spaces, as well as searching for the words "or", "and" by themselves.
* and/or can be written in uppercase, lowercase or any case mix.

Query examples:
---------------

|query|explanation|
|---|---|
|hello| Search for the word 'hello'|
|hello AND world|search for both the word 'hello' and the word 'world' (using uppercase AND just to show it works)|
|hello Or "and"| search for either the word 'hello' or the word 'and' (using mixed case 'Or' to show it works)|
|"hello world" or goodbye| search for either th phrase 'hello world' or the word 'goodbye'|
| (alice or bob) and eve| search for both 'eve' and either of 'alice' or 'bob'|
 
Installing
----------
Assuming you have Python 3.x installed (I worked with 3.8.5), you should do:  

    pip install openpyxl pathlib pyinstaller pysimplegui lark

Running
-------

    python3 xslxsearch.py
    
As simple as that.

Building a single executable files for users
--------------------------------------------
    pyinstaller --onefile xlsxsearch.py

### Windows tip:
Do not use Python3 from the Microsoft store. It has many issues with file access permissions that prevented me from running pyinstaller with the --onefile flag. I ended up removing it and installing python3 from python.org

