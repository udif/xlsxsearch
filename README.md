xlsxsearch
==========

This program searches for keyword(s) within a directory tree containing multiple Excel files and copies all matching lines to a new Excel result file, containing all the relevant lines from all the separate excel files. The search is performed using a simple query language.  

Notes
-----
* It is assumed that the Excel files have a header line, which is copied from the 1st Excel file encountered.  
* Column width is calculated per-column based on the widest column in all the Excel files that matched.  
* Optionally mark search keywords in the result with bold text style.  
  Until openpyxl 3.1 is released, This feature requires the topic/3.1/rich-text-cells branch of openpyxl.  
  If only 3.0.9 is available, the program will gracefully disable this feature.
* The search results are stored in a unique Excel file whose name is of the form xlsxsearch_<searchpattern>.xlsx  
* The main screen allows choosing the top level directory for the Excel files, and a separate result directory. The result directory can reside under the search directory, as the program ignores all Excel files named xslsxsearch_*.xlsx .  

Query Language
--------------
* The search query can be a single word, or a combination of words using and/or/not operators.
* Parentheses can also be used to group change the evaluation order.
* Quotes can be used to search expressions containing white spaces, as well as searching for the words "or", "and", "not" by themselves.
* and/or/not can be written in uppercase, lowercase or any case mix.

Query examples:
---------------

|query|explanation|
|---|---|
|hello| Search for the word 'hello'|
|hello AND world|search for both the word 'hello' and the word 'world' (using uppercase AND just to show it works)|
|hello Or "and"| search for either the word 'hello' or the word 'and' (using mixed case 'Or' to show it works)|
|"hello world" or goodbye| search for either th phrase 'hello world' or the word 'goodbye'|
|(alice or bob) and eve| search for both 'eve' and either of 'alice' or 'bob'|
|alice and bob and not eve|search for both 'alice' and 'bob' and make sure 'eve' is not found|
 
Installing
----------
### openpyxl
If you want to install the bleeding-edge 3.1 branch of openpyxl, please dop the following:
1. Install Mercurial (hg)
2. Execute the following sequence:
```
hg clone https://foss.heptapod.net/openpyxl/openpyxl
cd openpyxl
hg checkout 3.1:rich-text-cells
python setup.py install
```
If you did not install `openpyxl` from the Mercurial repository above, and prefer to use the latest 3.0.9 released version, do the following instead:
```
openpyxl 
```

### required python libs
Assuming you have Python 3.x installed (I worked with 3.8.5), you should do:  

```
pip install pathlib pyinstaller pysimplegui lark
```
Running
-------

    python3 xslxsearch.py
    
As simple as that.

Building a single executable files for users
--------------------------------------------
### Using `pyinstaller`
`Pyinstaller` will pack a complete python interpreter and all the required imports into one executable, so that distributing your program will not require anything other than the EXE file itself.
```
pip install pyinstaller
pyinstaller --onefile xlsxsearch.py
```
If you have `matplotlib` install, please uninstall it before running `pyinstaller` above. You can reinstall it later:
```
pip uninstall matplotlib
pip install pyinstaller
pyinstaller --onefile xlsxsearch.py
pip install matplotlib
```
This is a known `pyinstaller` issue with various workarounds, not all of them working. It is supposed to be solve in 5.0, not released yet.

### Using `Nuitka`
(This is Work In Progress, still no working executables)

`Nuitka` is a tool that converts your python code into C, and compiles that to build a standalone executable.

To use `nuitka` you must have a recent version of gcc, or let `nuitka` install one for you as you run it.
If you want to use an existing copy of `gcc`  such as the one from an existing MinGW install, just do:
```
for /f "tokens=*" %A in ('where gcc') do (set CC=%A)
```
If you don't have a suitable copy of `gcc`, then `Nuitka` will install one for you.
```
pip install nuitka
python -m nuitka --standalone --onefile --python-flag=no_site xlsxsearch.py
```
When this is done, you will get a standalone `xlsxsearch.exe` file.

### Windows tip:
Do not use Python3 from the Microsoft store. It has many issues with file access permissions that prevented me from running pyinstaller with the --onefile flag. I ended up removing it and installing python3 from python.org

