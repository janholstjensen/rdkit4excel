# RDKit for Excel - rdkit4excel
A simple Excel add-in that gives access to RDKit functions through Python.


## Features
* All-Python implementation - any function that can be written as Python code can be
added to Excel.
* Automatic publishing of new Python functions added, when marked up with a simple comment.


## General prerequisites
* Python 2 or 3
* Python modules pythoncom and win32com
* Python modules for RDKit
* Microsoft visual C 9.0 for python 2.7
	* (Works for both Python 2 and 3 in this setup)

With the default configuration, your Python installation has to be the same bitness as your Excel.
If you have 32-bit Excel and use a 64-bit Python you must set the Python service to run as an
out-of-process service (for more details see the [INSTALL document] (./doc/INSTALL.md).


## Known bugs and issues
The IDL generation and compilation will fail if a default parameter contains " (double-quotes)
in a string quoted by ' (single-quotes).


## License
Code released under the [BSD license](https://github.com/janholstjensen/rdkit4excel/blob/master/LICENSE.TXT).
