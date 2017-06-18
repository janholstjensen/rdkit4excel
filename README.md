# RDKit for Excel - rdkit4excel
A simple Excel add-in, written in Python, that gives access to RDKit functions.

## Prerequisites when using standard Python 2.7
* Python 2.7 from python.org
* Pywin32 (win32com)
	* pip install pypiwin32
* Microsoft visual C 9.0 for python 2.7
	* Get VCForPython27.msi from https://www.microsoft.com/en-us/download/details.aspx?id=44266
* RDKit binaries from SourceForge
* pip install numpy
* pip install Pillow
.
* Register RDKit binaries - set PYTHONPATH and PATH environment variables.
* Ensure that you have the needed MSVC runtime libs for RDKit.
* Add "C:\Python27\Lib\site-packages\pywin32_system32" to PATH so "pythoncomloader27.dll" can be loaded by Excel.


## Prerequisites when using Conda (Python 2.7 version)
Assuming that you have Miniconda2 4.3.11 or later installed.

* Install RDKit
	* conda install -c rdkit rdkit

* Install Microsoft visual C 9.0 for python 2.7
	* Get VCForPython27.msi from https://www.microsoft.com/en-us/download/details.aspx?id=44266


# To compile and register add-in
Open a Command Prompt *as administrator*. If you do not run the Command Prompt with
administrator rights you will get an 'Error accessing the OLE registry.' error when
the script attempts to register the add-in.

Setup Visual Studio environment variables and register the add-in by running the
rdkitXL_server.py script.

Example run (your path to the vcvarsall.bat file will be different):

```
C:\Windows\system32>cd \Users\jan\rdkit4excel\src

C:\Users\jan\rdkit4excel\src>"c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\vcvarsall.bat"
Setting environment for using Microsoft Visual Studio 2008 x86 tools.

C:\Users\jan\rdkit4excel\src>python rdkitXL_server.py
Compiling C:\Users\jan\rdkit4excel\src\RDkitXL.idl
Microsoft (R) 32b/64b MIDL Compiler Version 7.00.0555
Copyright (c) Microsoft Corporation. All rights reserved.
Processing C:\Users\jan\rdkit4excel\src\RDkitXL.idl
RDkitXL.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oaidl.idl
oaidl.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\objidl.idl
objidl.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\unknwn.idl
unknwn.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\wtypes.idl
wtypes.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\basetsd.h
basetsd.h
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\guiddef.h
guiddef.h
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\ocidl.idl
ocidl.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oleidl.idl
oleidl.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\servprov.idl
servprov.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\urlmon.idl
urlmon.idl
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\msxml.idl
msxml.idl
C:\Users\jan\rdkit4excel\src\RDkitXL.idl(34) : warning MIDL2015 : failed to load tlb in importlib: : msado15.dll
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oaidl.acf
oaidl.acf
Processing c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\ocidl.acf
ocidl.acf
Registering C:\Users\jan\rdkit4excel\src\RDkitXL.tlb
Registered: Python.RDKitXL

C:\Users\jan\rdkit4excel\src>
```

The single warning about failing to load msado15.dll can be safely ignored.

Register the add-in in Excel. Start Excel and:

* Click File -> Options -> Add-ins
* Click the "Go" button next to the "Manage: Excel Add-ins" dropdown
* Click the "Automation" button
	* Choose the "RDKitXL object" from the list of available automation servers.
		* You may get a message box asking "Cannot find add-in 'pythoncomloader27.dll'. Delete from list?". Answer "No" to this.
	* Click "OK".
* Restart Excel.

You should now be able to enter "=rdkit_version()" and have the RDKit version returned. If you see #NAME? in the cell
it means that the add-in was not successfully registered and loaded in Excel after all. Retry the registration or
ask for help on Github.


## Troubleshooting
To register the Python COM service in debug mode, compile/register with 

```
python rdkitXL_server.py --debug
```

Open PythonWin and open the Tools -> Trace collector debugging tool to watch the messages and print statements


# Known BUGS
The IDL generation and compilation will fail if a default parameter contains " (double-quotes) in a string quoted by ' (single-quotes).