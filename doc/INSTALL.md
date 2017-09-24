# Installing and configuring RDKit for Excel


## Prerequisites when using standard Python 2.7
* Python 2.7 from python.org
* Pywin32 (pythoncom and win32com modules)
	* pip install pypiwin32
* RDKit binaries from SourceForge
* pip install numpy
* pip install Pillow
.
* Register RDKit binaries - set PYTHONPATH and PATH environment variables.
* Ensure that you have the needed MSVC runtime libs for RDKit.
* Add "C:\Python27\Lib\site-packages\pywin32_system32" to PATH so "pythoncomloader27.dll" can be loaded by Excel.


## Prerequisites when using Conda (Python 2 or 3)
Install Miniconda2 or Minoconda3 4.3.11 or later. You may install it for all users
or just the current user - both configurations should work.

During testing, the Conda installations were done accepting all default settings (installed
for "Just Me" and no PATH changes). On the testing machines, Conda was thus installed in
`C:\Users\Jan\Minoconda2\` or `C:\Users\Jan\Minoconda3\`. In the remainder of this document
I will reference this Conda root folder as `%CONDA_ROOT%`.

Note: If you want to use Python 3 it is highly recommended that you install the 64-bit version
of Minconda3. The 32-bit version has far fewer releases of RDKit and may not have the latest
RDKit version.

* Install RDKit
	* `conda install -c rdkit rdkit`
	* You may have to set your path first, e.g. `set path=%PATH%;%CONDA_ROOT%;%CONDA_ROOT%\scripts`
* Add the `%CONDA_ROOT%` folder to your PATH.
	* It can be added to the current user's PATH or the SYSTEM path as you please.


## To compile and register the Excel add-in
First, register the add-in the Windows registry.

Open a Command Prompt *as administrator*. If you do not run the Command Prompt with
administrator rights you will get an 'Error accessing the OLE registry' error when
the script wants to register the add-in in the registry.

Example run:

```
C:\Windows\system32>cd \Users\jan\rdkit4excel\src

C:\Users\Jan\rdkit4excel\src>python RDKitXL_server.py
No IDL changes.
Registering C:\Users\Jan\rdkit4excel\src\RDKitXL.tlb.
Registered: Python.RDKitXL

C:\Users\Jan\rdkit4excel\src>
```

Next, register the add-in in Excel. Start Excel and:

* Click File -> Options -> Add-ins
* Click the "Go" button next to the "Manage: Excel Add-ins" dropdown
* Click the "Automation..." button
	* Choose the "RDKitXL add-in" from the list of available automation servers. Hint: Type "RD" on the keyboard to jump to servers starting with "RD".
	* Click "OK".
		* Note: If you now click the "RDKitXL add-in" in the add-in list you may get a message box asking "Cannot find add-in 'pythoncomloader27.dll'. Delete from list?". Answer "No" to this.
		* Note: If the server is configured to run as an out-of-process service the error message will include a path to an EXE instead of the pythoncomloader27.dll. You should still answer "No".
* Click "OK".
* Restart Excel.

You should now be able to enter `=rdkit_info_version()` in a cell and have the RDKit version string returned.
If you see "#NAME?" in the cell it means that the add-in was not successfully registered and loaded in Excel
after all. In that case, please double-check your PATH settings, retry the registration, and restart Excel.
If that doesn't work, please open an issue on GitHub.


## Adding new functions
You can easily add new Excel-callable functions to the add-in by adding functions to the
`CRDKitXL` class in `RDKitXL_server.py`.

Let's say that you add the infamous example function `rdkit_fat_mw()` that implements
an alternative and utterly useless variant of the molecular weight:

```
#RDKITXL: in:smiles:str, out:float
	def rdkit_fat_mw(self, smiles):
		self.rdkit_info_num_calls = self.rdkit_info_num_calls +1
		smiles = dispatch_to_str(smiles)

		mol = Chem.MolFromSmiles(smiles)		
		if mol != None:
			return Descriptors.MolWt(mol) * 5
		else:
			return 'ERROR: Cannot parse SMILES input.'
```

The source code comments in `RDKitXL_server.py` provide more details on how the `#RDKITXL:` comments should
be written.

The new function needs to be published to the type library. We do that by re-running `RDKitXL_server.py`.
If you just start a command prompt as administrator it will fail:

```
C:\Windows\system32>cd \Users\jan\rdkit4excel\src

C:\Users\Jan\rdkit4excel\src>python RDKitXL_server.py
Compiling C:\Users\Jan\rdkit4excel\src\RDKitXL.idl.
'midl' is not recognized as an internal or external command,
operable program or batch file.
Traceback (most recent call last):
  File "RDKitXL_server.py", line 202, in <module>
    main(sys.argv)
  File "RDKitXL_server.py", line 197, in main
    BuildTypelib(genfile)
  File "RDKitXL_server.py", line 160, in BuildTypelib
    raise RuntimeError("Compiling MIDL failed!")
RuntimeError: Compiling MIDL failed!

C:\Users\Jan\rdkit4excel\src>
```

If you already have a full Visual Studio install you can start a Visual Studio command prompt
as administrator and then you should be able to run without errors.

Alternatively, you can download and install a minimal suite of MS Visual Studio tools.

* Download the "Microsoft visual C 9.0 for Python 2.7" package VCForPython27.msi from https://www.microsoft.com/en-us/download/details.aspx?id=44266 and run it.
	* The MSI package installs per default for the current user only and so doesn't mess up your system settngs.

The above Python 2.7 tools are OK for Python 3 as well. We only use the MIDL compiler to
build a type library and that is independent of the Python version.

Example build in an administrator command prompt (your path to the vcvarsall.bat file will be different):

```
C:\Windows\system32>cd \Users\jan\rdkit4excel\src

C:\Users\jan\rdkit4excel\src>"c:\Users\jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\vcvarsall.bat"
Setting environment for using Microsoft Visual Studio 2008 x86 tools.

C:\Users\jan\rdkit4excel\src>python RDKitXL_server.py
Compiling C:\Users\Jan\rdkit4excel\src\RDKitXL.idl.
Microsoft (R) 32b/64b MIDL Compiler Version 7.00.0555
Copyright (c) Microsoft Corporation. All rights reserved.
Processing C:\Users\Jan\rdkit4excel\src\RDKitXL.idl
RDKitXL.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oaidl.idl
oaidl.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\objidl.idl
objidl.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\unknwn.idl
unknwn.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\wtypes.idl
wtypes.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\basetsd.h
basetsd.h
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\guiddef.h
guiddef.h
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\ocidl.idl
ocidl.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oleidl.idl
oleidl.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\servprov.idl
servprov.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\urlmon.idl
urlmon.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\msxml.idl
msxml.idl
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\oaidl.acf
oaidl.acf
Processing c:\Users\Jan\AppData\Local\Programs\Common\Microsoft\Visual C++ for Python\9.0\WinSDK\Include\ocidl.acf
ocidl.acf
Registering C:\Users\Jan\rdkit4excel\src\RDKitXL.tlb.
Registered: Python.RDKitXL

C:\Users\Jan\rdkit4excel\src>
```

(Re)Start Excel, and the new function should now be ready for use.


## Correcting or changing functions
Changing an implementation of a function is as simple as changing `RDKitXL_server.py` and restarting
Excel.

You only need to rebuild the type library (re-run `RDKitXL_server.py`) when changes are made 
to the published function signatures. That is: When any `#RDKITXL:` comment is changed, removed, or added.


## Deploying new functions to end users
If only function implementations are changed, a deployment of `RDKitXL_server.py` is all that is required.

Changes that cause the type library to change require the following files to be deployed to end users:

```
RDKitXL_server.py
RDKitXL.idl
RDKitXL.idl.previous
RDKitXL.tlb
```

You should not have to re-register the type library and COM service on end user PCs.


## 32-bit Excel and 64-bit Conda
If you have differing bitness of your Python and Office, e.g. 64-bit Python and 32-bit Office
you will have to change the configuration and registration a little.

You must set `_reg_clsctx_` in `RDKitXL_server.py`:

```
# Uncomment the next line to run the server in a separate process:
_reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
```

This will start up the Python COM service in a separate process instead of loading it in the memory space
of Excel. It will be slower, but it allows 32- and 64-bit code to communicate and also isolates your Python
code from Excel so one cannot crash the other.

After changing this line you must run `RDKitXL_server.py` successfully to get the registry entries updated.

After the registry update, you need to make the registered 64-bit COM service visible from Excel.
Start `regedit` and locate the following registry key:

`HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{e4d5c553-ebc8-49ca-bacf-4947ef110fc5}`

Export it to disk, open the .REG file in Notepad, and change all nine occurrences of "\Classes\" to "\Wow6432Node\Classes\" like this:

```
[HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{e4d5c553-ebc8-49ca-bacf-4947ef110fc5}]

becomes

[HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\CLSID\{e4d5c553-ebc8-49ca-bacf-4947ef110fc5}]
```

Load the edited .REG file into the registry by double-clicking it. The "RDKitXL add-in" should now be visible in Excel's list of
automation servers.


## Troubleshooting
To register the Python COM service in debug mode, compile/register with 

```
python rdkitXL_server.py --debug
```

Open PythonWin and open the Tools -> Trace collector debugging tool to watch the messages and print statements.
