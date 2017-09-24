# Modified from Pippo_server.py test example from win32com.
import sys, os
import pythoncom
import win32com
import winerror
from win32com.server.util import wrap
from win32com.client import Dispatch
import win32api

from py2idl import generate_idl

import rdkit
from rdkit import Chem
from rdkit.Chem import Descriptors, AllChem

# 'six' module is used by isinstance() to detect a string instance in a simple way.
from six import string_types

def dispatch_to_str(possible_dispatch):
	if not isinstance(possible_dispatch, string_types):
		return str(Dispatch(possible_dispatch))
	else:
		return str(possible_dispatch)


class CRDKitXL:
	#
	# COM registry declarations.
	#
	_reg_clsid_ = '{e4d5c553-ebc8-49ca-bacf-4947ef110fc5}'
	_reg_desc_ = "RDKitXL object"
	_reg_progid_ = "Python.RDKitXL"
	_reg_options_ = {"Programmable":''}

	# Per default the server runs as an in-process server, meaning that it gets loaded
	# as a DLL in Excel's memory space. If you run a 32-bit Excel and a 64-bit Python
	# you will have to uncomment the following line to run the server in a separate
	# process, since you cannot load 64-bit code into a 32-bit process.
	# The out-of-process service is slower, but has the benefit that it is isolated from
	# Excel and so cannot crash Excel.
	# Uncomment the next line to run the server in a separate process:
	# _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

	### Link to typelib
	_typelib_guid_ =  '{da5bc306-15b4-498a-9c1c-f560aa8b5c32}'
	_typelib_version_ = 1, 0
	_com_interfaces_ = ['IRDKitXL']

	def __init__(self):

# Structured comments starting with RDKITXL are used to markup published
# properties and functions. Only published properties and functions will
# be visible in Excel.
#
# Structured comments must immediately precede the line defining the
# property or function. So any comments describing the property or function
# must be placed *before* the structured comment.
#
# The marked-up properties and functions generate corresponding properties
# and functions in the RDKitXL.idl file which is then compiled to an
# RDKitXL.tlb type library that Excel can read.
#
# All data types have to be made explicit in the generated IDL. Accepted
# types in the structured comments are 'int', 'str', and 'float'.

# Get RDKit version string.
#RDKITXL: prop:str
		self.rdkit_info_version = rdkit.__version__

# Get number of function calls done. Requires all published functions to
# increment this value.
#RDKITXL: prop:int
		self.rdkit_info_num_calls = 0

# String parameters with default values:
#   Values must be unicode strings.
#   Values may not contain " inside the string.
#
# Debug output: Use win32api.OutputDebugString() and something like DbgView
# to trace it. Do not use print() - it will cause "Bad file descriptor" failures
# (suspect that it is a threading issue).

# Calculates a named RDKit descriptor value from a SMILES input.
#RDKITXL: in:smiles:str, inopt:descriptor:str, out:float
	def rdkit_descriptor(self, smiles, descriptor=u'MolLogP'):
		try:
			self.rdkit_info_num_calls = self.rdkit_info_num_calls + 1
			# win32api.OutputDebugString(str(type(smiles)) + " " + str(type(descriptor)))
			smiles = dispatch_to_str(smiles)
			descriptor = dispatch_to_str(descriptor)

			myfunction = getattr(Descriptors, descriptor)
			mol = Chem.MolFromSmiles(smiles)
			if mol != None:
				return myfunction(mol)
			else:
				# OK, so you are wondering how on Earth a function that is marked up
				# as out:float can return a string ?! Me too, but it works :-), and
				# makes it possible to show a decent error to the end user.
				# Apparently COM returns the value as a variant regardless of the
				# specified IDL retval type (?)...
				return 'ERROR: Cannot parse SMILES input.'
		except Exception as e:
			return "ERROR: " + str(e)
		
# Generates a molfile with 2D coordinates from SMILES input. Useful for depiction.
#RDKITXL: in:smiles:str, out:str
	def rdkit_smiles_to_molblock(self, smiles):
		self.rdkit_info_num_calls = self.rdkit_info_num_calls +1
		# win32api.OutputDebugString(str(type(smiles)))
		smiles = dispatch_to_str(smiles)

		mol = Chem.MolFromSmiles(smiles)		
		if mol != None:
			# Add coords for depiction.
			AllChem.Compute2DCoords(mol)
			return Chem.MolToMolBlock(mol)
		else:
			return 'ERROR: Cannot parse SMILES input.'


def BuildTypelib(idlfile = "RDKitXL.idl"):
	this_dir = os.path.dirname(__file__)
	idl = os.path.abspath(os.path.join(this_dir, idlfile))
	basename = idlfile.split('.')[0]
	tlb=os.path.splitext(idl)[0] + '.tlb'
	prev_idl = idl + ".previous"

	this_idl_txt = "".join(open(idl, 'r').readlines())
	previous_idl_txt = "does not exist"
	if os.path.isfile(prev_idl):
		previous_idl_txt = "".join(open(prev_idl, 'r').readlines())

	if this_idl_txt != previous_idl_txt:
		print("Compiling %s" % (idl,))
		rc = os.system ('midl "%s"' % (idl,))
		if rc:
			raise RuntimeError("Compiling MIDL failed!")
		# Can't work out how to prevent MIDL from generating the stubs.
		# just nuke them
		for fname in ("dlldata.c %s_i.c %s_p.c %s.h"%(basename, basename, basename)).split():
			os.remove(os.path.join(this_dir, fname))
		open(prev_idl, 'w').write("".join(open(idl, 'r').readlines()))

	else:
		print("No IDL changes.")

	print("Registering %s" % (tlb,))
	tli=pythoncom.LoadTypeLib(tlb)
	pythoncom.RegisterTypeLib(tli,tlb)

def UnregisterTypelib():
	k = CRDKitXL
	try:
		pythoncom.UnRegisterTypeLib(k._typelib_guid_, 
									k._typelib_version_[0], 
									k._typelib_version_[1], 
									0, 
									pythoncom.SYS_WIN32)
		print("Unregistered typelib")
	except pythoncom.error as details:
		if details[0]==winerror.TYPE_E_REGISTRYACCESS:
			pass
		else:
			raise

def main(argv=None):
	if argv is None: argv = sys.argv[1:]
	genfile = generate_idl(__file__, generatefile="RDKitXL.idl") #Introspective code, Yay!
	if '--unregister' in argv:
		# Unregister the type-libraries.
		UnregisterTypelib()
	else:
		# Build and register the type-libraries.
		BuildTypelib(genfile)
	import win32com.server.register 
	win32com.server.register.UseCommandLine(CRDKitXL)

if __name__=='__main__':
	main(sys.argv)
