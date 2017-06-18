# modified from Pippo_server.py test example from win32com
import sys, os
import pythoncom
import win32com
import winerror
from win32com.server.util import wrap
import rdkit
from rdkit import Chem
from rdkit.Chem import Descriptors, AllChem
from win32com.client import Dispatch
from py2idl import generate_idl

class CRDKitXL:
	#
	# COM declarations	
	#
	_reg_clsid_ = '{e4d5c553-ebc8-49ca-bacf-4947ef110fc5}' #"{05AC1CCE-3F9B-4d9a-B0B5-DFE8BE45AFA8}"
	_reg_desc_ = "RDKitXL object"
	_reg_progid_ = "Python.RDKitXL"
	_reg_options_ = {"Programmable":''}
	#_reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER	
	###
	#To generate new GUID's if starting a completely new project
	#import uuid
	#str(uuid.uuid4())
	
	### Link to typelib
	_typelib_guid_ =  '{da5bc306-15b4-498a-9c1c-f560aa8b5c32}'#'{41059C57-975F-4B36-8FF3-C5117426647A}' #
	_typelib_version_ = 1, 0
	_com_interfaces_ = ['IRDKitXL']

	def __init__(self):
#RDKITXL: prop:str
		self.rdkit_version = rdkit.__version__
#RDKITXL: prop:int
		self.number_of_calls = 0
		

#string variables given as default must be unicode and not contain " inside the string.
#RDKITXL: in:smiles:str, inopt:descriptor:str, out:float
	def rdkit_descriptor(self, smiles, descriptor=u'MolLogP'):
		self.number_of_calls = self.number_of_calls +1
		print type(smiles), type(descriptor)
		if not isinstance(smiles, unicode):
			print "converting smiles"
			smiles = str(Dispatch(smiles))
		if not isinstance(descriptor, unicode):
			print "converting descriptor"
			descriptor = str(Dispatch(descriptor))
		myfunction = getattr(Descriptors, descriptor)
		mol = Chem.MolFromSmiles(smiles)
		if mol != None:
			return myfunction(mol)
		else:
			return 'Error in parsing SMILES'
			
#RDKITXL: in:smiles:str, out:str
	def rdkit_SmilesToMolBlock(self, smiles):
		self.number_of_calls = self.number_of_calls +1
		print type(smiles)
		if not isinstance(smiles, unicode):
			print "converting smiles"
			smiles = str(Dispatch(smiles))
		mol = Chem.MolFromSmiles(smiles)
		
		if mol != None:
			#Add coords for depiction
			AllChem.Compute2DCoords(mol)
			return Chem.MolToMolBlock(mol)
		else:
			return 'Error in parsing SMILES'
			
	
def BuildTypelib(idlfile = "RDKitXL.idl"):
	from distutils.dep_util import newer
	this_dir = os.path.dirname(__file__)
	idl = os.path.abspath(os.path.join(this_dir, idlfile))
	basename = idlfile.split('.')[0]
	tlb=os.path.splitext(idl)[0] + '.tlb'
	if newer(idl, tlb):
		print "Compiling %s" % (idl,)
		rc = os.system ('midl "%s"' % (idl,))
		if rc:
			raise RuntimeError("Compiling MIDL failed!")
		# Can't work out how to prevent MIDL from generating the stubs.
		# just nuke them
		for fname in ("dlldata.c %s_i.c %s_p.c %s.h"%(basename, basename, basename)).split():
			os.remove(os.path.join(this_dir, fname))
	
	print "Registering %s" % (tlb,)
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
		print "Unregistered typelib"
	except pythoncom.error, details:
		if details[0]==winerror.TYPE_E_REGISTRYACCESS:
			pass
		else:
			raise

def main(argv=None):
	if argv is None: argv = sys.argv[1:]
	genfile = generate_idl(__file__, generatefile="RDkitXL.idl") #Introspective code, Yay!
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
