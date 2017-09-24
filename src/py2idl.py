"""Automatic creation of IDL from tagged python source code"""
import re

py2comtype = { 'str':'BSTR',
				'int':'LONG',
				'float':'FLOAT'}


def generate_idl(pythonfile, templatefile='template.idl', generatefile='generated.idl'):

	program = open(pythonfile,'r').read()
	
	
	#Match Properties
	regex = re.compile('\#RDKITXL:\s(.*)\n\s*self\.([a-zA-Z_\-0-9]*)\s')
	propertymatches = regex.findall(program)
	if __name__== "__main__": print(propertymatches)
	#Match Functions
	regex = re.compile('\#RDKITXL:\s(.+)\n\s+def\s+(.*)\((.*)\):\n')
	matches = regex.findall(program)

	

	#Generate IDL lines for .idl file
	idl_lines = []
	#[('prop:str', 'rdkit_version')]
	#[propget, id(1), helpstring("property MyProp1")] HRESULT MyProp1([out, retval] long *pVal);
	for i, propmatch in enumerate(propertymatches):
		tags = propmatch[0].split(':')
		idl_line = '[propget, id(%i), helpstring("Helpstring not exposed in Excel")] HRESULT _stdcall %s([out, retval] %s *pVal);'%(i+1, propmatch[1], py2comtype[tags[1]])
		idl_lines.append(idl_line)
		
	#[('in:smiles:str, inopt:descriptor:str, out:float',
	#  'rdkit_descriptor',
	#  "self, smiles, descriptor='MolLogP'")]
	#[id(2), helpstring("Calculate Descriptor from a SMILES")] HRESULT rdkit_descriptor([in] BSTR *smiles, [in] BSTR *descriptor, [out, retval] FLOAT *val);
	for i, match in enumerate(matches):
		#Each match should generate a line for the IDL
		idl_line = '[id(%i), helpstring("Helpstring not exposed in Excel")] HRESULT _stdcall '%(i+1+len(propertymatches))
		idl_line = idl_line + '%s('%match[1]
		parms = match[0].split(',')
		parms = [parm.strip() for parm in parms] #Strip whitespace
		for parm in parms:
			#Each parm to add, depending on type
			tags = parm.split(':')
			if tags[0] == 'in':
				idl_line = idl_line + '[in] %s %s, '%(py2comtype[tags[2]], tags[1])
			elif tags[0] == 'inopt':
				#identify default option
				funcopt = match[2]
				defaultoption = re.findall('.*%s=u(.*)'%tags[1],funcopt)[0]
				defaultoption = re.sub("'",'"',defaultoption)# ' must be " in the .idl file. TODO what if default option is 'something""something'??
				idl_line = idl_line + '[in, defaultvalue(%s)] %s %s, '%(defaultoption, py2comtype[tags[2]], tags[1]) #Default Value must be VARIANT or midl.exe complains
			elif tags[0] == 'out':
				idl_line = idl_line + '[out, retval] %s *val);'%py2comtype[tags[1]]
			else:
				print("Warning, unknown parameter type",match, parm)

		idl_lines.append(idl_line)
	
	if __name__ == "__main__":
		for l in idl_lines:
			print(l)

	#Join lines	
	insert = '\n'.join(idl_lines)

	#Substitue in template
	template = open(templatefile,'r')
	regex = re.compile('#RDKitXL')
	idl = regex.sub(insert, template.read())
	template.close()
	
	#Write new .idl file	
	f=open(generatefile,'w')
	f.write(idl)
	f.close()
	
	return generatefile
	
	
if __name__ == '__main__':
	# If started directly, assume that we are testing and generate a 'tester.idl' file.
	genfile = generate_idl('RDKitXL_server.py', generatefile='tester.idl')
	with file(genfile,'r') as f:
		print(f.read())

