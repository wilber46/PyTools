#!/usr/bin/python
"""
switch all html files to word document.
depends on win32com.
usage:
python extractWizData.py <data_dir> <out_dir>
"""

import os,sys
import zipfile
import shutil
import platform
import tempfile
import win32com.client

def handlePath(in_path,out_path,html):
	for f in os.listdir(in_path):
		path = in_path + os.sep + f
		# print(path)
		if os.path.isdir(path):
			handlePath(path,out_path + '/' + f,html)
		elif os.path.isfile(path):
			if html:
				unzipFile(path,out_path)
			else:
				handleHtml(path,out_path)
		pass
	pass

def handleHtml(file,out_path):
	# print(file + "\n" + out_path)
	(filepath,filename) = os.path.split(file);
	(name,ext) = os.path.splitext(filename)
	(prefix,suffix) = os.path.split(filepath)

	(outPrefix,outSuffix) = os.path.split(out_path)
	if filename == 'index.html':
		if not os.path.exists(outPrefix):
			os.makedirs(outPrefix)

		destFile = outPrefix + os.sep + suffix + ".doc"
		if not os.path.exists(destFile):
			handleDoc(file,destFile)

		# shutil.copy(file,outPrefix + os.sep + suffix + ".doc")

def handleDoc(html,document):
	# print('html -> ' + html)
	# print('doc -> ' + document)
	print("generate file:\t%s\n"%(document))

	word = win32com.client.Dispatch('Word.Application')

	doc = word.Documents.Add(html)
	doc.SaveAs(document, FileFormat=0)
	doc.Close()

	word.Quit()

def unzipFile(file,out_path):
	(filepath,filename) = os.path.split(file);
	(name,ext) = os.path.splitext(filename)
	if ext == '.ziw':
		destpath = out_path + os.sep + name
		destpath = destpath.replace(' ','_')
		if not os.path.exists(destpath):
			# print('mkdirs ' + destpath)
			os.makedirs(destpath)

		zfile = zipfile.ZipFile(file,'r')
		for pfile in zfile.namelist():
			zfile.extract(pfile,destpath)

		zfile.close()



if __name__ == '__main__':
	sysInfo = platform.system()
	print("current system platform is %s"%(sysInfo))
	if not sysInfo.lower() == "windows":
		print("only support windows!!!")
		exit(0)

	if len(sys.argv) < 2:
		print('please input in_file and out_file')
		exit(0)
    
	in_path = sys.argv[1]
	if not os.path.exists(in_path):
		print('not exist path ' + in_path)
		exit(0)

	print("in_path = " + in_path)

	out_path = sys.argv[2]
	print("out_path = " + out_path)

	# if not os.access(out_path,os.W_OK):
	# 	print("out path can not write,check please.. path = " + out_path)

	# in_path = '/home/wilber/tmp/test/apple'
	# out_path = '/home/wilber/tmp/test/apple_out'

	# in_path = "F:/tmp/test/apple"
	# out_path = "F:/tmp/test/wizDoc"
    	
	tmp_path = tempfile.mkdtemp()

	# handlePath(in_path,out_path)

	handlePath(in_path,tmp_path,True)
	handlePath(tmp_path,out_path,False)

	if os.path.exists(tmp_path):
		shutil.rmtree(tmp_path)
