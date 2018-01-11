# -*- coding: utf-8 -*-  
import os,bs4,re,argparse,datetime,time,subprocess
from bs4 import BeautifulSoup
from time import sleep
from models import run
from docx.enum.style import WD_STYLE_TYPE
from docx import *
from multiprocessing import Pool
import importlib

########################################################################
baseloca=os.path.dirname(os.getcwd())  #app同级目录
tmploca=os.path.join(baseloca,'templates')
resloca=os.path.join(baseloca,'result')
testfile=os.path.join(tmploca,'awrrpt_1_209199_209200.html')
########################################################################

def LocationJudge(file):
	return 1 if ':' in file or file[0]=='/' else 0

def OScmd(*args):
	os.system("python.exe models.py %s"%args[0])

def manager():
	parser = argparse.ArgumentParser()
	parser.add_argument('-f','--file', action="store",help='once analyze only one file,and you need to add filename to specify which file you want to analyze')
	parser.add_argument('-d','--dirs',action='store',help='analyze all the files which were placed in the ../templates/')
	args = parser.parse_args()
	p=Pool(4)
	if args.file:
		k=LocationJudge(args.file)
		os.system("python.exe models.py %s"%args.file)

	elif args.dirs:
		k=LocationJudge(args.dirs)
		dirs=[]
		baselocation=args.dirs if k else os.path.join(os.getcwd(),args.dirs)
		tmp=os.listdir(args.dirs) if k else os.listdir(baselocation)
		for i in tmp:
			if ".html" in os.path.splitext(i)[1]:
				dirs.append(i)
			else:
				pass

		for i in [os.path.join(baselocation,i) for i in dirs]:
			p.apply_async(OScmd,args=(i,))
		p.close()
		p.join()
	else:
		pass
	
if __name__ == "__main__":
	manager()
	print("文件生成目录为%s"%resloca)

