import sys, os  
from win32com.client import Dispatch, constants, gencache 


class Doc2PDF(object):
	def __init__(self,input1):
		self.input1 = input1
		self.output = input1.replace('docx','pdf')
		# print(self.output)

	def change2pdf(self):
		gencache.EnsureModule('{00020905-0000-s0000-C000-000000000046}', 0, 8, 4) 
		w = Dispatch("Word.Application") 
		flag = 0 
		try:
			doc = w.Documents.Open(self.input1, ReadOnly = 1)  
			doc.ExportAsFixedFormat(self.output, constants.wdExportFormatPDF,   
			Item = constants.wdExportDocumentWithMarkup, CreateBookmarks = constants.wdExportCreateHeadingBookmarks) 	
			flag = 0 
		except:
			flag = 1	
		finally:
			w.Quit(constants.wdDoNotSaveChanges)
			sys.exit(flag)

def test():
	input1 = os.path.join(os.getcwd().replace('app','result'),'tzsw_090708_090709.docx')
	print(input1)
	# output = r'C:\Users\galgamish\Desktop\work\程序\awr\AWR_DOCX版初代\result\APEX数据库性能分析报告(201711281700-1800).pdf'
	t=Doc2PDF(input1)
	t.change2pdf()


if __name__ == "__main__":
	test()
