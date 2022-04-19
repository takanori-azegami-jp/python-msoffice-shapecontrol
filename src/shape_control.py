import sys
import glob
import win32com.client
from abc import ABCMeta, abstractmethod

#-----------------------------------------------------
# Mainクラス
#-----------------------------------------------------
class Main:
	def __init__(self):
		args = sys.argv
		self.folder_pass =  args[1] #フォルダパス

	#main処理
	def main(self):
		print("フォルダパス：" + self.folder_pass)

		#Excelファイルを探索
		list  = File(self.folder_pass , "xlsx")
		for file in list.get_filelist():
			print(file)
			#Excel図形操作
			excel = Excel(file)
			excel.shape_control()

		#Wordファイルを探索
		list  = File(self.folder_pass , "docx")
		for file in list.get_filelist():
			print(file)
			#Word図形操作
			word = Word(file)
			word.shape_control()

		#PowerPointファイルを探索
		list  = File(self.folder_pass , "pptx")
		for file in list.get_filelist():
			print(file)
			#PowerPoint図形操作
			powerpoint = PowerPoint(file)
			powerpoint.shape_control()

#-----------------------------------------------------
# Fileクラス
#-----------------------------------------------------
class File:
	def __init__(self, folder_pass, extension):
		self.folder_pass = folder_pass #フォルダパス
		self.extension = extension #拡張子

	#	フォルダ配下（サブフォルダ含む）の指定拡張子のファイル一覧を取得
	def get_filelist(self):
		return glob.glob( self.folder_pass + "\\**\\*." + self.extension, recursive=True)

#-----------------------------------------------------
# Documentクラス（インターフェース）
#-----------------------------------------------------
class Document(metaclass=ABCMeta):
	@abstractmethod
	def shape_control(self):
		pass
#-----------------------------------------------------
# Excelクラス
#-----------------------------------------------------
class Excel(Document):
	def __init__(self, file_name):
		self.file_name = file_name

	#	Excel図形操作
	def shape_control(self):
		pythoncom.CoInitialize()  #win32com開始前にこれを呼び出す
		excel = win32com.client.DispatchEx("Excel.Application") #ファイルサーバなどのリモート処理ではDispatchではなくDispatchExを使う
		excel.Visible = False

		try:
			excel.DisplayAlerts = False
			doc = excel.Workbooks.Open(self.file_name,False ,False ,None ,"dummy@password" ,"dummy@password" ,True) #パスワード有りはダミーパスワードで無視
			sheet = doc.Worksheets(1) #1シート目を対象
			#シート内の図形を全て処理
			for shape in sheet.Shapes:
				try:
					print(shape.TextFrame.Characters().Text.strip())
				except:
					print("Error")
			doc.Close(False) #閉じる
		except:
			print("Error")
		finally:
			excel.DisplayAlerts = True
			excel.Quit()
			del excel
			pythoncom.CoUninitialize() #終了した後はこれを呼び出す

#-----------------------------------------------------
# Wordクラス
#-----------------------------------------------------
class Word(Document):
	def __init__(self, file_name):
		self.file_name = file_name

	# Word図形操作
	def shape_control(self):
		pythoncom.CoInitialize()  #win32com開始前にこれを呼び出す
		word = win32com.client.DispatchEx("Word.Application") #ファイルサーバなどのリモート処理ではDispatchではなくDispatchExを使う
		word.Visible = False

		try:
			word.DisplayAlerts =  False
			doc = word.Documents.Open(self.file_name, False, False, None,"dummy@password", "dummy@password", False, "dummy@password", "dummy@password") #パスワード有りはダミーパスワードで無視する
			#ドキュメント内の図形を全て処理
			for shape in doc.Shapes:
				try:
					print(shape.TextFrame.TextRange.Text.strip())
				except:
					print("Error")
			doc.Close(False) #閉じる
		except:
			print("Error")
		finally:
			word.DisplayAlerts =True
			word.Quit()
			del word 
			pythoncom.CoUninitialize() #win32com終了時にこれを呼び出す

#-----------------------------------------------------
# PowerPointクラス
#-----------------------------------------------------
class PowerPoint(Document):
	def __init__(self, file_name):
		self.file_name = file_name

	#	PowerPoint図形操作
	def shape_control(self):
		pythoncom.CoInitialize()  #win32com開始前にこれを呼び出す
		powerpoint =win32com.client.DispatchEx("PowerPoint.Application") #ファイルサーバなどのリモート処理ではDispatchではなくDispatchExを使う
		powerpoint.Visible = True

		try:
			doc = powerpoint.Presentations.Open(self.file_name +"::dummy@password::dummy@password") #パスワード有りはダミーパスワードで無視する
			slide = doc.slides[0] #1スライド目を対象
			#ドキュメント内の図形を全て処理
			for shape in slide.shapes:
				try:
					print(shape.TextFrame.TextRange.Text.strip())
				except:
					print("Error")
			doc.Close() #閉じる
		except:
			print("Error")
		finally:
			powerpoint.Quit()
			del powerpoint 
			pythoncom.CoUninitialize() #win32com終了時にこれを呼び出す


#main呼び出し
if __name__ == "__main__":
	main = Main()
	main.main()
