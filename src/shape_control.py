import win32com.client
import pythoncom
from abc import ABCMeta, abstractmethod

#main処理
def main():
	#Excel図形操作
	excel = Excel(r"Excelファイルのフルパス")
	excel.shape_control()

	#Word図形操作
	word = Word(r"Wordファイルのフルパス")
	word.shape_control()

	#PowerPoint図形操作
	powerpoint = PowerPoint(r"PowerPointファイルのフルパス")
	powerpoint.shape_control()

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
		excel.Visible = True
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
		word.Visible = True
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
	main()
