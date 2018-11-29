# cmd /k cd /d "$(CURRENT_DIRECTORY)" & C:\Python34\python "$(FULL_CURRENT_PATH)" &ECHO&PAUSE&EXIT
# cmd /k cd /d "$(CURRENT_DIRECTORY)" & D:\Python36\python "$(FULL_CURRENT_PATH)" &ECHO&PAUSE&EXIT

#使用pyinstaller打包应用程序时logo.ico，AdvertiserProcessor.py存放在以下目录下
# cd D:\LB\JAVASCRIPT
# D:\Python36\Scripts\pyinstaller -D -i logo.ico AdvertiserProcessor.py
# D:\Python36\Scripts\pyinstaller -D -w -i logo.ico AdvertiserProcessor.py
#打包完成后的文件会存在 D:\LB\JAVASCRIPT\dist中


import sys
import os
import time
import re

from PyQt5.QtWidgets import QApplication,QMainWindow,QHeaderView,QMessageBox,QFileDialog,QTableWidgetItem
from PyQt5.uic import loadUi
from PyQt5.QtGui import QStandardItemModel,QStandardItem

from xlrd import open_workbook
from xlwt import Workbook,easyxf,Formula

import sqlite3

import jieba
jieba.set_dictionary("./dict.txt")  
jieba.initialize()  

#————————————————————————————————————————————————————
#字符串处理函数
def stripLocation(s):
	#剔除所有地名
	wordArray=jieba.cut(s)
	a=[]
	for word in wordArray:
		w=re.sub(r".*[省市区县镇乡村]$","",word)
		if w is not "":
			a.append(w)
	
	result="".join(a)
	return result

def stripCompany(s):
	#剔除公司类型名
	wordArray=jieba.cut(s)
	a=[]
	for word in wordArray:
		w=re.sub(r"有限|责任|股份|合伙|公司","",word)
		if w is not "":
			a.append(w)
	
	result="".join(a)
	return result
	

#————————————————————————————————————————————————————

class MainWindow(QMainWindow):
	def __init__(self,parent=None):
		super(MainWindow,self).__init__(parent)
		loadUi("advertiser.ui",self)
		
		#按钮事件连接
		self.pushButton.clicked.connect(self.searchAdvertiser)
		self.lineEdit.returnPressed.connect(self.searchAdvertiser)
		self.pushButton_2.clicked.connect(self.update)
		self.toolButton.clicked.connect(self.findSourceFile)
		self.pushButton_3.clicked.connect(self.processAdvertisers)
		self.toolButton_2.clicked.connect(self.findSourceFile0)
		self.pushButton_4.clicked.connect(self.one2one)
		self.toolButton_3.clicked.connect(self.findSourceFile1)
		self.toolButton_4.clicked.connect(self.findSourceFile2)
		self.pushButton_5.clicked.connect(self.statistics)
		#self.tabWidget.currentChanged.connect(self.tabCurrentChanged)
		
	#__________________________________________		
	def findSourceFile(self):
		fileName,fileType= QFileDialog.getOpenFileName(self,"选取要导入本地数据库的广告主总单","D:/","Excel Files (*.xlsx);;Excel Files (*.xls)")
		self.lineEdit_3.setText(fileName)
		
	def findSourceFile0(self):
		fileName,fileType= QFileDialog.getOpenFileName(self,"选取要导入本地数据库的广告主总单","D:/","Excel Files (*.xls);;Excel Files (*.xlsx)")
		self.lineEdit_4.setText(fileName)
		
	def findSourceFile1(self):
		fileName,fileType= QFileDialog.getOpenFileName(self,"选取要导入本地数据库的广告主总单","D:/","Excel Files (*.xls);;Excel Files (*.xlsx)")
		self.lineEdit_5.setText(fileName)
	
	def findSourceFile2(self):
		fileName,fileType= QFileDialog.getOpenFileName(self,"选取要导入本地数据库的广告主总单","D:/","Excel Files (*.xls);;Excel Files (*.xlsx)")
		self.lineEdit_6.setText(fileName)
		
	#__________________________________________
	#检索本地数据库中已经存在的广告主
	def searchAdvertiser(self):
		keywords=self.lineEdit.text()
		self.statusbar.showMessage(keywords)
		
		if keywords.strip() is "":
			#self.statusbar.showMessage('请输入查询关键词！')
			QMessageBox.critical(self,"注意","请输入查询关键词！")
			return
		else:
			keywords=keywords.strip()
			
		keywords=stripLocation(stripCompany(keywords))
		
		keywordscollection=jieba.cut(keywords)
		keys=""
		
		for keyword in keywordscollection:
			keys=keys+"%"+keyword		
		
		sql="SELECT DISTINCT ename, cname FROM advertiser WHERE cname LIKE '"+keys+"%' OR ename LIKE '"+keys+"%' ORDER BY ename ASC LIMIT 100000"			
			
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		cursor.execute(sql)	
		conn.commit
		
		rows = cursor.fetchall()
		
		model=QStandardItemModel()
		model.setHorizontalHeaderLabels(["广告主（中文名）","广告主（英文名）"])
		
		if rows:
			aCount=len(rows)
			self.statusbar.showMessage("查到"+str(aCount)+"条相关广告主")
			rowNum=0
			for row in rows:				
				item0=QStandardItem(row[0])
				item1=QStandardItem(row[1])
				model.setItem(rowNum,0,item1)
				model.setItem(rowNum,1,item0)
				rowNum+=1
				
			self.tableView.setModel(model)
			self.tableView.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
			
		else:
			QMessageBox.critical(self,"注意","查不到!")
			
		cursor.close()
		conn.close()
		
	#————————————————————————————————————————————————————
	def update(self):
		sourceFileName=self.lineEdit_3.text()
		excelFileName=sourceFileName.strip()
		filenameOk=excelFileName.endswith(".xlsx") or excelFileName.endswith(".xls")

		if excelFileName is "" or not filenameOk :
			QMessageBox.critical(self,"注意","请选择要导入广告主总表Excel文件！")
			return
		#___________________________________________
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		sql='DROP TABLE IF EXISTS advertiser'
		cursor.execute(sql)
		conn.commit()
		
		sql='CREATE TABLE IF NOT EXISTS advertiser(ename varchar(128), cname varchar(128),bname varchar(128))'
		cursor.execute(sql)
		conn.commit()	

		#将广告主总表导入本地sqlite数据库	
		self.statusbar.showMessage("正在读取EXCEL数据 请耐心等待......")
		
		recordCount=0
		
		args=[]
		book=open_workbook(excelFileName)
		for sheet in book.sheets():
			for i in range(sheet.nrows):
				v1=sheet.cell(i,1).value
				v2=sheet.cell(i,2).value
				v3=sheet.cell(i,3).value
				arg=tuple([v1,v2,v3])
				args.append(arg)
				
				recordCount+=1
				self.statusbar.showMessage(str(recordCount))
		
		#批量执行sql数据插入
		try:
			sql = 'INSERT INTO advertiser(ename,cname,bname) values(?,?,?)'
			cursor.executemany(sql, args)
		except Exception as e:
			print(e)
			self.statusbar.showMessage(e)
		finally:
			cursor.close()
			conn.commit()
			conn.close()	
			QMessageBox.information(self,"恭喜","数据已成功导入数据库advertisers.db")
			
	
	
	#________________________________________________________________
	#新广告主表处理
	def processAdvertisers(self):
		targetFileName=self.lineEdit_4.text().strip()
		#先判断文件选择是否正确
		if targetFileName is "":
			QMessageBox.critical(self,'注意', '请先选择要处理的新广告主EXCEL表!')
			return
		filenameOk=targetFileName.endswith(".xlsx") or targetFileName.endswith(".xls")
		if filenameOk is False:
			QMessageBox.critical(self,"注意","请输入正确的EXCEL文件名！")
			return
		print("当前选定要处理的广告主EXCEL表：",targetFileName)
		
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		sql='DROP TABLE IF EXISTS workbook'
		cursor.execute(sql)
		conn.commit()
		
		sql='CREATE TABLE IF NOT EXISTS workbook(a varchar(128), b varchar(128), c varchar(128), d varchar(128), e varchar(128), f varchar(128), g varchar(128), h varchar(128), i varchar(128), j varchar(128), k varchar(128), l varchar(128), m varchar(128), n varchar(128), o varchar(128), p varchar(128), q varchar(128), r varchar(128), s varchar(128), t varchar(128), u varchar(128), v varchar(128), w varchar(128), x varchar(128), y varchar(128), z varchar(128))'
		cursor.execute(sql)
		conn.commit()
		
		#将工作表数据存入SQLITE数据库表
		args=[]
		book=open_workbook(targetFileName)
		for sheet in book.sheets():
			for i in range(1,sheet.nrows):
				vArray=[]
				for j in range(26):
					if j<sheet.ncols:
						v=str(sheet.cell(i,j).value).strip()
					else:
						v=""
					vArray.append(v)
				arg=tuple(vArray)
				args.append(arg)
		
		#批量执行sql数据插入
		try:
			sql = 'INSERT INTO workbook(a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
			cursor.executemany(sql, args)
		except Exception as e:
			print(e)
		finally:
			cursor.close()
			conn.commit()
			conn.close()
			self.statusbar.showMessage("Excel表已经成功导入本地SQLITE数据库")
			
		
		#--------------------------------------------------------------------------------
		#测试是否成功
		
		self.statusbar.showMessage("正在比对数据，请稍候...")
		
		#创建EXCEL表格准备写入数据
		book = Workbook()
		sheet= book.add_sheet('Sheet1',cell_overwrite_ok=True)
		#单元格格式水蓝色
		cellStyle=easyxf('font: name 宋体;pattern: pattern solid, fore_colour aqua')
		
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		#通过联合多表UPDATE来导入已有广告主到新表中	
		sql="UPDATE workbook SET j=(SELECT cname FROM advertiser WHERE bname=workbook.d) WHERE j=''"
		cursor.execute(sql)	
		conn.commit
		
		sql="UPDATE workbook SET l=(SELECT ename FROM advertiser WHERE cname=workbook.j) WHERE l=''"
		cursor.execute(sql)	
		conn.commit
		
		
		
		#从本地数据库表中查询数据输出到EXCEL表
		sql="SELECT * FROM workbook ORDER BY l DESC,j ASC,d ASC"
		cursor.execute(sql)	
		conn.commit
		rows = cursor.fetchall()		
					
		if rows:	
			rown=1
			for row in rows:	
				for i in range(26):			
					if (i==3 or i==9):
						if row[i]:
							link='HYPERLINK("https://www.baidu.com/s?wd='+str(row[i]).replace("&","+")+'";"'+str(row[i])+'")'
							sheet.write(rown,i,Formula(link),cellStyle)
						else:
							sheet.write(rown,i,"",cellStyle)
					elif i==5:
						link='HYPERLINK("https://www.baidu.com/s?wd=site:www.tianyancha.com '+str(row[5])+'";"'+str(row[5])+'")'
						sheet.write(rown,i,Formula(link),cellStyle)
					elif i==17:
						link='HYPERLINK("https://translate.google.cn/#zh-CN/en/'+str(row[9])+'";"英译")'
						sheet.write(rown,i,Formula(link),cellStyle)
					elif i==18:
						link='HYPERLINK("https://www.baidu.com/s?wd=site:www.tianyancha.com '+str(row[9])+'";"天眼查")'
						sheet.write(rown,i,Formula(link),cellStyle)
					else:
						if row[i]:
							sheet.write(rown,i,str(row[i]),cellStyle)
						else:
							sheet.write(rown,i,"",cellStyle)				
				rown=rown+1
			self.statusbar.showMessage("共计写入"+str(rown)+"行")
		
		#设置EXCEL表格格式
		sheet.col(0).hidden=True
		sheet.col(1).hidden=True
		sheet.col(2).hidden=True
		sheet.col(3).width=8000
		sheet.col(4).hidden=True
		sheet.col(5).width=3000
		sheet.col(6).hidden=True
		sheet.col(7).hidden=True
		sheet.col(8).hidden=True
		sheet.col(9).width=10000
		sheet.col(10).hidden=True
		sheet.col(11).width=10000	
		sheet.col(12).width=2000	
		sheet.col(13).width=2000
		sheet.col(14).width=2000
		sheet.col(15).width=2000
		sheet.col(16).width=2000
		sheet.col(17).width=2000
		sheet.col(18).width=2000
		
		savedFileName=re.sub(r'\.xlsx|\.xls','',targetFileName)+'_ok.xls'
		if os.path.exists(savedFileName):
			t=time.localtime()
			savedFileName=re.sub(r'\.xlsx|\.xls','',targetFileName)+'_ok_'+str(t.tm_year)+'_'+str(t.tm_mon)+'_'+str(t.tm_mday)+'_'+str(t.tm_hour)+'_'+str(t.tm_min)+'_'+str(t.tm_sec)+'.xls'
		book.save(savedFileName)
		
		hyperlinker="<a href='file:///"+savedFileName+"' title='点击打开'>"+savedFileName+"</a>"
		QMessageBox.information(self,"注意","处理完的广告主文件已经保存到<br/><br/>"+hyperlinker+"<br/>（点击链接打开文件）<br/><br/>")
		
		#关闭SQLITE数据库
		cursor.close()
		conn.close()
			
	#————————————————————————————————————————————————————
	#两表互导
	def one2one(self):
		sourceFile0=self.lineEdit_5.text()
		sourceFile1=self.lineEdit_6.text()
		
		print(sourceFile0)
		print(sourceFile1)
		
		if sourceFile0.strip()=="" or sourceFile1.strip()=="":
			QMessageBox.critical(self,"注意","请选择数据读取表和写入表！")
			return
		else:
			sourceFile0=sourceFile0.strip()
			sourceFile1=sourceFile1.strip()			
		
		filenameOk0=sourceFile0.endswith(".xlsx") or sourceFile0.endswith(".xls")
		filenameOk1=sourceFile1.endswith(".xlsx") or sourceFile1.endswith(".xls")
		filenameOk=filenameOk0 and filenameOk1
		
		if filenameOk is False:
			QMessageBox.critical(self,"注意","必须选择.xls或者.xlsx格式EXCEL文件！")
			return
		
		columnName0=self.comboBox_0.currentText()
		columnName1=self.comboBox_1.currentText()
		columnName2=self.comboBox_2.currentText()
		columnName3=self.comboBox_3.currentText()
		columnName4=self.comboBox_4.currentText()
		columnName5=self.comboBox_5.currentText()
		
		print(columnName0,columnName1,columnName2,columnName3,columnName4,columnName5)	
		#--------------------------------------------------------
		
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		sql='DROP TABLE IF EXISTS book0'
		cursor.execute(sql)
		conn.commit()
		
		sql='CREATE TABLE IF NOT EXISTS book0(a varchar(128), b varchar(128), c varchar(128), d varchar(128), e varchar(128), f varchar(128), g varchar(128), h varchar(128), i varchar(128), j varchar(128), k varchar(128), l varchar(128), m varchar(128), n varchar(128), o varchar(128), p varchar(128), q varchar(128), r varchar(128), s varchar(128), t varchar(128), u varchar(128), v varchar(128), w varchar(128), x varchar(128), y varchar(128), z varchar(128))'
		cursor.execute(sql)
		conn.commit()
		
		#将工作表数据存入SQLITE数据库表
		args=[]
		book0=open_workbook(sourceFile0)
		for sheet in book0.sheets():
			for i in range(1,sheet.nrows):
				vArray=[]
				for j in range(26):
					if j<sheet.ncols:
						v=str(sheet.cell(i,j).value).strip()
					else:
						v=''
					vArray.append(v)
				arg=tuple(vArray)
				args.append(arg)
		
		#批量执行sql数据插入
		try:
			sql = 'INSERT INTO book0(a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
			cursor.executemany(sql, args)
		except Exception as e:
			print(e)
		finally:
			cursor.close()
			conn.commit()
			conn.close()
			self.statusbar.showMessage("数据读取表已经成功导入本地SQLITE数据库")
			
		#-----------------------------------------------------------------------
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		sql='DROP TABLE IF EXISTS book1'
		cursor.execute(sql)
		conn.commit()
		
		sql='CREATE TABLE IF NOT EXISTS book1(a varchar(128), b varchar(128), c varchar(128), d varchar(128), e varchar(128), f varchar(128), g varchar(128), h varchar(128), i varchar(128), j varchar(128), k varchar(128), l varchar(128), m varchar(128), n varchar(128), o varchar(128), p varchar(128), q varchar(128), r varchar(128), s varchar(128), t varchar(128), u varchar(128), v varchar(128), w varchar(128), x varchar(128), y varchar(128), z varchar(128))'
		cursor.execute(sql)
		conn.commit()
		
		#将工作表数据存入SQLITE数据库表
		args=[]
		book1=open_workbook(sourceFile1)
		for sheet in book1.sheets():
			for i in range(1,sheet.nrows):
				vArray=[]
				for j in range(26):
					if j<sheet.ncols:
						v=str(sheet.cell(i,j).value).strip()
					else:
						v=''
					vArray.append(v)
				arg=tuple(vArray)
				args.append(arg)
		
		#批量执行sql数据插入
		try:
			sql = 'INSERT INTO book1(a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
			cursor.executemany(sql, args)
		except Exception as e:
			print(e)
		finally:
			cursor.close()
			conn.commit()
			conn.close()
			self.statusbar.showMessage("数据写入表已经成功导入本地SQLITE数据库")
		#--------------------------------------------------------------------------------
		#比较两个表	
		self.statusbar.showMessage("正在比对数据 请稍候")	
		#创建EXCEL表格准备写入数据
		book = Workbook()
		sheet= book.add_sheet('Sheet1',cell_overwrite_ok=True)
		#单元格格式水蓝色
		cellStyle=easyxf('font: name 宋体;pattern: pattern solid, fore_colour aqua')	
		#建立SQLITE3数据库连接
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()	
		#通过联合多表UPDATE来导入已有广告主到新表中
		#覆盖原数据
		#sql="UPDATE book1 SET "+columnName3+"=(SELECT "+columnName2+" FROM book0 WHERE book0."+columnName0+"=book1."+columnName1+")"
		#不覆盖原数据
		sql="UPDATE book1 SET "+columnName3+"=(SELECT "+columnName2+" FROM book0 WHERE book0."+columnName0+"=book1."+columnName1+") WHERE book1."+columnName3+"=''"
		print(sql)
		cursor.execute(sql)	
		conn.commit
		#覆盖原数据
		#sql="UPDATE book1 SET "+columnName5+"=(SELECT "+columnName4+" FROM book0 WHERE book0."+columnName0+"=book1."+columnName1+")"
		#不覆盖原数据
		sql="UPDATE book1 SET "+columnName5+"=(SELECT "+columnName4+" FROM book0 WHERE book0."+columnName0+"=book1."+columnName1+") WHERE book1."+columnName5+"=''"
		print(sql)
		cursor.execute(sql)	
		conn.commit	
		
		
		#从本地数据库表中查询数据输出到EXCEL表
		sql="SELECT * FROM book1 ORDER BY l DESC,j ASC,d ASC"
		cursor.execute(sql)	
		conn.commit
		rows = cursor.fetchall()		
		
		
		if rows:	
			rown=1
			for row in rows:	
				for i in range(26):			
					if (i==3 or i==9):
						if row[i]:
							link='HYPERLINK("https://www.baidu.com/s?wd='+str(row[i]).replace("&","+")+'";"'+str(row[i])+'")'
							#print(link)
							sheet.write(rown,i,Formula(link),cellStyle)
						else:
							sheet.write(rown,i,"",cellStyle)
					elif i==5:
						link='HYPERLINK("https://www.baidu.com/s?wd=site:www.tianyancha.com '+str(row[5])+'";"'+str(row[5])+'")'
						sheet.write(rown,i,Formula(link),cellStyle)
					elif i==17:
						link='HYPERLINK("https://translate.google.cn/#zh-CN/en/'+str(row[9])+'";"英译")'
						sheet.write(rown,i,Formula(link),cellStyle)
					elif i==18:
						link='HYPERLINK("https://www.baidu.com/s?wd=site:www.tianyancha.com '+str(row[9])+'";"天眼查")'
						sheet.write(rown,i,Formula(link),cellStyle)
					else:
						if row[i]:
							sheet.write(rown,i,str(row[i]),cellStyle)
						else:
							sheet.write(rown,i,"",cellStyle)
				rown=rown+1

			self.statusbar.showMessage("共计写入"+str(rown)+"行")
			
		
		#设置EXCEL表格格式
		sheet.col(0).hidden=True
		sheet.col(1).hidden=True
		sheet.col(2).hidden=True
		sheet.col(3).width=8000
		sheet.col(4).hidden=True
		sheet.col(5).width=3000
		sheet.col(6).hidden=True
		sheet.col(7).hidden=True
		sheet.col(8).hidden=True
		sheet.col(9).width=10000
		sheet.col(10).hidden=True
		sheet.col(11).width=10000	
		sheet.col(12).width=2000	
		sheet.col(13).width=2000
		sheet.col(14).width=2000
		sheet.col(15).width=2000
		sheet.col(16).width=2000
		sheet.col(17).width=2000
		sheet.col(18).width=2000
				
		savedFileName=re.sub(r'\.xlsx|\.xls','',sourceFile1)+'_ok.xls'
		if os.path.exists(savedFileName):
			t=time.localtime()
			savedFileName=re.sub(r'\.xlsx|\.xls','',sourceFile1)+'_ok_'+str(t.tm_year)+'_'+str(t.tm_mon)+'_'+str(t.tm_mday)+'_'+str(t.tm_hour)+'_'+str(t.tm_min)+'_'+str(t.tm_sec)+'.xls'
		
		book.save(savedFileName)
		
		hyperlinker="<a href='file:///"+savedFileName+"' title='点击打开'>"+savedFileName+"</a>"
		QMessageBox.information(self,"注意","处理完的广告主文件已经保存到<br/><br/>"+hyperlinker+"<br/>（点击链接打开文件）<br/><br/>")
		
		#关闭SQLITE数据库
		cursor.close()
		conn.close()
		
	#————————————————————————————————————————————————————
	#统计相关数据
	def statistics(self):
		print("doing statistics!")
		self.tableWidget.clear()
		self.tableWidget.setColumnCount(2)
		self.tableWidget.setRowCount(1000)
		self.tableWidget.setColumnWidth(0,300)
		self.tableWidget.setHorizontalHeaderLabels(["广告主","产品数"])		
		
		self.statusbar.showMessage("正在统计数据，请耐心等待...")
		#连接数据库
		conn = sqlite3.connect('advertisers.db')
		cursor = conn.cursor()
		
		#统计广告主总数
		sql="SELECT COUNT(DISTINCT cname) FROM advertiser"		
		cursor.execute(sql)	
		conn.commit
		
		advertiserCount=cursor.fetchone()[0]
		
		print("广告主总数：",advertiserCount)
		self.lcdNumber.display(advertiserCount)
		
		#统计产品总数
		sql="SELECT COUNT(DISTINCT bname) FROM advertiser"		
		cursor.execute(sql)	
		conn.commit
		
		productCount=cursor.fetchone()[0]
		
		print("产品总数：",productCount)
		self.lcdNumber_2.display(productCount)
		
		#分组统计各个广告主拥有的产品数并排序
		sql="SELECT cname, COUNT(bname) AS cnt FROM advertiser GROUP BY cname ORDER BY cnt DESC LIMIT 1000"
		cursor.execute(sql)	
		conn.commit
		
		rows=cursor.fetchall()
		
		if rows:
			rowNumber=0
			for row in rows:
				cell0 = QTableWidgetItem(row[0]) 
				cell1 = QTableWidgetItem(str(row[1])) 
				self.tableWidget.setItem(rowNumber, 0, cell0)
				self.tableWidget.setItem(rowNumber, 1, cell1) 
				rowNumber+=1	
			
		cursor.close()
		conn.close()
		
		self.statusbar.showMessage("统计已经完成，显示的是最新数据")
		QMessageBox.information(self,"统计完成","统计完成，请查阅数据！")
		
	
			
#————————————————————————————————————————————————————
app=QApplication(sys.argv)
w=MainWindow()
w.show()
sys.exit(app.exec())
