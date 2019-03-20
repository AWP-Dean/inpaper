#!/usr/bin/python
# -*- coding: UTF-8 -*-
from tkinter import *
from tkinter import ttk
from PIL import Image,ImageTk
import tkinter.messagebox
import time
import rxls
import w2xls
import function
import w2accdb
import logger

class App(object):
	def __init__(self, object):

		v = IntVar()
		#已存信息部分
		self.xlsxFrame = LabelFrame(object, width=260, height=60, text='已存信息')
		xlsxlist = rxls.getsheets()
		xlsxCombobox = ttk.Combobox(self.xlsxFrame)
		xlsxCombobox["values"] = tuple(xlsxlist)
		xlsxCombobox.current(len(xlsxlist)-1)#默认显示最近一个
		#绑定选择后直接运行sort功能,好像不太对
		#xlsxCombobox.bind("<<ComboboxSelected>>",print('test'))
		xlsxCombobox["state"] = "readonly"
		xlsxCombobox.place(x = 5,y = 5)

		#来文信息部分
		self.inpaperFrame = LabelFrame(object, width=380, height=480, text='来文信息')
		self.inpaperFrame.place(x=0, y=60)

		#显示扫描件的地方
		#self.scanFrame = LabelFrame(object, width=390, height=480, text='图片信息')
		#self.scanFrame.place(x=385, y=60)
		im = Label(self.inpaperFrame,width = 360,height = 420)
		im.place(x=400, y=0)

		#来文单位
		aptnamelabel = Label(self.inpaperFrame,text='来文单位',height=3)
		aptnamelabel.place(x = 5,y = 0)
		aptnameEntry = Entry(self.inpaperFrame,width = 40)
		aptnameEntry.place(x=70, y=20)
		#来文时间
		inpapertime = Label(self.inpaperFrame,text='来文时间',height=3)
		inpapertime.place(x = 5,y = 40)
		timeEntry = Entry(self.inpaperFrame,width = 18)
		timeEntry.insert(0,"%d年%d月%d日"%(time.localtime().tm_year,\
										time.localtime().tm_mon,\
										time.localtime().tm_mday))
		timeEntry.place(x=70, y=60)
		minreq = Label(self.inpaperFrame,text='是否精确',height=3)
		minreq.place(x=210, y=43)


		#来文名称
		inpapername = Label(self.inpaperFrame, text='来文名称',height=3)
		inpapername.place(x=5, y=80)
		inpaperText = Text(self.inpaperFrame,height = 2,width=40)
		inpaperText.place(x= 70,y = 100)

		#开会时间、地点
		meettime = Label(self.inpaperFrame, text='开会时间',height=3)
		meettime.place(x= 5,y = 140)
		meettimeEntry = Entry(self.inpaperFrame)
		meettimeEntry.place(x= 70,y = 160)
		meetaddr = Label(self.inpaperFrame, text='开会地点',height=3)
		meetaddr.place(x= 5,y = 180)
		meetaddrEntry = Entry(self.inpaperFrame,width = 40)
		meetaddrEntry.place(x= 70,y = 200)

		#主要内容
		content = Label(self.inpaperFrame, text='主要内容',height=3)
		content.place(x= 5,y = 220)
		contentText = Text(self.inpaperFrame,height = 10,width = 40)
		contentText.place(x= 70,y = 240)

		#拟办意见
		plan = Label(self.inpaperFrame, text='拟办意见', height=3)
		plan.place(x= 5,y = 370)
		default_value = StringVar()
		default_value.set('请王蕾同志阅示。')
		planEntry = Entry(self.inpaperFrame,textvariable = default_value,width = 25)
		planEntry.place(x= 70,y = 390)

		# 拟办意见2
		planText = Text(self.inpaperFrame,height = 2,width=40)
		planText.place(x= 70,y = 420)
		planText.insert(1.0, '    拟请')



		#初始化填写界面
		def init():
			logger.loggerprint("进入初始化填写界面功能，init（）")
			aptnameEntry.delete(0, END)
			timeEntry.delete(0, END)
			inpaperText.delete(0.0, END)
			meettimeEntry.delete(0, END)
			meetaddrEntry.delete(0, END)
			contentText.delete(0.0, END)
			planEntry.delete(0, END)
			planText.delete(0.0, END)
			logger.loggerprint("初始化填写界面完毕")

		def initbox():
			logger.loggerprint("进入初始化下拉菜单功能，initbox（）")
			xlsxlist = rxls.getsheets()
			logger.loggerprint("进入初始化下拉菜单功能，initbox（）")
			xlsxCombobox["values"] = tuple(xlsxlist)
			xlsxCombobox.current(len(xlsxlist) - 1)
			xlsxCombobox["state"] = "readonly"
			xlsxCombobox.place(x = 5,y = 5)
			logger.loggerprint("重新初始化下拉菜单成功")

		def message(str):
			tkinter.messagebox.showinfo('提示', str)

		def scanview():
			logger.loggerprint('进入扫描件预览功能，scanview（）')
			im.destroy()
			scanfile = Image.open(r'C:\Users\Administrator\Desktop\HP0003.jpg')
			resized_scanfile = function.resize(360, 420, scanfile)
			show_scanfile = ImageTk.PhotoImage(resized_scanfile)
			img = Label(self.inpaperFrame, image=show_scanfile, width=360, height=420)
			img.image = show_scanfile  # 因为缺这么一个死活不出来，麻辣隔壁，上面那一句的image=img还得在，原理不知道，马勒戈壁
			img.place(x=400, y=0)
			logger.loggerprint('完成扫描件预览功能，scanview（）')

		def sort():
			logger.loggerprint('进入查询功能，sort（）')
			init()
			sortstr = xlsxCombobox.get()
			logger.loggerprint('下拉菜单选择的是：')
			logger.loggerprint(sortstr)
			#print('sort',sortstr)
			list = rxls.mess(sortstr)
			logger.loggerprint('查询到的信息是：')
			logger.loggerprint(list)
			scanview()
			if len(list) == 8:
				aptnameEntry.insert(0, list[0])
				timeEntry.insert(0, list[1])
				inpaperText.insert(INSERT, list[2])
				meettimeEntry.insert(0, list[3])
				meetaddrEntry.insert(0, list[4])
				contentText.insert(INSERT, list[5])
				planEntry.insert(0, list[6])
				planText.insert(INSERT, list[7])
				logger.loggerprint('有会议，插入成功，查询完毕')
			elif len(list) == 6:
				aptnameEntry.insert(0, list[0])
				timeEntry.insert(0, list[1])
				inpaperText.insert(INSERT, list[2])
				contentText.insert(INSERT, list[3])
				planEntry.insert(0, list[4])
				planText.insert(INSERT, list[5])
				logger.loggerprint('无会议，插入成功，查询完毕')

		def timecall():
			logger.loggerprint('进入是否精确功能，timecall（）')
			timeEntry.delete(0, END)
			if v.get() == 1:
				logger.loggerprint("是否精确，选择的‘是’")
				timeEntry.insert(0, "%d年%d月%d日\n%s" % (time.localtime().tm_year, \
												   time.localtime().tm_mon, \
												   time.localtime().tm_mday,\
												   time.strftime("%H:%M",time.localtime())))
				logger.loggerprint("timecall（）完毕")
			elif v.get() == 0:
				logger.loggerprint("是否精确，选择的‘否’")
				timeEntry.insert(0, "%d年%d月%d日" % (time.localtime().tm_year, \
												   time.localtime().tm_mon, \
												   time.localtime().tm_mday))
				logger.loggerprint("timecall（）完毕")
			else :
				message("我也不知道怎么会到这一步")

		def save():
			logger.loggerprint("进入保存功能，save（）")
			aptname = aptnameEntry.get()
			time = timeEntry.get()
			inpaper = inpaperText.get('1.0',END)[:-1]#这里加[:-1]是为了删除最后多出来的换行符
			meettime = meettimeEntry.get()
			meetaddr = meetaddrEntry.get()
			content = contentText.get('1.0',END)[:-1]
			plan = planEntry.get()
			plan1 = planText.get('1.0',END)[:-1]
			logger.loggerprint("保存前（判断前）获取信息如下：")
			logger.loggerprint([aptname,time,inpaper,meettime,meetaddr,content,plan,plan1])
			if aptname == "" and inpaper == "" and content == "" :
				message('请填写信息！！！')
			elif meettime == '':
				no = int(rxls.getno()) + 1
				list = [no,aptname,time,inpaper,content,plan,plan1]
				#print(no)
				logger.loggerprint("保存前（没会）获取信息如下：")
				logger.loggerprint(list)
				w2xls.withoutmeeting(list)
				w2xls.tempfilewithoutmeeting(list)
				w2accdb.insert(list)
				logger.loggerprint("保存完毕，重新初始化下拉框之前")
				initbox()  # 将保存好的sheet显示在下拉框里
				logger.loggerprint("保存完毕，重新初始化下拉框之后")
				message('保存成功')
			else:
				no = int(rxls.getno()) + 1
				list = [no, aptname, time, inpaper, meettime, meetaddr, content, plan, plan1]
				# print(no)
				logger.loggerprint("保存前（有会）获取信息如下：")
				logger.loggerprint(list)
				w2xls.withmeeting(list)
				w2xls.tempfilewithmeeting(list)
				w2accdb.insert(list)
				logger.loggerprint("保存完毕，重新初始化下拉框之前")
				initbox()  # 将保存好的sheet显示在下拉框里
				logger.loggerprint("保存完毕，重新初始化下拉框之后")
				message('保存成功')



			"""
			if meettime == '':
				list = [aptname,time,inpaper,content,plan,plan1]
				no = int(rxls.getno()) + 1
				#print(no)
				w2xls.withoutmeeting(no,list)
			else:
				list = [aptname,time,inpaper,meettime,meetaddr,content,plan,plan1]
				no = int(rxls.getno()) + 1
				#print(no)
				w2xls.withmeeting(no, list)
			"""

		def saveandprint():
			logger.loggerprint("进入保存并且打印功能")
			save()
			printxls()
			logger.loggerprint("完成保存并且打印功能")

		def printxls():
			logger.loggerprint("进入打印功能，printxls（）")
			logger.loggerprint("进入打印功能，下拉菜单选项，确定打印哪个")
			sortstr = xlsxCombobox.get()
			logger.loggerprint("打印的是：" + sortstr)
			list = [sortstr[5:]] + rxls.mess(sortstr)
			logger.loggerprint("内容是：")
			logger.loggerprint(list)
			if len(list) == 9:
				w2xls.tempfilewithmeeting(list)
				logger.loggerprint("创建有会议的tempfile成功")
			elif len(list) == 7:
				w2xls.tempfilewithoutmeeting(list)
				logger.loggerprint("创建无会议的tempfile成功")
			function.printer('C:\\tempfile.xls')
			logger.loggerprint("打印完毕")
			message('打印成功')

		def shutdown():
			logger.loggerprint("--------关闭界面--------")
			root.destroy()

		#时间是否精确
		minreqbutton1 = Radiobutton(self.inpaperFrame, text='是', value=1,command=timecall, variable=v)
		minreqbutton2 = Radiobutton(self.inpaperFrame, text='否', value=0,command=timecall, variable=v)
		minreqbutton1.place(x=260, y=60)
		minreqbutton2.place(x=330, y=60)
		minreqbutton2.select()

		#查询按钮
		checkbutton = Button(self.xlsxFrame, text='查询(Q)',command=sort)
		checkbutton.place(x = 180,y = 0)
		self.xlsxFrame.place(x = 270,y = 0)
		# 按钮
		self.action1 = ttk.Button(object, text="保存(S)", command=save)  # 创建一个按钮, text
		self.action1.place(x=70, y=550)
		# 按钮
		self.action2 = ttk.Button(object, text="保存并打印(S&P)", command=saveandprint)  # 创建一个按钮, text
		self.action2.place(x=230, y=550)
		# 按钮
		self.action3 = ttk.Button(object, text="打印(P)", command=printxls)  # 创建一个按钮, text
		self.action3.place(x=400, y=550)
		# 按钮
		self.action4 = ttk.Button(object, text="关闭(C)", command=shutdown)  # 创建一个按钮, text
		self.action4.place(x=560, y=550)


root = Tk()
root.title("来文阅办信息化管理系统")
root.geometry("800x600")
root.geometry("+500+200")
app = App(root)
logger.loggerprint("------------开始-----------")
root.mainloop()
