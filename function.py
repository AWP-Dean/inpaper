#!/usr/bin/python
# -*- coding: UTF-8 -*-

import win32.win32print as win32print#本来是import win32print 和import win32api 在pycharm中虽然能运行，但一直报错，后来改成现在这样子，加了个路径，对了
import win32.win32api as win32api
import time
import logger
from PIL import Image

def resize(w_box, h_box, pil_image):  # 参数是：要适应的窗口宽、高、Image.open后的图片
	w, h = pil_image.size  # 获取图像的原始大小
	f1 = 1.0 * w_box / w
	f2 = 1.0 * h_box / h
	factor = min([f1, f2])
	width = int(w * factor)
	height = int(h * factor)
	return pil_image.resize((width, height), Image.ANTIALIAS)
"""
#str_1原始字符串
#index位置
#str插入的字符
def insert(str_1,index,str):
	str_list = list(str_1)
	#nPos = str_list.index(index) + 1
	str_list.insert(index + 1, str)
	str_2 = "".join(str_list)
	return str_2

def saveformat(list):
	#print(len(list[2]))
	#print(list[2].find('日')+1)
	if len(list[1]) > list[1].find('日') + 1:
		str = insert(list[1],list[1].find('日'),'\n')

	list[5] = insert(list[5],-1, '    ')
	list[5] = insert(list[5],list[5].find('\n'),'    ')

	#return [len(list[1]),list[1].find('日') + 1,'长',str,list[5]]
	return list[5]
"""




#打印功能
#Canon iR2018 UFRII LT
#print /d:USB001 c:\ODBC\1.docx，没用这个
def printer(file):
	logger.loggerprint("开始打印功能，打印的是：")
	logger.loggerprint(file)
	filename = file
	win32api.ShellExecute(
		0,
		"print",
		filename,
		'/d:"%s"' % win32print.GetDefaultPrinter(),
		".",
		0
	)

def currenttime():
	return time.strftime("%Y-%m-%d  %H:%M:%S",time.localtime())

def time1():
	return time.strftime("%Y-%m-%d",time.localtime())

