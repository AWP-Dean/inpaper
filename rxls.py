#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlrd
import logger

def mess(str):
	logger.loggerprint("开始读取xls，获取信息")
	wb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\来文来电\2019.xls")
	ws = wb.sheet_by_name(str)
	#print(wb.sheet_names())
	list = []
	if ws.cell_value(4, 0) == '会议时间':
		list = [ws.cell_value(2, 1),\
				ws.cell_value(2, 3),\
				ws.cell_value(3, 1),\
				ws.cell_value(4, 1),\
				ws.cell_value(4, 3),\
				ws.cell_value(5, 1),\
				ws.cell_value(7, 1),\
				ws.cell_value(8, 1)]
		logger.loggerprint("确定有'会议时间'，获取到的信息是：")
		logger.loggerprint(list)
	else:
		list = [ws.cell_value(2, 1), \
				ws.cell_value(2, 3), \
				ws.cell_value(3, 1), \
				ws.cell_value(4, 1), \
				ws.cell_value(6, 1), \
				ws.cell_value(7, 1)]
		logger.loggerprint("确定无'会议时间'，获取到的信息是：")
		logger.loggerprint(list)
	return list

def getno():
	#wb = xlrd.open_workbook("C:\\Users\\Administrator\\Desktop\\相关工作\\来文来电\\2019.xls")
	#list = wb.sheet_names()

	list = getsheets()
	str = list[-1]
	logger.loggerprint("获取到的收文号，getno（）是：", + str[str.find('-')+1:])
	#print(str[str.find('-')+1:])
	return str[str.find('-')+1:]

def getsheets():

	wb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\来文来电\2019.xls")
	list = wb.sheet_names()
	logger.loggerprint("获取sheet列表，getsheets（）是：")
	logger.loggerprint(list)
	return list


"""
	备注：
	有会议：
	收文号:	ws.cell_value(0,3)
	来文单位：ws.cell_value(2,1)
	收文时间：ws.cell_value(2,3)
	来文名称：ws.cell_value(3,1)
	会议时间：ws.cell_value(4,1)
	会议地点：ws.cell_value(4,3)
	主要内容：ws.cell_value(5,1)
	拟办意见1：ws.cell_value(7,1)
	拟办意见2：ws.cell_value(8,1)
	
	无会议：
	收文号:	ws.cell_value(0,3)
	来文单位：ws.cell_value(2,1)
	收文时间：ws.cell_value(2,3)
	来文名称：ws.cell_value(3,1)
	主要内容：ws.cell_value(4,1)
	拟办意见1：ws.cell_value(6,1)
	拟办意见2：ws.cell_value(7,1)
	
	
"""
