#!/usr/bin/python
# -*- coding: UTF-8 -*-
import function

def loggerprint(str):
	filepath = r'C:\Users\Administrator\Desktop\来文来电\\logger'
	filename = function.time1() + '.txt'
	log = open(filepath + filename, 'a',newline='')
	print(function.currenttime(), file=log)
	print('\r\n', file=log)
	print(str, file=log)
	print('\r\n',file=log)
	log.close()