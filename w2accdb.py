#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pypyodbc
import logger

def insert(list):
	logger.loggerprint("写入access数据库，开始")
	db = 'Driver={Microsoft Access Database (*.mdb,*.accdb)};DBQ=A:\\来文来电\\2019.mdb'
	conn = pypyodbc.win_connect_mdb(db)
	curser = conn.cursor()
	sql_insert = '''INSERT INTO 2019(ID,来文单位,来文日期,来文名称,主要内容) VALUES(?,?,?,?,?)'''
	insert_value = (list[0], list[1],list[2],list[3],list[4])
	logger.loggerprint("sql语句：" + sql_insert + insert_value)
	#print(list)
	#(ID，来文单位，来文日期，来文名称，主要内容，备注)
	curser.execute(sql_insert,insert_value)
	#curser.execute("INSERT INTO 2019(ID,来文单位,来文日期,来文名称,主要内容) VALUES('100', 'asd', 'asd', 'asd', 'asd')")
	conn.commit()  # 没他不报错，也不插入，因为没有这条语句，折腾了一天，我日他娘。
	conn.close()
	logger.loggerprint("写入access数据库，完毕")
	#print('success')

#%(list[0],list[1],list[2],list[3],list[4])
#insert into 2019 values(45, 'asda', '2019年3月7日', 'asdasd', 'a', '');
#insert into 2019 (ID,来文单位,来文日期,来文名称,主要内容) values(66,'aa','2019年3月7日','aa','aa');



