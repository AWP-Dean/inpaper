#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlwt
import xlrd
from xlutils.copy import copy
import logger


def withmeeting(list):
	logger.loggerprint("有会议，写入xls，开始")
	rb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\来文来电\2019.xls", formatting_info=True)
	wb = copy(rb)
	ws = wb.add_sheet('2019-' + str(list[0]))

	# 设置字体，字号（仿宋_GB2312，14）song_font
	song_font = xlwt.Font()
	song_font.name = '仿宋_GB2312'
	song_font.height = 280  # 字体大小，220就是11号字体，大概就是11*20得来的吧

	# 设置字体，字号（黑体，14）black_font
	black_font = xlwt.Font()
	black_font.name = '黑体'
	black_font.height = 280

	# 设置字体，字号（方正小标宋_GBK，20）black_font1
	black_font1 = xlwt.Font()
	black_font1.name = '方正小标宋_GBK'
	black_font1.height = 400

	# 设置边框（上下左右）
	all_border = xlwt.Borders()  # 给单元格加框线
	all_border.left = xlwt.Borders.THIN  # 左
	all_border.top = xlwt.Borders.THIN  # 上
	all_border.right = xlwt.Borders.THIN  # 右
	all_border.bottom = xlwt.Borders.THIN  # 下
	all_border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	all_border.right_colour = 0x40
	all_border.top_colour = 0x40
	all_border.bottom_colour = 0x40

	# 设置边框（上下右）
	udr_border = xlwt.Borders()  # 给单元格加框线
	udr_border.top = xlwt.Borders.THIN  # 上
	udr_border.right = xlwt.Borders.THIN  # 右
	udr_border.bottom = xlwt.Borders.THIN  # 下
	# 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	udr_border.right_colour = 0x40
	udr_border.top_colour = 0x40
	udr_border.bottom_colour = 0x40

	# 设置边框（上）
	u_border = xlwt.Borders()  # 给单元格加框线
	u_border.top = xlwt.Borders.THIN  # 上
	u_border.top_colour = 0x40

	# 设置边框（上下）
	ud_border = xlwt.Borders()  # 给单元格加框线
	ud_border.top = xlwt.Borders.THIN  # 上
	ud_border.bottom = xlwt.Borders.THIN  # 下
	ud_border.top_colour = 0x40  # 上
	ud_border.bottom_colour = 0x40  # 下

	# 自动换行，水平垂直居中
	wrap_mid_al = xlwt.Alignment()
	wrap_mid_al.horz = 0x02  # 设置水平居中
	wrap_mid_al.vert = 0x01  # 设置垂直居中
	wrap_mid_al.wrap = 1

	# 自动换行，水平居左，垂直居中
	wrap_left_al = xlwt.Alignment()
	wrap_left_al.horz = 0x01  # 设置水平居左
	wrap_left_al.vert = 0x01  # 设置垂直居中
	wrap_left_al.wrap = 1

	# 不换行，水平居左，垂直居中
	nowrap_left_al = xlwt.Alignment()
	nowrap_left_al.horz = 0x01  # 设置水平居左
	nowrap_left_al.vert = 0x01  # 设置垂直居中

	# 黑体14号字，水平垂直居中，自动换行，上下右边框
	# 来文单位，来文名称，主要内容，领导批示，拟办意见，处理结果
	black_mid_wrap_udr_style = xlwt.XFStyle()
	black_mid_wrap_udr_style.alignment = wrap_mid_al
	black_mid_wrap_udr_style.font = black_font
	black_mid_wrap_udr_style.borders = udr_border

	# 黑体14号字，水平垂直居中，自动换行，全边框
	# 收文时间，会议地点
	black_mid_wrap_all_style = xlwt.XFStyle()
	black_mid_wrap_all_style.alignment = wrap_mid_al
	black_mid_wrap_all_style.font = black_font
	black_mid_wrap_all_style.borders = all_border

	# 仿宋GB2312,14号字，水平垂直居中，自动换行，仅上框
	# 来文单位，来文名称，领导批示，收文时间，会议地点
	# 上面的那些内容
	song_mid_wrap_u_style = xlwt.XFStyle()
	song_mid_wrap_u_style.alignment = wrap_mid_al
	song_mid_wrap_u_style.font = song_font
	song_mid_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，仅上框
	# 主要内容，拟办意见
	# 上面的那些内容
	song_left_wrap_u_style = xlwt.XFStyle()
	song_left_wrap_u_style.alignment = wrap_left_al
	song_left_wrap_u_style.font = song_font
	song_left_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，上下框
	# 处理结果
	# 内容
	song_left_wrap_ud_style = xlwt.XFStyle()
	song_left_wrap_ud_style.alignment = wrap_left_al
	song_left_wrap_ud_style.font = song_font
	song_left_wrap_ud_style.borders = ud_border

	# 黑体14号，垂直居中，水平居左，不换行，无框
	# 紧急程度，省总工会办公室，承办人：张腾
	black_left_nowrap_nobord_style = xlwt.XFStyle()
	black_left_nowrap_nobord_style.alignment = nowrap_left_al
	black_left_nowrap_nobord_style.font = black_font

	# 宋体14号，垂直居中，水平居左，换行，无框
	# 收文号
	song_left_nowrap_nobord_style = xlwt.XFStyle()
	song_left_nowrap_nobord_style.alignment = wrap_left_al
	song_left_nowrap_nobord_style.font = song_font

	# 方正小标宋_GBK 20号，水平垂直居中，换行，无框
	# 山西省总工会来文、来电阅办卡片
	black_mid_wrap_nobord_style = xlwt.XFStyle()
	black_mid_wrap_nobord_style.alignment = wrap_mid_al
	black_mid_wrap_nobord_style.font = black_font1

	# 设置单元格宽度
	ws.col(0).width = 3500
	ws.col(1).width = 8000
	ws.col(2).width = 4500
	ws.col(3).width = 8000
	ws.row(0).set_style(xlwt.easyxf('font:height 600;'))  # 行高37.5
	ws.row(1).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(2).set_style(xlwt.easyxf('font:height 650;'))
	ws.row(3).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(4).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(5).set_style(xlwt.easyxf('font:height 3000;'))  # 行高180
	ws.row(6).set_style(xlwt.easyxf('font:height 3000;'))  # 行高180
	ws.row(7).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(8).set_style(xlwt.easyxf('font:height 780;'))  # 行高50
	ws.row(9).set_style(xlwt.easyxf('font:height 480;'))  # 行高30
	ws.row(10).set_style(xlwt.easyxf('font:height 570;'))  # 行高30
	ws.row(11).set_style(xlwt.easyxf('font:height 300;'))  # 行高18

	ws.write(0, 0, '紧急程度:', black_left_nowrap_nobord_style)
	ws.write(0, 3, '       收文号：[2019]'+ str(list[0]), song_left_nowrap_nobord_style)
	ws.write_merge(1, 1, 0, 3, '山西省总工会来文、来电阅办卡片', black_mid_wrap_nobord_style)
	ws.write(2, 0, '来文单位', black_mid_wrap_udr_style)
	ws.write(2, 1, list[1], song_mid_wrap_u_style)
	ws.write(2, 2, '收文时间', black_mid_wrap_all_style)
	ws.write(2, 3, list[2], song_mid_wrap_u_style)
	ws.write(3, 0, '来文名称', black_mid_wrap_udr_style)
	ws.write_merge(3, 3, 1, 3, list[3], song_mid_wrap_u_style)
	ws.write(4, 0, '会议时间', black_mid_wrap_udr_style)
	ws.write(4, 1, list[4], song_mid_wrap_u_style)
	ws.write(4, 2, '会议地点', black_mid_wrap_all_style)
	ws.write(4, 3, list[5], song_mid_wrap_u_style)
	ws.write(5, 0, '主要内容', black_mid_wrap_udr_style)
	ws.write_merge(5, 5, 1, 3, list[6], song_left_wrap_u_style)
	ws.write(6, 0, '领\n导\n批\n示', black_mid_wrap_udr_style)
	ws.write_merge(6, 6, 1, 3, '', song_left_wrap_u_style)
	ws.write_merge(7, 9, 0, 0, '拟办意见', black_mid_wrap_udr_style)
	ws.write_merge(7, 7, 1, 3, list[7], song_left_wrap_u_style)
	ws.write_merge(8, 8, 1, 3, list[8], song_left_nowrap_nobord_style)
	ws.write(10, 0, '处理结果', black_mid_wrap_udr_style)
	ws.write_merge(10, 10, 1, 3, '', song_left_wrap_ud_style)
	ws.write(11, 0, ' 省总工会办公室', black_left_nowrap_nobord_style)  # 5个空格
	ws.write(11, 3, '        承办人：张  腾', black_left_nowrap_nobord_style)

	wb.save(r"C:\Users\Administrator\Desktop\来文来电\2019.xls")
	logger.loggerprint("有会议，写入xls，完毕")

def withoutmeeting(list):

	logger.loggerprint("无会议，写入xls，开始")
	rb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\来文来电\2019.xls", formatting_info=True)
	wb = copy(rb)
	ws = wb.add_sheet('2019-' + str(list[0]))

	# 设置字体，字号（仿宋_GB2312，14）song_font
	song_font = xlwt.Font()
	song_font.name = '仿宋_GB2312'
	song_font.height = 280  # 字体大小，220就是11号字体，大概就是11*20得来的吧
	# 设置字体，字号（黑体，14）black_font
	black_font = xlwt.Font()
	black_font.name = '黑体'
	black_font.height = 280
	# 设置字体，字号（方正小标宋_GBK，20）black_font1
	black_font1 = xlwt.Font()
	black_font1.name = '方正小标宋_GBK'
	black_font1.height = 400

	# 设置边框（上下左右）
	all_border = xlwt.Borders()  # 给单元格加框线
	all_border.left = xlwt.Borders.THIN  # 左
	all_border.top = xlwt.Borders.THIN  # 上
	all_border.right = xlwt.Borders.THIN  # 右
	all_border.bottom = xlwt.Borders.THIN  # 下
	all_border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	all_border.right_colour = 0x40
	all_border.top_colour = 0x40
	all_border.bottom_colour = 0x40

	# 设置边框（上下右）
	udr_border = xlwt.Borders()  # 给单元格加框线
	udr_border.top = xlwt.Borders.THIN  # 上
	udr_border.right = xlwt.Borders.THIN  # 右
	udr_border.bottom = xlwt.Borders.THIN  # 下
	# 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	udr_border.right_colour = 0x40
	udr_border.top_colour = 0x40
	udr_border.bottom_colour = 0x40

	# 设置边框（上）
	u_border = xlwt.Borders()  # 给单元格加框线
	u_border.top = xlwt.Borders.THIN  # 上
	u_border.top_colour = 0x40

	# 设置边框（上下）
	ud_border = xlwt.Borders()  # 给单元格加框线
	ud_border.top = xlwt.Borders.THIN  # 上
	ud_border.bottom = xlwt.Borders.THIN  # 下
	ud_border.top_colour = 0x40  # 上
	ud_border.bottom_colour = 0x40  # 下

	# 自动换行，水平垂直居中
	wrap_mid_al = xlwt.Alignment()
	wrap_mid_al.horz = 0x02  # 设置水平居中
	wrap_mid_al.vert = 0x01  # 设置垂直居中
	wrap_mid_al.wrap = 1

	# 自动换行，水平居左，垂直居中
	wrap_left_al = xlwt.Alignment()
	wrap_left_al.horz = 0x01  # 设置水平居左
	wrap_left_al.vert = 0x01  # 设置垂直居中
	wrap_left_al.wrap = 1

	# 不换行，水平居左，垂直居中
	nowrap_left_al = xlwt.Alignment()
	nowrap_left_al.horz = 0x01  # 设置水平居左
	nowrap_left_al.vert = 0x01  # 设置垂直居中

	# 黑体14号字，水平垂直居中，自动换行，上下右边框
	# 来文单位，来文名称，主要内容，领导批示，拟办意见，处理结果
	black_mid_wrap_udr_style = xlwt.XFStyle()
	black_mid_wrap_udr_style.alignment = wrap_mid_al
	black_mid_wrap_udr_style.font = black_font
	black_mid_wrap_udr_style.borders = udr_border

	# 黑体14号字，水平垂直居中，自动换行，全边框
	# 收文时间，会议地点
	black_mid_wrap_all_style = xlwt.XFStyle()
	black_mid_wrap_all_style.alignment = wrap_mid_al
	black_mid_wrap_all_style.font = black_font
	black_mid_wrap_all_style.borders = all_border

	# 仿宋GB2312,14号字，水平垂直居中，自动换行，仅上框
	# 来文单位，来文名称，领导批示，收文时间，会议地点
	# 上面的那些内容
	song_mid_wrap_u_style = xlwt.XFStyle()
	song_mid_wrap_u_style.alignment = wrap_mid_al
	song_mid_wrap_u_style.font = song_font
	song_mid_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，仅上框
	# 主要内容，拟办意见
	# 上面的那些内容
	song_left_wrap_u_style = xlwt.XFStyle()
	song_left_wrap_u_style.alignment = wrap_left_al
	song_left_wrap_u_style.font = song_font
	song_left_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，上下框
	# 处理结果
	# 内容
	song_left_wrap_ud_style = xlwt.XFStyle()
	song_left_wrap_ud_style.alignment = wrap_left_al
	song_left_wrap_ud_style.font = song_font
	song_left_wrap_ud_style.borders = ud_border

	# 黑体14号，垂直居中，水平居左，不换行，无框
	# 紧急程度，省总工会办公室，承办人：张腾
	black_left_nowrap_nobord_style = xlwt.XFStyle()
	black_left_nowrap_nobord_style.alignment = nowrap_left_al
	black_left_nowrap_nobord_style.font = black_font

	# 宋体14号，垂直居中，水平居左，换行，无框
	# 收文号
	song_left_nowrap_nobord_style = xlwt.XFStyle()
	song_left_nowrap_nobord_style.alignment = wrap_left_al
	song_left_nowrap_nobord_style.font = song_font

	# 方正小标宋_GBK 20号，水平垂直居中，换行，无框
	# 山西省总工会来文、来电阅办卡片
	black_mid_wrap_nobord_style = xlwt.XFStyle()
	black_mid_wrap_nobord_style.alignment = wrap_mid_al
	black_mid_wrap_nobord_style.font = black_font1

	# 设置单元格宽度
	ws.col(0).width = 3500
	ws.col(1).width = 8000
	ws.col(2).width = 4500
	ws.col(3).width = 8000
	ws.row(0).set_style(xlwt.easyxf('font:height 600;'))  # 行高37.5
	ws.row(1).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(2).set_style(xlwt.easyxf('font:height 650;'))
	ws.row(3).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(4).set_style(xlwt.easyxf('font:height 3000;'))
	ws.row(5).set_style(xlwt.easyxf('font:height 3900;'))
	ws.row(6).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(7).set_style(xlwt.easyxf('font:height 780;'))  # 行高50
	ws.row(8).set_style(xlwt.easyxf('font:height 480;'))  # 行高30
	ws.row(9).set_style(xlwt.easyxf('font:height 570;'))  # 行高30
	ws.row(10).set_style(xlwt.easyxf('font:height 300;'))  # 行高18

	ws.write(0, 0, '紧急程度:', black_left_nowrap_nobord_style)
	ws.write(0, 3, '       收文号：[2019]' + str(list[0]), song_left_nowrap_nobord_style)
	ws.write_merge(1, 1, 0, 3, '山西省总工会来文、来电阅办卡片', black_mid_wrap_nobord_style)
	ws.write(2, 0, '来文单位', black_mid_wrap_udr_style)
	ws.write(2, 1, list[1], song_mid_wrap_u_style)
	ws.write(2, 2, '收文时间', black_mid_wrap_all_style)
	ws.write(2, 3, list[2], song_mid_wrap_u_style)
	ws.write(3, 0, '来文名称', black_mid_wrap_udr_style)
	ws.write_merge(3, 3, 1, 3, list[3], song_mid_wrap_u_style)
	ws.write(4, 0, '主要内容', black_mid_wrap_udr_style)
	ws.write_merge(4, 4, 1, 3, list[4], song_left_wrap_u_style)
	ws.write(5, 0, '领\n导\n批\n示', black_mid_wrap_udr_style)
	ws.write_merge(5, 5, 1, 3, '', song_left_wrap_u_style)
	ws.write_merge(6, 8, 0, 0, '拟办意见', black_mid_wrap_udr_style)
	ws.write_merge(6, 6, 1, 3, list[5], song_left_wrap_u_style)
	ws.write_merge(7, 7, 1, 3, list[6], song_left_nowrap_nobord_style)
	ws.write(9, 0, '处理结果', black_mid_wrap_udr_style)
	ws.write_merge(9, 9, 1, 3, '', song_left_wrap_ud_style)
	ws.write(10, 0, ' 省总工会办公室', black_left_nowrap_nobord_style)  # 5个空格
	ws.write(10, 3, '        承办人：张  腾', black_left_nowrap_nobord_style)

	wb.save(r"C:\Users\Administrator\Desktop\来文来电\2019.xls")
	logger.loggerprint("无会议，写入xls，完毕")

def tempfilewithmeeting(list):
	logger.loggerprint("有会议，写入临时xls，开始")
	wb = xlwt.Workbook()
	ws = wb.add_sheet('tempfile')

# 设置字体，字号（仿宋_GB2312，14）song_font
	song_font = xlwt.Font()
	song_font.name = '仿宋_GB2312'
	song_font.height = 280  # 字体大小，220就是11号字体，大概就是11*20得来的吧

	# 设置字体，字号（黑体，14）black_font
	black_font = xlwt.Font()
	black_font.name = '黑体'
	black_font.height = 280

	# 设置字体，字号（方正小标宋_GBK，20）black_font1
	black_font1 = xlwt.Font()
	black_font1.name = '方正小标宋_GBK'
	black_font1.height = 400

	# 设置边框（上下左右）
	all_border = xlwt.Borders()  # 给单元格加框线
	all_border.left = xlwt.Borders.THIN  # 左
	all_border.top = xlwt.Borders.THIN  # 上
	all_border.right = xlwt.Borders.THIN  # 右
	all_border.bottom = xlwt.Borders.THIN  # 下
	all_border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	all_border.right_colour = 0x40
	all_border.top_colour = 0x40
	all_border.bottom_colour = 0x40

	# 设置边框（上下右）
	udr_border = xlwt.Borders()  # 给单元格加框线
	udr_border.top = xlwt.Borders.THIN  # 上
	udr_border.right = xlwt.Borders.THIN  # 右
	udr_border.bottom = xlwt.Borders.THIN  # 下
	# 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	udr_border.right_colour = 0x40
	udr_border.top_colour = 0x40
	udr_border.bottom_colour = 0x40

	# 设置边框（上）
	u_border = xlwt.Borders()  # 给单元格加框线
	u_border.top = xlwt.Borders.THIN  # 上
	u_border.top_colour = 0x40

	# 设置边框（上下）
	ud_border = xlwt.Borders()  # 给单元格加框线
	ud_border.top = xlwt.Borders.THIN  # 上
	ud_border.bottom = xlwt.Borders.THIN  # 下
	ud_border.top_colour = 0x40  # 上
	ud_border.bottom_colour = 0x40  # 下

	# 自动换行，水平垂直居中
	wrap_mid_al = xlwt.Alignment()
	wrap_mid_al.horz = 0x02  # 设置水平居中
	wrap_mid_al.vert = 0x01  # 设置垂直居中
	wrap_mid_al.wrap = 1

	# 自动换行，水平居左，垂直居中
	wrap_left_al = xlwt.Alignment()
	wrap_left_al.horz = 0x01  # 设置水平居左
	wrap_left_al.vert = 0x01  # 设置垂直居中
	wrap_left_al.wrap = 1

	# 不换行，水平居左，垂直居中
	nowrap_left_al = xlwt.Alignment()
	nowrap_left_al.horz = 0x01  # 设置水平居左
	nowrap_left_al.vert = 0x01  # 设置垂直居中

	# 黑体14号字，水平垂直居中，自动换行，上下右边框
	# 来文单位，来文名称，主要内容，领导批示，拟办意见，处理结果
	black_mid_wrap_udr_style = xlwt.XFStyle()
	black_mid_wrap_udr_style.alignment = wrap_mid_al
	black_mid_wrap_udr_style.font = black_font
	black_mid_wrap_udr_style.borders = udr_border

	# 黑体14号字，水平垂直居中，自动换行，全边框
	# 收文时间，会议地点
	black_mid_wrap_all_style = xlwt.XFStyle()
	black_mid_wrap_all_style.alignment = wrap_mid_al
	black_mid_wrap_all_style.font = black_font
	black_mid_wrap_all_style.borders = all_border

	# 仿宋GB2312,14号字，水平垂直居中，自动换行，仅上框
	# 来文单位，来文名称，领导批示，收文时间，会议地点
	# 上面的那些内容
	song_mid_wrap_u_style = xlwt.XFStyle()
	song_mid_wrap_u_style.alignment = wrap_mid_al
	song_mid_wrap_u_style.font = song_font
	song_mid_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，仅上框
	# 主要内容，拟办意见
	# 上面的那些内容
	song_left_wrap_u_style = xlwt.XFStyle()
	song_left_wrap_u_style.alignment = wrap_left_al
	song_left_wrap_u_style.font = song_font
	song_left_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，上下框
	# 处理结果
	# 内容
	song_left_wrap_ud_style = xlwt.XFStyle()
	song_left_wrap_ud_style.alignment = wrap_left_al
	song_left_wrap_ud_style.font = song_font
	song_left_wrap_ud_style.borders = ud_border

	# 黑体14号，垂直居中，水平居左，不换行，无框
	# 紧急程度，省总工会办公室，承办人：张腾
	black_left_nowrap_nobord_style = xlwt.XFStyle()
	black_left_nowrap_nobord_style.alignment = nowrap_left_al
	black_left_nowrap_nobord_style.font = black_font

	# 宋体14号，垂直居中，水平居左，换行，无框
	# 收文号
	song_left_nowrap_nobord_style = xlwt.XFStyle()
	song_left_nowrap_nobord_style.alignment = wrap_left_al
	song_left_nowrap_nobord_style.font = song_font

	# 方正小标宋_GBK 20号，水平垂直居中，换行，无框
	# 山西省总工会来文、来电阅办卡片
	black_mid_wrap_nobord_style = xlwt.XFStyle()
	black_mid_wrap_nobord_style.alignment = wrap_mid_al
	black_mid_wrap_nobord_style.font = black_font1

	# 设置单元格宽度
	ws.col(0).width = 3500
	ws.col(1).width = 8000
	ws.col(2).width = 4500
	ws.col(3).width = 8000
	ws.row(0).set_style(xlwt.easyxf('font:height 600;'))  # 行高37.5
	ws.row(1).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(2).set_style(xlwt.easyxf('font:height 650;'))
	ws.row(3).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(4).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(5).set_style(xlwt.easyxf('font:height 3000;'))  # 行高180
	ws.row(6).set_style(xlwt.easyxf('font:height 3000;'))  # 行高180
	ws.row(7).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(8).set_style(xlwt.easyxf('font:height 780;'))  # 行高50
	ws.row(9).set_style(xlwt.easyxf('font:height 480;'))  # 行高30
	ws.row(10).set_style(xlwt.easyxf('font:height 570;'))  # 行高30
	ws.row(11).set_style(xlwt.easyxf('font:height 300;'))  # 行高18

	ws.write(0, 0, '紧急程度:', black_left_nowrap_nobord_style)
	ws.write(0, 3, '       收文号：[2019]'+ str(list[0]), song_left_nowrap_nobord_style)
	ws.write_merge(1, 1, 0, 3, '山西省总工会来文、来电阅办卡片', black_mid_wrap_nobord_style)
	ws.write(2, 0, '来文单位', black_mid_wrap_udr_style)
	ws.write(2, 1, list[1], song_mid_wrap_u_style)
	ws.write(2, 2, '收文时间', black_mid_wrap_all_style)
	ws.write(2, 3, list[2], song_mid_wrap_u_style)
	ws.write(3, 0, '来文名称', black_mid_wrap_udr_style)
	ws.write_merge(3, 3, 1, 3, list[3], song_mid_wrap_u_style)
	ws.write(4, 0, '会议时间', black_mid_wrap_udr_style)
	ws.write(4, 1, list[4], song_mid_wrap_u_style)
	ws.write(4, 2, '会议地点', black_mid_wrap_all_style)
	ws.write(4, 3, list[5], song_mid_wrap_u_style)
	ws.write(5, 0, '主要内容', black_mid_wrap_udr_style)
	ws.write_merge(5, 5, 1, 3, list[6], song_left_wrap_u_style)
	ws.write(6, 0, '领\n导\n批\n示', black_mid_wrap_udr_style)
	ws.write_merge(6, 6, 1, 3, '', song_left_wrap_u_style)
	ws.write_merge(7, 9, 0, 0, '拟办意见', black_mid_wrap_udr_style)
	ws.write_merge(7, 7, 1, 3, list[7], song_left_wrap_u_style)
	ws.write_merge(8, 8, 1, 3, list[8], song_left_nowrap_nobord_style)
	ws.write(10, 0, '处理结果', black_mid_wrap_udr_style)
	ws.write_merge(10, 10, 1, 3, '', song_left_wrap_ud_style)
	ws.write(11, 0, ' 省总工会办公室', black_left_nowrap_nobord_style)  # 5个空格
	ws.write(11, 3, '        承办人：张  腾', black_left_nowrap_nobord_style)

	wb.save('C:\\tempfile.xls')
	logger.loggerprint("有会议，写入临时xls，完毕")

def tempfilewithoutmeeting(list):
	logger.loggerprint("无会议，写入临时xls，开始")
	wb = xlwt.Workbook()
	ws = wb.add_sheet('tempfile')

	# 设置字体，字号（仿宋_GB2312，14）song_font
	song_font = xlwt.Font()
	song_font.name = '仿宋_GB2312'
	song_font.height = 280  # 字体大小，220就是11号字体，大概就是11*20得来的吧
	# 设置字体，字号（黑体，14）black_font
	black_font = xlwt.Font()
	black_font.name = '黑体'
	black_font.height = 280
	# 设置字体，字号（方正小标宋_GBK，20）black_font1
	black_font1 = xlwt.Font()
	black_font1.name = '方正小标宋_GBK'
	black_font1.height = 400

	# 设置边框（上下左右）
	all_border = xlwt.Borders()  # 给单元格加框线
	all_border.left = xlwt.Borders.THIN  # 左
	all_border.top = xlwt.Borders.THIN  # 上
	all_border.right = xlwt.Borders.THIN  # 右
	all_border.bottom = xlwt.Borders.THIN  # 下
	all_border.left_colour = 0x40  # 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	all_border.right_colour = 0x40
	all_border.top_colour = 0x40
	all_border.bottom_colour = 0x40

	# 设置边框（上下右）
	udr_border = xlwt.Borders()  # 给单元格加框线
	udr_border.top = xlwt.Borders.THIN  # 上
	udr_border.right = xlwt.Borders.THIN  # 右
	udr_border.bottom = xlwt.Borders.THIN  # 下
	# 设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
	udr_border.right_colour = 0x40
	udr_border.top_colour = 0x40
	udr_border.bottom_colour = 0x40

	# 设置边框（上）
	u_border = xlwt.Borders()  # 给单元格加框线
	u_border.top = xlwt.Borders.THIN  # 上
	u_border.top_colour = 0x40

	# 设置边框（上下）
	ud_border = xlwt.Borders()  # 给单元格加框线
	ud_border.top = xlwt.Borders.THIN  # 上
	ud_border.bottom = xlwt.Borders.THIN  # 下
	ud_border.top_colour = 0x40  # 上
	ud_border.bottom_colour = 0x40  # 下

	# 自动换行，水平垂直居中
	wrap_mid_al = xlwt.Alignment()
	wrap_mid_al.horz = 0x02  # 设置水平居中
	wrap_mid_al.vert = 0x01  # 设置垂直居中
	wrap_mid_al.wrap = 1

	# 自动换行，水平居左，垂直居中
	wrap_left_al = xlwt.Alignment()
	wrap_left_al.horz = 0x01  # 设置水平居左
	wrap_left_al.vert = 0x01  # 设置垂直居中
	wrap_left_al.wrap = 1

	# 不换行，水平居左，垂直居中
	nowrap_left_al = xlwt.Alignment()
	nowrap_left_al.horz = 0x01  # 设置水平居左
	nowrap_left_al.vert = 0x01  # 设置垂直居中

	# 黑体14号字，水平垂直居中，自动换行，上下右边框
	# 来文单位，来文名称，主要内容，领导批示，拟办意见，处理结果
	black_mid_wrap_udr_style = xlwt.XFStyle()
	black_mid_wrap_udr_style.alignment = wrap_mid_al
	black_mid_wrap_udr_style.font = black_font
	black_mid_wrap_udr_style.borders = udr_border

	# 黑体14号字，水平垂直居中，自动换行，全边框
	# 收文时间，会议地点
	black_mid_wrap_all_style = xlwt.XFStyle()
	black_mid_wrap_all_style.alignment = wrap_mid_al
	black_mid_wrap_all_style.font = black_font
	black_mid_wrap_all_style.borders = all_border

	# 仿宋GB2312,14号字，水平垂直居中，自动换行，仅上框
	# 来文单位，来文名称，领导批示，收文时间，会议地点
	# 上面的那些内容
	song_mid_wrap_u_style = xlwt.XFStyle()
	song_mid_wrap_u_style.alignment = wrap_mid_al
	song_mid_wrap_u_style.font = song_font
	song_mid_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，仅上框
	# 主要内容，拟办意见
	# 上面的那些内容
	song_left_wrap_u_style = xlwt.XFStyle()
	song_left_wrap_u_style.alignment = wrap_left_al
	song_left_wrap_u_style.font = song_font
	song_left_wrap_u_style.borders = u_border

	# 仿宋GB2312,14号字，垂直居中，水平居左，自动换行，上下框
	# 处理结果
	# 内容
	song_left_wrap_ud_style = xlwt.XFStyle()
	song_left_wrap_ud_style.alignment = wrap_left_al
	song_left_wrap_ud_style.font = song_font
	song_left_wrap_ud_style.borders = ud_border

	# 黑体14号，垂直居中，水平居左，不换行，无框
	# 紧急程度，省总工会办公室，承办人：张腾
	black_left_nowrap_nobord_style = xlwt.XFStyle()
	black_left_nowrap_nobord_style.alignment = nowrap_left_al
	black_left_nowrap_nobord_style.font = black_font

	# 宋体14号，垂直居中，水平居左，换行，无框
	# 收文号
	song_left_nowrap_nobord_style = xlwt.XFStyle()
	song_left_nowrap_nobord_style.alignment = wrap_left_al
	song_left_nowrap_nobord_style.font = song_font

	# 方正小标宋_GBK 20号，水平垂直居中，换行，无框
	# 山西省总工会来文、来电阅办卡片
	black_mid_wrap_nobord_style = xlwt.XFStyle()
	black_mid_wrap_nobord_style.alignment = wrap_mid_al
	black_mid_wrap_nobord_style.font = black_font1

	# 设置单元格宽度
	ws.col(0).width = 3500
	ws.col(1).width = 8000
	ws.col(2).width = 4500
	ws.col(3).width = 8000
	ws.row(0).set_style(xlwt.easyxf('font:height 600;'))  # 行高37.5
	ws.row(1).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(2).set_style(xlwt.easyxf('font:height 650;'))
	ws.row(3).set_style(xlwt.easyxf('font:height 700;'))
	ws.row(4).set_style(xlwt.easyxf('font:height 3000;'))
	ws.row(5).set_style(xlwt.easyxf('font:height 3900;'))
	ws.row(6).set_style(xlwt.easyxf('font:height 650;'))  # 行高40.5
	ws.row(7).set_style(xlwt.easyxf('font:height 780;'))  # 行高50
	ws.row(8).set_style(xlwt.easyxf('font:height 480;'))  # 行高30
	ws.row(9).set_style(xlwt.easyxf('font:height 570;'))  # 行高30
	ws.row(10).set_style(xlwt.easyxf('font:height 300;'))  # 行高18

	ws.write(0, 0, '紧急程度:', black_left_nowrap_nobord_style)
	ws.write(0, 3, '       收文号：[2019]' + str(list[0]), song_left_nowrap_nobord_style)
	ws.write_merge(1, 1, 0, 3, '山西省总工会来文、来电阅办卡片', black_mid_wrap_nobord_style)
	ws.write(2, 0, '来文单位', black_mid_wrap_udr_style)
	ws.write(2, 1, list[1], song_mid_wrap_u_style)
	ws.write(2, 2, '收文时间', black_mid_wrap_all_style)
	ws.write(2, 3, list[2], song_mid_wrap_u_style)
	ws.write(3, 0, '来文名称', black_mid_wrap_udr_style)
	ws.write_merge(3, 3, 1, 3, list[3], song_mid_wrap_u_style)
	ws.write(4, 0, '主要内容', black_mid_wrap_udr_style)
	ws.write_merge(4, 4, 1, 3, list[4], song_left_wrap_u_style)
	ws.write(5, 0, '领\n导\n批\n示', black_mid_wrap_udr_style)
	ws.write_merge(5, 5, 1, 3, '', song_left_wrap_u_style)
	ws.write_merge(6, 8, 0, 0, '拟办意见', black_mid_wrap_udr_style)
	ws.write_merge(6, 6, 1, 3, list[5], song_left_wrap_u_style)
	ws.write_merge(7, 7, 1, 3, list[6], song_left_nowrap_nobord_style)
	ws.write(9, 0, '处理结果', black_mid_wrap_udr_style)
	ws.write_merge(9, 9, 1, 3, '', song_left_wrap_ud_style)
	ws.write(10, 0, ' 省总工会办公室', black_left_nowrap_nobord_style)  # 5个空格
	ws.write(10, 3, '        承办人：张  腾', black_left_nowrap_nobord_style)

	wb.save('C:\\tempfile.xls')
	logger.loggerprint("无会议，写入临时xls，完毕")