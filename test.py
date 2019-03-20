import tempfile
import win32api
import win32print



file_name = "C:\\tempfile.xls"
#file_name = tempfile.mktemp('.txt')
#open(file_name,'w').write('test')

win32api.ShellExecute(
		0,
		"print",
		file_name,
		'/d:"%s"' % win32print.GetDefaultPrinter(),
		".",
		0
	)