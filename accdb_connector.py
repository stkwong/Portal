import pyodbc

def retri_outgoing(path, tablename, order_col, des, *colref, **keyref):
	conn = pyodbc.connect(
		r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_data.accdb;")
	cursor = conn.cursor()
	rows = []
	reflist = []

	colreflist = []
	for col in colref:
		colreflist.append(col)
	colrefstr = " , ".join(colreflist)

	if len(keyref) == 0:

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename + " ORDER BY " + order_col + " " + des + ";")
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	else:
		for k, v in keyref.items():
			reflist.append(k + " = " + v)
		refstr = " AND ".join(reflist)

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename + " WHERE " + refstr + " ORDER BY " + order_col + " " + des + ";")
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	return rows

def retri_sql(path, tablename, *colref, **keyref):
	conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_data.accdb;")
	cursor = conn.cursor()
	rows = []
	reflist = []

	colreflist = []
	for col in colref:
		colreflist.append(col)
	colrefstr = " , ".join(colreflist)

	if len(keyref) == 0:

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename)
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	else:
		for k, v in keyref.items():
			reflist.append(k + " = " + v)
		refstr = " AND ".join(reflist)

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename + " WHERE " + refstr + ";")
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	return rows

def insert_sql(path, tablename, **keyref):
	conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_data.accdb;")
	cursor = conn.cursor()
	keylst = []
	vallst = []
	strlst = []
	for k, v in keyref.items():
		keylst.append(k)
		vallst.append(v)
	for i in range(len(keylst)):
		strlst.append("?")
	keylst_str = " , ".join(keylst)
	strlst_str = " , ".join(strlst)
	val = tuple(vallst)
	query = 'INSERT INTO ' + tablename + ' ( ' + keylst_str + ' ) ' + ' VALUES(' + strlst_str + ');'

	try:
		cursor.execute( query , val)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

def update_sql_2key(tablename, colind, keyrefA, keyrefB, keyindA, keyindB, data):
	query = "UPDATE " + tablename + " SET " + colind + " = %s WHERE " + keyrefA + "= %s AND " + keyrefB + "= %s"
	input = (str(data), keyindA, keyindB)
	try:
		conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" r"DBQ=C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Database\SMS_data.accdb;")
		cursor = conn.cursor()
		cursor.execute(query, input)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

def update_sql_1key(path, tablename, colind, keyref, keyind, data):
	query = "UPDATE " + tablename + " SET " + colind + " = ? WHERE " + keyref + "= ?"
	input = (str(data), keyind)
	try:
		conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_data.accdb;")
		cursor = conn.cursor()
		cursor.execute(query, input)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

def update_sql_3key(tablename, colind, keyrefA, keyrefB, keyrefC, keyindA, keyindB, keyindC, data):
	query = "UPDATE " + tablename + " SET " + colind + " = %s WHERE " + keyrefA + "= %s AND " + keyrefB + "= %s AND " + keyrefC + "= %s"
	input = (str(data), keyindA, keyindB, keyindC)
	try:
		conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" r"DBQ=C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Database\SMS_data.accdb;")
		cursor = conn.cursor()
		cursor.execute(query, input)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

def retri_path(path, tablename, *colref, **keyref):
	conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_path.accdb;")
	cursor = conn.cursor()
	rows = []
	reflist = []

	colreflist = []
	for col in colref:
		colreflist.append(col)
	colrefstr = " , ".join(colreflist)

	if len(keyref) == 0:

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename)
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	else:
		for k, v in keyref.items():
			reflist.append(k + " = " + v)
		refstr = " AND ".join(reflist)

		try:
			cursor.execute("SELECT " + colrefstr + " FROM " + tablename + " WHERE " + refstr + ";")
			rows = cursor.fetchall()

		finally:
			cursor.close()
			conn.close()

	return rows

def insert_path(path, tablename, **keyref):
	#conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" r"DBQ=C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Database\SMS_data.accdb;")
	conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_path.accdb;")
	cursor = conn.cursor()
	keylst = []
	vallst = []
	strlst = []
	for k, v in keyref.items():
		keylst.append(k)
		vallst.append(v)
	for i in range(len(keylst)):
		strlst.append("?")
	keylst_str = " , ".join(keylst)
	strlst_str = " , ".join(strlst)
	val = tuple(vallst)
	query = 'INSERT INTO ' + tablename + ' ( ' + keylst_str + ' ) ' + ' VALUES(' + strlst_str + ');'

	try:
		cursor.execute( query , val)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

def update_path_1key(path, tablename, colind, keyref, keyind, data):
	query = "UPDATE " + tablename + " SET " + colind + " = ? WHERE " + keyref + "= ?"
	input = (str(data), keyind)
	try:
		conn = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_path.accdb;")
		cursor = conn.cursor()
		cursor.execute(query, input)
		conn.commit()

	finally:
		cursor.close()
		conn.close()

# def update_sql_lst(path, tablename, *colind_data, **keyref):
# 	reflist = [ ]
# 	colind_lst = [ ]
# 	data_lst = []
# 	for update_item in colind_data:
# 		update_item_lst = update_item.split("=")
# 		colind_lst.append(update_item_lst[0] + " = ? ")
# 		data_lst.append(update_item_lst[ 1 ])
# 	colind_lst_str = ", ".join( colind_lst )
# 	print ("colind_lst_str : ", colind_lst_str)
#
# 	for k , v in keyref.items():
# 		reflist.append( k + " = ? ")
# 		data_lst.append(v)
# 	ref_lst_str = " AND ".join( reflist )
# 	print("ref_lst_str : ", ref_lst_str)
#
# 	query = "UPDATE " + tablename + " SET " + colind_lst_str + " WHERE " + ref_lst_str
# 	input = tuple(data_lst)
# 	print(query, input)
# 	#print (query)
# 	try:
# 		conn = pyodbc.connect( r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" "DBQ=" + path + r"\Database\SMS_path.accdb;")
# 		cursor = conn.cursor()
# 		cursor.execute(query, input)
# 		conn.commit()
#
# 	finally:
# 		cursor.close()
# 		conn.close()