import xlrd
import MySQLdb
import re,fnmatch
import datetime

#For installing new packages
#import mysqlclient for python 3.x
#apt-get install python-dev libmysqlclient-dev
#pip install MySQL-python

# Open the workbook and define the worksheet
book = xlrd.open_workbook("/home/apollo/samp.xls")
sheet = book.sheet_by_index(0)
sheet_names=book.sheet_names()
sheet_names=[i.encode('UTF-8') for i in sheet_names]
database = MySQLdb.connect (host="localhost", user = "root", passwd = "root", db = "test")
cursor = database.cursor()
database.set_character_set('utf8')
cursor.execute('SET NAMES utf8;')
cursor.execute('SET CHARACTER SET utf8;')
cursor.execute('SET character_set_connection=utf8;')



col_u=[sheet.cell(0,i).value for i in range(sheet.ncols)]
col_names=[i.encode('UTF-8') for i in col_u]
col_names=[i.replace(" ","_") for i in col_names]
date=fnmatch.translate('*Date*')
date=[f for f in col_names if re.match(date,f)]
ind=[col_names.index(date[i]) for i in range(len(date))]		
del col_u
col_name=','.join(map(str, col_names))

n=len(col_names)
m='%s'*n
m=m.replace('%s','%s,').rstrip(',')
query = "INSERT INTO orders ("+col_name+") VALUES ("+m+");"
values=()
for r in range(1,sheet.nrows):
	values=()
	col_val=[sheet.cell(r,i).value for i in range(sheet.ncols)]
	for i in range(len(ind)):	
		col_val[ind[i]] = datetime.datetime(*xlrd.xldate_as_tuple(col_val[ind[i]], book.datemode))	
	values = tuple(col_val)
	cursor.execute(query,values)
	del values
	del col_val		
cursor.close()
database.commit()
database.close()

print('Added '+str(sheet.nrows)+' rows and '+str(sheet.ncols)+' columns')
