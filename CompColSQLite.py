import sqlite3
import string
import xlsxwriter
import sys

#FUNCIONES
def borrar_tabla(tabla,conn):
  cur = conn.cursor()
  sql="delete from " + tabla
  cur.execute(sql)
  conn.commit()

conn = None
conn = sqlite3.connect('C:/Users/jprieto/Documents/M&S/issues/Discrepante.sqlite3')

print("Borrando Tablas")
borrar_tabla("JavOrd", conn)
borrar_tabla("LalOrd", conn)
borrar_tabla("OrdJavNotInLal", conn)
print("Tablas Borradas")

print("Insertando JavOrd")
cur = conn.cursor()

dataFile = open('cargaJavord.txt', 'r')

for eachLine in dataFile:
  cell = str(eachLine)
  sql="insert into JavOrd(OrderID) values ('" + str(cell) + "');"
  cur.execute(sql)
  conn.commit()

dataFile.close()

print("Insertando LalOrd")

dataFile = open('cargaLalord.txt', 'r')

for eachLine in dataFile:
  cell = eachLine
  sql="insert into LalOrd(OrderID) values ('" + str(cell) + "');"
  cur.execute(sql)
  conn.commit()

dataFile.close()

print("Insertando Excepciones")

cur.execute("SELECT OrderID FROM JavOrd WHERE OrderID NOT IN ( SELECT OrderID FROM LalOrd ) ORDER BY OrderID;")

dataFile = open('discrepantes.txt', 'w')

for row in cur.fetchall():
  insert="insert into ordjavnotinlal(OrderID) values ('" + str(row[0]) + "');"
  cur.execute(insert)
  conn.commit()
  dataFile.write(str(row[0]))

dataFile.close()

print("Creando XLSX")

#Creacion de la spreadsheet modulo xlsxwriter
workbook = xlsxwriter.Workbook('ordersJavLal.xlsx')
worksheet = workbook.add_worksheet()
#escribimos cabeceras
worksheet.write(0, 0, 'Jav OrderID')
worksheet.write(0, 1, 'Lalith OrderID')
worksheet.write(0, 2, 'JavOrd not in Lal')

#insertamos Orders Javier
print("insertamos Orders Javier")
cur.execute("SELECT OrderID FROM JavOrd ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 0, cell)
  i=i+1

#insertamos Orders Lalith
print("insertamos Orders Lalith")
cur.execute("SELECT OrderID FROM LalOrd ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 1, cell)
  i=i+1

#insertamos discrepantes
print("insertamos discrepantes")
cur.execute("SELECT OrderID FROM JavOrd WHERE OrderID NOT IN ( SELECT OrderID FROM LalOrd ) ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 2, cell)
  i=i+1

workbook.close()

#salida programa
cur.close()

conn.close()

sys.exit()
