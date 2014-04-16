("JavOrd", conn)
borrar_tabla("LalOrd", conn)
borrar_tabla("OrdJavNotInLal", conn)
print("Tablas Borradas")

cur = conn.cursor()

print("Insertando JavOrd")

dataFile = open('cargaJavord.txt', 'r')

for eachLine in dataFile:
    cell = str(eachLine)
    sql="""insert into JavOrd(OrderID) values (%s)"""
    cur.execute(sql, cell)
    conn.commit()

#cur.execute("SELECT OrderID FROM javord")

dataFile.close()


print("Insertando LalOrd")

dataFile = open('cargaLalord.txt', 'r')

for eachLine in dataFile:
    cell = eachLine
    sql="""insert into LalOrd(OrderID) values (%s)"""
    cur.execute(sql, cell)
    conn.commit()

#cur.execute("SELECT OrderID FROM javord")

dataFile.close()

print("Insertando Excepciones")

cur.execute("SELECT OrderID FROM ordercomp.JavOrd WHERE OrderID NOT IN ( SELECT OrderID FROM ordercomp.LalOrd ) ORDER BY OrderID;")

dataFile = open('discrepantes.txt', 'w')

i=0
for row in cur.fetchall():
    cell=row
    i=i+1
    insert="""insert into ordjavnotinlal(OrderID) values (%s)"""
    cur.execute(insert, cell)
    conn.commit()
    dataFile.write(str(cell[0]))

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
cur.execute("SELECT OrderID FROM ordercomp.JavOrd ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 0, cell)
  i=i+1

#insertamos Orders Lalith
print("insertamos Orders Lalith")
cur.execute("SELECT OrderID FROM ordercomp.LalOrd ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 1, cell)
  i=i+1

#insertamos discrepantes
print("insertamos discrepantes")
cur.execute("SELECT OrderID FROM ordercomp.JavOrd WHERE OrderID NOT IN ( SELECT OrderID FROM ordercomp.LalOrd ) ORDER BY OrderID;")

i=1
for row in cur.fetchall():
  cell=row[0]
  worksheet.write(i, 2, cell)
  i=i+1

workbook.close()

cur.close()

conn.close()
