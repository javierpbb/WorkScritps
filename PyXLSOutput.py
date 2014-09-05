import sys
import xlwt
import datetime

d=datetime.date.today()
t=datetime.datetime.now()

print t

workbook = xlwt.Workbook() 
sheet = workbook.add_sheet("Prueba") 

sheet.write(0,0,"Entero")
sheet.write(0,1,"Texto")
sheet.write(0,2,"Fecha Dia")
sheet.write(0,3,"Fecha Hora")

for i in range (1,600):
	linea = "Cualquiera " + str(i)
	sheet.write(i,0,i)
	sheet.write(i,1,linea)
	sheet.write(i,2,d)
	sheet.write(i,3,t)
	t=t-datetime.timedelta(hours=2)
	d=d-datetime.timedelta(days=10)

workbook.save("PruJav.xls") 

print sys.version
