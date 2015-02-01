REM  *****  BASIC  *****

Sub LoadDataFromDBNueva

rem base de datos
dim oDatabase As Object
Dim oTable As Object, oField As Object, i As Integer

Set oDatabase = Application.OpenDatabase("/home/javier/Desarrollo/libreoffice/ODOO8.odb","odoo","jprietob",)

rem Hoja de calculo
doc = ThisComponent
addr = doc.getCurrentSelection().getCellAddress()
sheet = doc.getSheets().getByIndex(addr.Sheet)

rem calculamos el limite celda J8 (9,7)
cell = sheet.getCellByPosition (9, 7)
limite = cell.Value	

print "abierta base"

Dim orsRecords As Object, lCount As Long
Set orsRecords = oDatabase.OpenRecordset("QueryAccount_Accounts", , , )

lCount = 1
With orsRecords
	If Not .BOF Then		'	An empty recordset has both .BOF and .EOF set to True
		Do While Not .EOF and lCount < limite
			'asignamos valor columna A
			cell = sheet.getCellByPosition (0, lCount)
	        cell.Value = .fields(0).value
	        
	        'asignamos texto columna K
	        cell = sheet.getCellByPosition (10, lCount)
	        cell.String = .fields(0).value
	        
	        'asignamos demas columnas
	        cell = sheet.getCellByPosition (1, lCount)
	        cell.String = .fields(1).Value
	        
	        cell = sheet.getCellByPosition (2, lCount)
	        cell.Value = .fields(4).value
	        
			.MoveNext
			
			lCount = lCount + 1
		Loop
	End If
	print lCount
	.mClose()
End With

Print "Number of records = " & lCount

With oDatabase
	For i = 0 To .TableDefs.Count - 1
		Set oTable = .TableDefs(i)
		DebugPrint oTable.Name
	Next i
	
	print "Tablas", i
	
End With

oDatabase.mClose()

End Sub


