' de aqui he podido sacar los metodos aunque es Basic https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
'
' Compara los datos de dos columnas y pone los existentes en la primera y no en la segunda en una tercera.
' Importante. Las columnas clasificadas ascendentemente para optimizar.
'
'
' CONSTANTES A PREFIJAR: Rangos del las columnnas
'
' Tab a 2 espacios
'
Sub BuscandoNonInLal()

  ' k sera el indice de las orders validas en la columna R
 
  
  i = 2
  j = 2
  k = 2
  p = 2
  
  Dim oCell1 As Range
  Dim oCell2 As Range
 
  For Each oCell1 In Range("C2:C3282")
    
    igual = 0
    j = p
    
    Set oCell2 = Sheets("OutstandingApr15").Cells(j, 9)
    
    Do While j <= 181 And igual = 0 And oCell2.Text <= oCell1.Text
      
      If oCell2.Text = oCell1.Text Then
        igual = 1
        p = j
      End If
      
      j = j + 1
      Set oCell2 = Sheets("OutstandingApr15").Cells(j, 9)
    
    Loop
    
    If igual = 0 Then
      Cells(k, 14) = oCell1
      Cells(k, 15) = p
      k = k + 1
    End If
    
  Next


End Sub
