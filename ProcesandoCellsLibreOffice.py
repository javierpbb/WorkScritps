# de aqui he podido sacar los metodos aunque es Basic https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
#
# Compara los datos de dos columnas y pone los existentes en la primera y no en la segunda en una tercera.
# Importante. Las columnas clasificadas ascendentemente para optimizar. 
#
#
# CONSTANTES A PREFIJAR: Rangos del las columnnas 
#
# Tab a 2 espacios
#
#
import uno

oDoc = XSCRIPTCONTEXT.getDocument()

def ProcesandoSheet():
  # k sera el indice de las orders validas en la columna R 
  
  i=2
  j=2
  k=2
  p=2
  
  oSheet = oDoc.CurrentController.ActiveSheet

  for i in range(2,3283):  # El rango de X tiene que ser uno mayor ya que empieza en zero
    oCell1 = oSheet.getCellRangeByName("C" + str(i))

    igual=0
    j=p

    oCell2 = oSheet.getCellRangeByName("I" + str(j))
    while j <= 181 and igual == 0 and oCell2.String <= oCell1.String:
      if oCell2.String == oCell1.String:
        igual = 1
        p=j
        
      j=j+1
      oCell2 = oSheet.getCellRangeByName("I" + str(j))
    
    if igual == 0:
      oCell3 = oSheet.getCellRangeByName("O" + str(k))
      oCell3.String = oCell1.String
      oCellp = oSheet.getCellRangeByName("P" + str(k))
      oCellp.Value = p
      k=k+1

#  for x in range(3,10):
#  	oCell1.Value = x    # usando value hemos conseguido asignar valor
	
#  oCell1.String = "Mierda"
#  oCell2 = oSheet.getCellRangeByName("I1")
#  oCell2.Value = oCell2.Value + 1
