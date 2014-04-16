require "watir-webdriver"
require 'test/unit'

class TC_article_example < Test::Unit::TestCase

  def test_search
    
    # Contador de Shipments
    cont = 1
    
    # Creator del Browser
    ff = Watir::Browser.new :ff 
      
    # Abrimos el fichero con los shipments a spoofear
    File.open('C:\Users\jprieto\Documents\M&S\Issues\shipments.txt').each_line do |ship|
      
      # Navegamos a pagina Spoofer
          
      ff.goto("http://zzzzzzzzzzzzzzzzzzzz.amazon.com/CoreWebsite/ASNSpoofer")

      # Cargamos el shipment en el field ShipID
      ff.text_field(:name, "shipmentId").set(ship)
      ff.button(:value, "Load").click
    
      # OPCIONAL Lanzamos el DeletePackage si el shipment esta en cond 6009
      # ff.button(:name, "DeletePackages").click
      
      # Lanzamos el boton Spoof ASN
      ff.button(:name, "SpoofASN").click
      
      # busqueda de error en el retorno del spoof cuando hay partial ASN. Falla cuando sale bien el ASN, solo funcion
      # si da error el spoofASN
      
      text1 = nil
      # text1 = ff.span(:class, "errorMsg").text
      if text1 != nil
        puts text1
        text1 = nil
      end
      
      # fin de la busqueda del error en el retorno de un partial ASN
      
      # Logeamos actividad
      print "Shipment num ", cont, " - ", ship
      
      cont = cont + 1
       
    end
    
    #Cerramos Browser
    ff.close
          
      
  end

end
