Public Sub Main()
    Dim rstCantidad
    Dim nCantidad
    Dim rstProd                                  
    Dim Articulo
    Dim rstArticulo 
    Dim Precio
    Dim Cantidad
    Dim Descripcion


    'Set rstSuma = CreaRecordSet( "SELECT SUM( cantidad * precio * ( 1 - (descuento / 100) ) * ( 1 + (impuesto / 100) ) ) As importe FROM partvta WHERE venta = " & Me.Venta, Ambiente.Connection )
    
    'nCantidad = Val2( rstSuma(0) )  


    'If nCantidad >= 1500 Then      
    '   CambiaDescuento 30
    '   Eventos
    'Else
    '   txtFields(3) = "En ventas de mas de $1,500.00 descuento del 30%"
    '   CambiaDescuento 0
    '   Eventos
    'End If                     

'// <-- calculo del descuento limite 1,000 en punto de venta.  
	 
' obtiene el total acumulado en la venta actual por todos los articulos vendidos actualmente     
    Set rstSuma = CreaRecordSet( "SELECT SUM( cantidad * precio * ( 1 + (impuesto / 100) ) ) As importe FROM partvta WHERE venta = " & Me.Venta, Ambiente.Connection )
' convierte la cantidad a valor numerico
	nCantidad = Val2( rstSuma(0) )
    
'despues de calcular el monto del dia veo si no rebasa 1,000 pesos y si es lunes                
	Dim DiaSemana
	Dim descuento
	const porc = 27
	DiaSemana = WeekDay(date)
	if DiaSemana = 2 then  
		if nCantidad > 1000 then
			' aplica un descuento a la venta que no rebasa 1,000
			descuento = porc*1000/nCantidad
		else	
			descuento = porc   
		end if     
		cambiadescuento descuento 
	end if            
'msgbox(descuento)
    
'// termina calculo de descuento limite 1,000 -->

    if UltimaPartida > 0 Then
       Articulo = fg2.TextMatrix( UltimaPartida, 0 )
       Cantidad =  fg2.TextMatrix( UltimaPartida, 1 )
       Precio =  fg2.TextMatrix( UltimaPartida, 2 )      
       Importe = Val2( fg2.TextMatrix( UltimaPartida, 16 ) )
       Descripcion =  Mid( Trim(fg2.TextMatrix( UltimaPartida, 6 )), 1, 30 )             
    end if 

    if Ambiente.rstEstacion("torreta") <> 0 Then
       if clAt( "LPT", Ambiente.rstEstacion("ptorreta") ) > 0 Then
          Out Trim(Ambiente.rstEstacion("ptorreta")), Chr( 12 ) & Descripcion & "  $" & Importe
       else
          if Ambiente.torreta.PortOpen Then
             Ambiente.Torreta.OutPut = Chr( 12 ) & Descripcion & "  $" & Importe
          end if 
       end if
    end if   
       

    On Error Resume Next
    ColocaAsociados

End Sub


Sub CambiaDescuento( nDescuento )

    ' La variable venta contiene el número de venta de la operación que se esta
    ' realizando 
    Ambiente.Connection.Execute "UPDATE partvta SET descuento = " & nDescuento & " WHERE venta = " & venta

    For n = 1 To fg2.Rows - 1
        
        If clEmpty( fg2.TextMatrix( n, 0 ) ) Then
           Exit For
        End If   

        fg2.TextMatrix( n, 3 ) = nDescuento
          
    Next

    CalculaImportes

End Sub




Public Sub CambiaPrecio( nPrecio )

       Set rstPartidas = CreaRecordSet( "SELECT prods.articulo, partvta.impuesto, partvta.id_salida, prods.precio2, prods.precio1 FROM partvta INNER JOIN prods ON partvta.articulo = prods.articulo WHERE partvta.venta = " & Me.Venta, Ambiente.Connection )

       While Not rstPartidas.EOF
             
             if rstPartidas("Precio" & nPrecio) > 0 Then
                Query.Reset
                Query.strState = "UPDATE"
                Query.AddField "partvta", "precio", rstPartidas("Precio" & nPrecio)
                Query.Condition = "id_salida = " & rstPartidas("id_salida")
                Query.CreateQuery
                Query.Execute

                For n = 1 To fg2.Rows - 1

                    If clEmpty(fg2.TextMatrix(n, 0)) Then
                       Exit For
                    End If

                    If Trim(fg2.TextMatrix(n, 0)) = Trim(rstPartidas("articulo")) Then
                       fg2.TextMatrix(n, 2) = Formato(rstPartidas("Precio" & nPrecio) * ( 1 + (rstPartidas("impuesto") / 100) ), Ambiente.FDinero)
                       fg2.TextMatrix(n, 9) = 2
                    End If

                Next

             end if


             rstPartidas.MoveNext
       Wend

       CalculaImportes

End Sub



Sub ColocaAsociados()    
    Dim rstAsociados
    Dim cAsociado 

    cAsociado = ""
    Set rstAsociados = CreaRecordSet("SELECT asociados.articulo, asociados.observ, prods.descrip FROM asociados INNER JOIN prods ON asociados.articulo = prods.articulo WHERE padre = '" & rstArticulo("Articulo") & "'", Ambiente.Connection)
    
    While Not rstAsociados.EOF
          cAsociado = cAsociado & " " & Trim(rstAsociados("articulo")) & " " & Trim(rstAsociados("descrip")) & " " & Trim(rstAsociados("Observ")) & vbCrLf
          rstAsociados.MoveNext 
    Wend

    If Not clEmpty( (cAsociado) ) Then
       cAsociado = "Recomendaciones: " & cAsociado
       EnviaMensajeGrid (cAsociado)  
    End If

End Sub





