Public Sub Main()                                 
    Dim Articulo
    Dim Importe 
    Dim Precio
    Dim Cantidad
    Dim Descripcion
	Dim porcGlob
    
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

   'Se agrega lineas para descuento y lista 2
'	if WeekDay(date) = 2 Then
'      porcGlob = 25
'    else
'      porcGlob = 0
'    end if
'	Por ahora 15% todos los dias
'	asi como esta ahorita no va a haber precios dos
	porcGlob = 15
	CambiaPrecioDos(porcGlob)
	DescuentoFnc(porcGlob)

    On Error Resume Next
    ColocaAsociados
End Sub

Sub CambiaPrecioDos(porcGlob)
	Dim precio
	Dim numLista
	Dim cantPrecioDos
	Dim flagCalcular
	flagCalcular = 0
	For n = 1 To fg2.Rows - 1
		If clEmpty( fg2.TextMatrix( n, 0 ) ) Then
       			Exit For
       	End If
		cantPrecioDos = CreaRecordSet( "SELECT C2 FROM prods where articulo = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30))", Ambiente.Connection )
		cantPrecioDos = Val2(cantPrecioDos(0))
		if fg2.TextMatrix( n, 1 ) >= cantPrecioDos  and cantPrecioDos > 0 then
			if porcGlob = 0 and (Val2(fg2.TextMatrix( n, 1 )) Mod Val2(cantPrecioDos)) = 0 then
				numLista = 2		
			else
				numLista = 1
    		end if
			precio = CreaRecordSet( "SELECT PRECIO" & numLista & " FROM prods where articulo = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30))", Ambiente.Connection )
			precio = Val2(precio(0))
			fg2.TextMatrix(n, 2) = precio
			Ambiente.Connection.Execute "UPDATE partvta SET precio = " & precio & ", lista = " & numLista & " WHERE (venta = " & venta & " AND ARTICULO = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30)) AND cast(cantidad as int) = " & fg2.TextMatrix( n, 1 ) & ")"
			flagCalcular = 1
    	end if
	Next
	if flagCalcular = 1 then
		CalculaImportes
	end if
End Sub

Sub DescuentoFnc(porcGlob)                   
	if porcGlob > 0 then
		const desLimite = 1000
		Dim montoAcum
		Dim montosinDes
		Dim descFin 
		Dim flagCalcular
		flagCalcular = 0
        ' obtiene el total acumulado en la venta actual por todos los articulos vendidos actualmente     
	    montoAcum = 0
		For n = 1 To fg2.Rows - 1
        	If clEmpty( fg2.TextMatrix( n, 0 ) ) Then
       			Exit For
       		End If
			flagSinDes = 0
			descFin = porcGlob 
			montoSinDes = fg2.TextMatrix( n, 1 ) * fg2.TextMatrix( n, 2 ) * ( 1 + (fg2.TextMatrix( n, 5 ) / 100) )
			' PONER QUE SI VALE MENOS DE 6 PESOS, SI SON MENOS 6 UNID 0%, 6 A 10 15%, MAS DE 10 COMPLETO
			if fg2.TextMatrix( n, 2 ) * ( 1 + (fg2.TextMatrix( n, 5 ) / 100) ) < 6 then
				If fg2.TextMatrix( n, 1 ) < 5 then
					descFin = 0
				ElseIf fg2.TextMatrix( n, 1 ) < 10 then
					descFin = porcGlob*3/5
				end If
			end if
			' revisar si el articulo tiene descuento especial
			Set descuentoEsp = CreaRecordSet( "SELECT * from descuentoEspecial where ARTICULO = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30))", Ambiente.Connection )
			if Val2(descuentoEsp("articulo")) <> 0 then
				if descuentoEsp("pocoDesc") then
					if Val2(descuentoEsp("descuento")) < porcGlob then
						descFin = Val2(descuentoEsp("descuento"))
					end if	
				else
					descFin = 0
				end if
			end if 
			' para revisar limite de descuento
			if descFin <> 0 then
				montoAcum = montoAcum + montoSinDes
			end if   
			if montoAcum > desLimite then
				if montoAcum-montoSinDes < desLimite then
					descFin = descFin * (desLimite-(montoAcum-montoSinDes)) / montoSinDes
				else
					descFin = 0
            	end if
			end if
			' actualizar descuento
			if	Val2(fg2.TextMatrix( n, 3 )) <> Val2(descFin) then
        		fg2.TextMatrix( n, 3 ) = descFin
				Ambiente.Connection.Execute "UPDATE partvta SET descuento = " & descFin & " WHERE (venta = " & venta & " AND ARTICULO = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30)))"
				flagCalcular = 1
			end if
      	Next
		' recalcula importes
		if flagCalcular = 1 then
      		CalculaImportes
		end if
	end if
End Sub


Sub CambiaDescuento( nDescuento )
    ' La variable venta contiene el n�mero de venta de la operaci�n que se esta
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