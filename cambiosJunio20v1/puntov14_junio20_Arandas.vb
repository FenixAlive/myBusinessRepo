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
	if WeekDay(date) = 2 Then
      porcGlob = 25
    else
      porcGlob = 0
    end if
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
	const desLimite = 1000
	Dim montoAcum
	Dim montosinDes
	Dim porcPoco
	Dim prodSinDes
	Dim prodPocoDes
	Dim descFin
	Dim flagSinDes
	' 				   GELMICIN,       B.PINAVERIO,    CAPTOPRIL,       CAPTOPRIL,       DICLOFENACO,     DICLOFENACO,     DICLOFENACO,     DICLOFENACO,      ENALAPRIL,      ENALAPRIL,       ENALAPRIL,       GABAPENTINA 15, KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      LORATADINA 10,   LORATADINA 20,   LORATADINA 20,   LOSARTAN 50MG,   LOSARTAN 100MG   METFORMINA,      METFORMINA,      METOPROLOL,      METOPROLOL,      OMEPRAZOL 14,    OMEPRAZOL 14,    OMEPRAZOL 14,    OMEPRAZOL 30,    OMEPRAZOL 30,    PANTOPRAZOL 20 7, PANTOPRAZOL 20 14, PANTOPRAZOL 40 14, SILDENAFIL 50 1, SILDENAFIL 50 1, SILDENAFIL 50 1, SILDENAFIL 100 1, SILDENAFIL 100 1, PIROXICAM,       ATORVASTATINA ,  ALCOHOL 108, CUBRE 118, CUBRE 158, CUBRE Y GEL 156, GEL 154, GEL 155, GEL MUN,         GEL BRITZ,       GEL JALOMA,      GUANTE 121, GUANTE ESTERIL,  TOALLAS DESINF,  GEL 157, MASCARA,        SEDALMERK 126, TERRAMICINA 155
	prodSinDes = Array("780083148966", "821998000434", "7503000422405", "7501537102142", "7503000422368", "7502216802964", "7502009746017", "7501482201952", "7502216792845", "7502216791299", "7502227872246", "785120754919", "7501493888760", "7501557140308", "7501349024847", "7501573900559", "7502216791312", "7502009740244", "7501075716924", "7502216790339", "7502216793415", "7502009742828", "7502240450070", "7502240450711", "7501075717020", "7503004908776", "7501075714173", "7501075718881", "7501573902584", "7501571201221", "7501493888746", "7501342803562", "7502216792760", "7501349022485",  "7501349028364",   "7501349028364",   "7502216796812", "7501258210133", "7502009744457", "7502009744914", "7502216796836",   "7501537102371", "7501349024526", "108",       "118",     "158",     "156",           "154",   "155",   "7508006012117", "7503009088664", "759684154140", "121",       "7502224240659", "7503006503474", "157",   "7506329100214", "126",        "155")
	'                   CARBAMAZEPINA,   BUTILHIOSINA,   VENGESIC,       ITRACONAZOL,          PAROXETINA,      ATORVASTATINA 40 10, ATORVASTATINA 40 10, PAROXETINA 10,   PAROXETINA 10,   DULOXETINA,      ADIOLOL,        ENFO-KI,         RMFLEX,          INFACAR ET 1,    INFACAR ET 2,    ACENOCUMAROL,    TRIBEDOCE COMP  
	prodPocoDes = Array("7501075710724", "821998000601", "780083140731", "7501075717174",      "7501349024939", "785120754858",      "7501349028913",     "7501349024939", "7501075718041", "7502009745478", "725742762145", "7501590285608", "7508304309513", "7502253072191", "7502253072207", "7501471889611", "7501537164713")
 	porcPoco =    Array(17.5,            10,             4.55,           6.25,                 6.25,            6.25,                6.25,                6.25,            6.25,            5,               2.5,            10,              11.22,           6.25,            6.25,            19.7,            16.67)   
	
	if porcGlob > 0 then
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
					descFin = 15
				end If
			end if
			For i = LBound(prodSinDes) To Ubound(prodSinDes)
            	if prodSinDes(i) = fg2.TextMatrix( n, 0 ) then
               		descFin = 0
					flagSinDes = 1
      			end if
         	Next
			if flagSinDes = 0 then
				montoAcum = montoAcum + montoSinDes
			end if
		 	if montoAcum > desLimite then
				if montoAcum-montoSinDes < desLimite then
					if flagSinDes = 0 then
						' revisa los articulos que llevan un descuento menor al propuesto
     		       		For i = LBound(prodPocoDes) To Ubound(prodPocoDes)
       		       	 		if prodPocoDes(i) = fg2.TextMatrix( n, 0 ) then
         	            		descFin = porcPoco(i) 
          	     			end if
           		 		Next
					end if
					descFin = descFin * (desLimite-(montoAcum-montoSinDes)) / montoSinDes
				else
					descFin = 0
            	end if
		 	else
         		' revisa los articulos que llevan un descuento menor al propuesto
				if flagSinDes = 0 then
     		       	For i = LBound(prodPocoDes) To Ubound(prodPocoDes)
       		       	 	if prodPocoDes(i) = fg2.TextMatrix( n, 0 ) then
         	            	descFin = porcPoco(i) 
          	     		end if
           		 	Next
				end if
			end if
        	fg2.TextMatrix( n, 3 ) = descFin
			Ambiente.Connection.Execute "UPDATE partvta SET descuento = " & descFin & " WHERE (venta = " & venta & " AND ARTICULO = cast(" & fg2.TextMatrix( n, 0 ) & " AS nvarchar(30)))"
      	Next
      	CalculaImportes
	end if
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