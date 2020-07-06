Sub Main()                 

    'ImprimeSalida
    ImprimeEntradaPorTicket

End Sub     



Sub ImprimeEntradaPorTicket()
    Dim c
    Dim rstEntrada
    Dim rstPartidas
    Dim cCantidad
    Dim cPrecio
    Dim cImporte
    Dim nCantidadTotal
    Dim nImporteTotal
    Dim rstProd
    
    c = ""
    c = c & Ambiente.Empresa & vbCrLf
    c = c & Ambiente.Direccion1 & vbCrLf
    c = c & Ambiente.Direccion2 & vbCrLf
    c = c & Ambiente.Telefonos & vbCrLf
    'c = c & Ambiente.Web & vbCrLf
    c = c & Ambiente.Uid & " " & Formato( Date(), "dd-MM-yyyy" ) & " " & Formato( Time(), "hh:mm:ss" ) & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf

    Set rstEntrada = CreaRecordSet( "SELECT * FROM entradas WHERE entrada = " & Documento, Ambiente.Connection )

    If rstEntrada.EOF Then
       MsgBox "El dato no existe", vbInformation
       Exit Sub
    End If    
   
    c = c & "Número de entrada: " & rstEntrada("entrada") & vbCrLf
    c = c & "Fecha: " & Formato( rstEntrada("f_emision"), "dd-MM-yyyy" ) & vbCrLf  
    c = c & vbCrLf
    c = c & vbCrLf

    Set rstPartidas = CreaRecordSet( "SELECT * FROM entpart WHERE entrada = " & Documento, Ambiente.Connection )

    While Not rstPartidas.EOF 

          Set rstProd = CreaRecordSet( "SELECT prods.descrip, (prods.precio1*(case when impuesto = 'IVA' then 1.16 else 1 end)) as precio FROM prods WHERE articulo = '" & rstPartidas("articulo") & "'", Ambiente.Connection )
          cCantidad = PadL(Formato( rstPartidas( "cantidad" ), "##,##" ), 10 )
          'cPrecio = PadL(Formato( rstProd( "precio" ), "##,##0.00" ), 10 )
          'cImporte = PadL(Formato( rstPartidas( "cantidad" ) * Val2(rstProd( "precio" )), "##,##0.00" ), 10 )          
          'c = c & mid(Trim(rstProd("Descrip")),1,40) & vbCrLf
          'c = c & Trim(rstPartidas("articulo")) & ", " & mid(Trim(rstProd("Descrip")),1,40) & vbCrLf
          'c = c & cCantidad & cPrecio & cImporte & vbCrLf
          'c = c & vbCrLf

          c = c & mid(Trim(rstProd("Descrip")),1,40) & vbCrLf
          c = c & cCantidad & vbCrLf
          
          nCantidadTotal = nCantidadTotal + rstPartidas( "cantidad" )
          nImporteTotal = nImporteTotal + (rstPartidas("cantidad") * Val2(rstProd( "precio" )))

          rstPartidas.MoveNext
    Wend

    'c = c & vbCrLf
    'c = c & vbCrLf
    c = c & vbCrLf
    c = c & "-------------------------------------------" & vbCrLf    
    c = c & "Cantidad total: " & Formato( nCantidadTotal, "##,##" ) & vbCrLf    
    c = c & "total: " & Formato( nImporteTotal, "##,##0.00" ) & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf
    c = c & vbCrLf

    if Ambiente.rstEstacion("ticketcorte") <> 0 Then
       c = c & Chr(27) & Chr(105)        
    end if

    Script.sendToPrinter Ambiente, (c), prn.Pantalla 
    c = ""

End Sub




Public Sub ImprimeSalida()

' Formato de entradas al inventario
' Por() Daniel

    Dim rstEncabezado  'El recordSet de los datos generales
    Dim rstPartidas    'Las partidas de la factura
    Dim rstTipoMovim   'El cliente de la venta
    Dim rstLotes       'Lista de lotes 
    Dim rstSeries      'Lista de números de serie
    Dim nRenglon       'El número de renglón sobre el que se va a trabajar    
    Dim strSql         'Una cadena con el query que se va a ejecutar
    Dim Estado         'El estado de la venta cancelado o facturado
    Dim ImportePartida 'El importe de cada partida
    Dim ImporteTotal   'El importe total de la factura
    Dim Impuesto       'El impuesto de la factura
    Dim rstMoneda      'La moneda de la venta 
    Dim nLineas        'Un contador para los campos de tipo memo 

    ' Creamos el recordSet del encabezado de la venta
    Set rstEncabezado = Rst("SELECT * FROM entradas WHERE entrada = " & prn.Documento, Ambiente.Connection )

    ' Verificamos que la venta que se desea imprimir exista
    if rstEncabezado.EOF Then
       MsgBox "No existe la entrada seleccionada",vbInformation
       Exit Sub
    end if 

    ' Creamos el recordSet de las partidas que componen la entrada
    Set rstPartidas = Rst("SELECT * FROM entpart WHERE entrada = " & rstEncabezado.fields("Entrada"), Ambiente.Connection )

    ' Traemos todos los datos del proveedor
    Set rstTipoMovim = Rst("SELECT * FROM tipominv WHERE tipo_movim = '" & rstEncabezado.fields("Tipo_doc") & "'", Ambiente.Connection )

    ' Iniciamos la Impresión del encabezado de la forma
    IniciaDocumento
    'Encabezado
    Call Encabezado()
    nRenglon = 12

    EstableceFuente "Courier New", 8

    While Not rstPartidas.EOF

          ' No congelamos el sistema
          Ev = DoEvents          
          if nRenglon > 60 Then
              PaginaNueva
              Call Encabezado()              
              nRenglon = 12
           end if


          Say Row(nRenglon), Col(06), Trim(rstPartidas.Fields("Articulo"))
          Say Row(nRenglon), Col(43), Formato(rstPartidas.Fields("Precio"), "###,###,###.00")
          Say Row(nRenglon), Col(52), rstPartidas.Fields("Cantidad")

          ' Calculamos el importe
          ImportePartida = rstPartidas.Fields("Precio") * rstPartidas.Fields("Cantidad") 
          ' Calculamos el total
          ImporteTotal = ImporteTotal + ImportePartida

          Say Row(nRenglon), Col(67), Formato(ImportePartida,"###,###,###.00")

          ' Contamos las lineas que tiene el campo memo
          ' Imprimimos las lineas al final para poder incrementar los renglones
          ' sin afectar a las partidas
          nLineas = CuantasLineas( rstPartidas.Fields("Observ") )
          
          ' Imprimimos cada una de ellas
          For n = 1 to nLineas 
              Say Row(nRenglon), Col(16), strLinea( Trim(rstPartidas.Fields("Observ")), n)
              nRenglon = nRenglon + 1
          Next

          if nRenglon > 60 Then
              PaginaNueva
              Call Encabezado()              
              nRenglon = 12
           end if

          ' Si no existen comentarios no disminuimos en un renglon.
          if nLineas > 0 Then
             nRenglon = nRenglon - 1
          end if

          rstPartidas.MoveNext
          nRenglon = nRenglon + 1

    Wend
    
    if nRenglon > 60 Then
       PaginaNueva
       Call Encabezado()              
       nRenglon = 12
    end if


    ' Imprimimos las observaciones de la operacion
    nLineas = CuantasLineas( rstEncabezado.Fields("Observ") )
          
    ' Imprimimos cada una de ellas
    For n = 1 to nLineas 
        nRenglon = nRenglon + 1
        Say Row(nRenglon), Col(16), strLinea( rstEncabezado.Fields("Observ"), n)
    Next

    if nRenglon > 55 Then
       PaginaNueva
       Call Encabezado()              
       nRenglon = 12
    end if

    Set rstMoneda = rst("SELECT * FROM monedas WHERE moneda = '" & Ambiente.Moneda & "'", Ambiente.Connection )
    nRenglon = nRenglon + 1
    EstableceFuente "Courier New", 10
    Linea Col(6), Row(nRenglon), Col(84), Row(nRenglon)
    Say Row(nRenglon), Col(06), Letra(ImporteTotal, rstMoneda.fields("Descrip"), True,rstMoneda.fields("Nombre") )
    Say Row(nRenglon+1), Col(55), "Importe : " & PadL(Formato(ImporteTotal,"###,###,###.00"),12)
    Say Row(nRenglon+2), Col(55), "Total   : " & PadL(Formato(ImporteTotal + Impuesto,"###,###,###.00"),12)

    ' Se da por terminado el documento
    FinDocumento    

End Sub     

Public Sub Encabezado()
' Metemos en una función el encabezado del reporte
   Dim rstEncabezado
   Dim rstTipoMovim

 ' Creamos el recordSet del encabezado de la venta
    Set rstEncabezado = Rst("SELECT * FROM ENTRADAS WHERE ENTRADA = " & prn.Documento, Ambiente.Connection )

' Traemos todos los datos del proveedor
    Set rstTipoMovim = Rst("SELECT * FROM TIPOMINV WHERE TIPO_MOVIM = '" & rstEncabezado.fields("Tipo_doc") & "'", Ambiente.Connection )

  
  EstableceFuente "Courier New", 20
    FontBold True    

    ' Imprimimos el logo del programa
    Picture Ambiente.path & "\Images\Business.jpg", Col(6), Row(1), Col(7),Row(4)

    
    ' Imprimimos el Nombre de la empresa
    FilledBox 1.5,0.4,5,0.43, RGB( 200, 200, 200 )
    Say Row(3),Col(15), Ambiente.Empresa

    EstableceFuente "Courier New", 10
    Say Prn.Row(5), Col(15), "Entrada: "  & rstEncabezado.Fields("Entrada") 
    Say Row(5), Col(70), "Página: " & Prn.Paginas


    if rstEncabezado.Fields("Estado") = "CO" Then
       Estado = "CONFIRMADA"
    end if

    if rstEncabezado.Fields("Estado") = "PE" Then
       Estado = "PENDIENTE"
    end if

    FontBold False
 
    Say Row(06), Col(6), "Estado             : " & Estado & " " & Formato( rstEncabezado.Fields("F_Emision"), "dd/MM/yyyy" )
    Say Row(07), Col(6), "Tipo de movimiento : " & rstEncabezado.Fields("Tipo_doc") & " " & rstTipoMovim.Fields("Descrip") & " " & rstEncabezado.Fields("almacen")

    Linea Col(6), Row(9), Col(74), Row(9)
    Say Row(10), Col(6), "Articulo          Descripción              Costo    Cantidad              Importe"
    Linea Col(6), Row(11), Col(74), Row(11)
    EstableceFuente "Courier New", 8

End Sub


Sub Imprimir( c )

    If Ambiente.rstEstacion("ticket") <> 0 Then
       if clAt( "LPT", Ambiente.rstEstacion("pticket") ) > 0 Then
          Out Trim(Ambiente.rstEstacion("pticket")), c
       else
          if Ambiente.Ticket.PortOpen Then
             Ambiente.Ticket.Output = c
          end if 
       end if
    End if

End Sub
