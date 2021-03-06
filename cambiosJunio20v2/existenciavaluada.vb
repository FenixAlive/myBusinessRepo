Public Sub Main()

    ' Colocamos los datos del rango
    ParamData.ParametrosRequeridos "Articulos", "Artículo inicial", "Articulos", "Artículo final"
    ParamData.Check = True    
    ParamData.CheckValue = True
    ParamData.CheckLeyend = "Solo con existencia positiva o negativa"

    ' Mostramos la ventana de rangos
    Rangos Ambiente, False

    ' Si se presiono el boton cancelado detenemos la operación
    if Cancelado Then
       Exit Sub
    end if           

    Reporte.Titulo = "Existencia Valuada"

    cCondicion = ""

    if Not ParamData.Todos1 Then 
       cCondicion = cCondicion & " AND (cast(prods.articulo as bigint) >= cast(" & ParamData.BusquedaIni & " as bigint) AND cast(prods.articulo as bigint) <= cast(" & ParamData.busquedaFin & " as bigint))"
       Reporte.Titulo2 = "Articulos entre el siguiente rango " & Trim(ParamData.BusquedaIni) & " - " & Trim(ParamData.BusquedaFin)
    End if

    if ParamData.CheckValue Then
       Reporte.Titulo2 = Reporte.Titulo2 & " (Solo artículos con existencia)"
       cCondicion = cCondicion & " AND (existencia > 0 OR existencia < 0)"
    end if 

                                  
    IniciaDocumento()
          strSQL = "" 
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "prods.articulo, "
          strSQL = strSQL & "prods.descrip, "
          strSQL = strSQL & "prods.existencia, "
          strSQL = strSQL & "(prods.precio1*(case when impuesto = 'IVA' then 1.16 else 1 end)) as Precio, "  
		  strSQL = strSQL & "(prods.existencia * prods.precio1 * (case when impuesto = 'IVA' then 1.16 else 1 end)) as Total "
          strSQL = strSQL & "FROM prods "
          strSQL = strSQL & "WHERE prods.articulo <> 'SYS' "
          strSQL = strSQL & cCondicion
          strSQL = strSQL & "ORDER BY prods.descrip "
          Reporte.SQL = strSQL 
          Reporte.RetrieveColumns 
          
          Reporte.Columns("articulo").Titulo = "ARTICULO"

          Reporte.Columns("Descrip").Titulo = "DESCRIPCION"
          Reporte.Columns("Descrip").Ancho = 37

          Reporte.Columns("existencia").Titulo = "Existencia"
          Reporte.Columns("existencia").Formato = Ambiente.FDinero
          Reporte.Columns("existencia").Font = "Terminal"
          Reporte.Columns("existencia").FontSize = 7
          Reporte.Columns("existencia").Ancho = 9
          Reporte.Columns("existencia").AnchoCelda = 8
          Reporte.Columns("existencia").AnchoTitulo = 7
          Reporte.Columns("existencia").Align = 1
          
		  Reporte.Columns("Precio").Titulo = "Precio"
          Reporte.Columns("Precio").Formato = Ambiente.FDinero
          Reporte.Columns("Precio").Font = "Terminal"
          Reporte.Columns("Precio").FontSize = 7
          Reporte.Columns("Precio").Ancho = 9
          Reporte.Columns("Precio").AnchoCelda = 8
          Reporte.Columns("Precio").AnchoTitulo = 7
          Reporte.Columns("Precio").Align = 1 
           
		  Reporte.Columns("Total").Acumulado = True
          Reporte.Columns("Total").Titulo = "Total"
          Reporte.Columns("Total").Formato = Ambiente.FDinero
          Reporte.Columns("Total").Font = "Terminal"
          Reporte.Columns("Total").FontSize = 7
          Reporte.Columns("Total").Ancho = 13
          Reporte.Columns("Total").AnchoCelda = 8
          Reporte.Columns("Total").AnchoTitulo = 7
          Reporte.Columns("Total").Align = 1 

          Reporte.ImprimeReporte
    FinDocumento()

End Sub

