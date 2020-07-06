Public Sub Main()

    ' Colocamos los datos del rango
    ParamData.ParametrosRequeridos ,,,,,,,,True 

    ' Mostramos la ventana de rangos
    Rangos Ambiente, False

    ' Si se presiono el boton cancelado detenemos la operación
    if Cancelado Then
       Exit Sub
    end if

    if Not ParamData.TodasLasFechas Then
       cCondicion = " AND entradas.f_emision >= " & FechaSQL(ParamData.FechaInicial, Ambiente.Connection) & " AND entradas.f_emision <= " & FechaSQL( ParamData.FechaFinal, Ambiente.Connection )
       Reporte.Titulo2 = "Del día " & Formato(ParamData.FechaInicial, "dd-MM-yyyy") & " al día " & Formato(ParamData.FechaFinal, "dd-MM-yyyy")
    End if

    IniciaDocumento()

          strSQL = "" 
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "entradas.tipo_doc AS 'DOCUMENTO', "
          strSQL = strSQL & "entradas.entrada AS 'NUMERO', "
          strSQL = strSQL & "entradas.f_emision AS 'FECHA', "
          strSQL = strSQL & "entradas.estado AS 'ESTATUS', "
          strSQL = strSQL & "entpart.articulo AS 'articulo', "
		  strSQL = strSQL & "prods.descrip AS 'descripción', "
          strSQL = strSQL & "entpart.cantidad AS 'cantidad', "
          strSQL = strSQL & "(prods.precio1*(case when impuesto = 'IVA' then 1.16 else 1 end)) as 'Precio', "
          strSQL = strSQL & "(prods.precio1*(case when impuesto = 'IVA' then 1.16 else 1 end)) * entpart.cantidad AS 'Importe' "
          strSQL = strSQL & "FROM entpart INNER JOIN entradas ON entpart.entrada = entradas.entrada "
		  strSQL = strSQL & "INNER JOIN prods ON entpart.articulo = prods.articulo "
          strSQL = strSQL & "WHERE (entradas.estado = 'CO' Or entradas.estado = 'CA') "
          strSQL = strSQL & cCondicion
          strSQL = strSQL & "ORDER BY entradas.entrada, prods.descrip "
          Reporte.SQL = strSQL 
          Reporte.Titulo = "Entradas al inventario"
          Reporte.RetrieveColumns 

          For n = 1 to Reporte.Columns.Count
              Reporte.Columns(n).EColor = RGB( 100,100,100 )
              Reporte.Columns(n).EColorCondition = Prepara("ResultSet('Estatus') = 'CA'")
          Next

		  Reporte.Columns("NUMERO").Grupo = True
          Reporte.Columns("NUMERO").GrupoTitulo = "Numero Documento: "
          Reporte.Columns("NUMERO").GrupoData = Prepara("ResultSet('Numero')")
          Reporte.Columns("NUMERO").GrupoTotales = True
          Reporte.Columns("NUMERO").Visible = False  
          Reporte.Columns("NUMERO").GrupoTotalLeyenda = "Total por Documento: "
          Reporte.Columns("NUMERO").Ancho = 1

          Reporte.Columns("DOCUMENTO").Grupo = True
          Reporte.Columns("DOCUMENTO").GrupoTitulo = "Tipo Documento: "
          Reporte.Columns("DOCUMENTO").GrupoTotales = False
          Reporte.Columns("DOCUMENTO").Visible = False  
          Reporte.Columns("DOCUMENTO").Ancho = 1


		  Reporte.Columns("Fecha").Grupo = True
          Reporte.Columns("Fecha").GrupoTitulo = "Fecha: "
          Reporte.Columns("Fecha").GrupoTotales = False
          Reporte.Columns("Fecha").Visible = False  
          Reporte.Columns("Fecha").Ancho = 1
		  'Reporte.Columns("Fecha").Formato = "dd-MM-yyyy"

		  Reporte.Columns("descripción").Ancho = 35
          Reporte.Columns("descripción").Data = Prepara("Mid(Resultset('descripción'),1,60)")
          Reporte.Columns("descripción").Font = "Courier New"           
          Reporte.Columns("descripción").FontSize = 7
		  Reporte.Columns("descripción").Visible = True
		
          Reporte.Columns("cantidad").Acumulado = True
          Reporte.Columns("cantidad").Formato = Ambiente.FDinero
          Reporte.Columns("cantidad").Font = "Courier New"
          Reporte.Columns("cantidad").FontSize = 7      
          Reporte.Columns("cantidad").Anchocelda = 11     
          Reporte.Columns("cantidad").Ancho = 8
          Reporte.Columns("cantidad").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")


          Reporte.Columns("Precio").Formato = Ambiente.FDinero
          Reporte.Columns("Precio").Font = "Courier New"
          Reporte.Columns("Precio").FontSize = 7      
          Reporte.Columns("Precio").Anchocelda = 11     
          Reporte.Columns("Precio").Ancho = 8
          Reporte.Columns("Precio").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")


          Reporte.Columns("Importe").Acumulado = True
          Reporte.Columns("Importe").Formato = Ambiente.FDinero
          Reporte.Columns("Importe").Font = "Courier New"
          Reporte.Columns("Importe").FontSize = 7      
          Reporte.Columns("Importe").Anchocelda = 11     
          Reporte.Columns("Importe").Ancho = 8
          Reporte.Columns("Importe").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

          Reporte.Columns("Estatus").Titulo = "Est."
          Reporte.Columns("Estatus").Ancho = 1 
		  Reporte.Columns("Estatus").Visible = False

          Reporte.ImprimeReporte

    FinDocumento()

End Sub

