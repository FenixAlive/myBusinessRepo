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
       cCondicion = " AND salidas.f_emision >= " & FechaSQL(ParamData.FechaInicial, Ambiente.Connection) & " AND salidas.f_emision <= " & FechaSQL( ParamData.FechaFinal, Ambiente.Connection )
       Reporte.Titulo2 = "Del día " & Formato(ParamData.FechaInicial, "dd-MM-yyyy") & " al día " & Formato(ParamData.FechaFinal, "dd-MM-yyyy")
    End if

    IniciaDocumento()

          strSQL = "" 
          strSQL = strSQL & "SELECT "
          strSQL = strSQL & "salidas.tipo_doc AS 'DOCUMENTO', "
          strSQL = strSQL & "salidas.salida AS 'NUMERO', "
          strSQL = strSQL & "salidas.f_emision AS 'FECHA', "
          strSQL = strSQL & "salidas.estado AS 'ESTATUS', "
          strSQL = strSQL & "salpart.articulo AS 'articulo', "
          strSQL = strSQL & "salpart.cantidad AS 'cantidad', "
          strSQL = strSQL & "salpart.costo AS 'costo', "
          strSQL = strSQL & "salpart.costo * salpart.cantidad AS 'Valor' "
          strSQL = strSQL & "FROM salidas INNER JOIN salpart ON salidas.salida = salpart.salida "
          strSQL = strSQL & "WHERE (salidas.estado = 'CO' Or salidas.estado = 'CA') "
          strSQL = strSQL & cCondicion
          strSQL = strSQL & "ORDER BY salidas.salida "
          Reporte.SQL = strSQL 
          Reporte.Titulo = "Salidas al inventario"
          Reporte.RetrieveColumns 

          For n = 1 to Reporte.Columns.Count
              Reporte.Columns(n).EColor = RGB( 100,100,100 )
              Reporte.Columns(n).EColorCondition = Prepara("ResultSet('Estatus') = 'CA'")
          Next

          Reporte.Columns("cantidad").Acumulado = True
          Reporte.Columns("cantidad").Formato = Ambiente.FDinero
          Reporte.Columns("cantidad").Font = "Courier New"
          Reporte.Columns("cantidad").FontSize = 7      
          Reporte.Columns("cantidad").Anchocelda = 11     
          Reporte.Columns("cantidad").Ancho = 8
          Reporte.Columns("cantidad").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")


          Reporte.Columns("costo").Acumulado = True
          Reporte.Columns("costo").Formato = Ambiente.FDinero
          Reporte.Columns("costo").Font = "Courier New"
          Reporte.Columns("costo").FontSize = 7      
          Reporte.Columns("costo").Anchocelda = 11     
          Reporte.Columns("costo").Ancho = 8
          Reporte.Columns("costo").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")


          Reporte.Columns("valor").Acumulado = True
          Reporte.Columns("valor").Formato = Ambiente.FDinero
          Reporte.Columns("valor").Font = "Courier New"
          Reporte.Columns("valor").FontSize = 7      
          Reporte.Columns("valor").Anchocelda = 11     
          Reporte.Columns("valor").Ancho = 8
          Reporte.Columns("valor").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

          Reporte.Columns("fecha").Formato = "dd-MM-yyyy"

          Reporte.Columns("Estatus").Titulo = "Est."
          Reporte.Columns("Estatus").Ancho = 5

          Reporte.ImprimeReporte

    FinDocumento()

End Sub