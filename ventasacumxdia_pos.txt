' Elaborado por Luis Angel Muñoz Franco
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
       cCondicion = " AND ventas.f_emision >= " & FechaSQL(ParamData.FechaInicial, Ambiente.Connection) & " AND ventas.f_emision <= " & FechaSQL( ParamData.FechaFinal, Ambiente.Connection )
       Reporte.Titulo2 = "Del día " & Formato(ParamData.FechaInicial, "dd-MM-yyyy") & " al día " & Formato(ParamData.FechaFinal, "dd-MM-yyyy")
    End if

    IniciaDocumento()  
		strSQL = ""    
		strSQL = strSQL & "SELECT f_emision as Fecha, "
		strSQL = strSQL & "SUM( importe * tipo_cam * "   
        strSQL = strSQL & "CASE "
        strSQL = strSQL & "WHEN tipo_doc = 'DEV' THEN -1 "
        strSQL = strSQL & "ELSE 1 "      
        strSQL = strSQL & "END "
		strSQL = strSQL & ") as SubTotal, "
		strSQL = strSQL & "sum(impuesto*tipo_cam * "  
        strSQL = strSQL & "CASE "
        strSQL = strSQL & "WHEN tipo_doc = 'DEV' THEN -1 "
        strSQL = strSQL & "ELSE 1 "      
        strSQL = strSQL & "END "
		strSQL = strSQL & ") as IVA, " 
		strSQL = strSQL & "sum(iespecial*tipo_cam * "
        strSQL = strSQL & "CASE "
        strSQL = strSQL & "WHEN tipo_doc = 'DEV' THEN -1 "
        strSQL = strSQL & "ELSE 1 "      
        strSQL = strSQL & "END "     
        strSQL = strSQL & ") as IEPS, "    
		' ver como hacer una función para evitar tanto codigo
		strSQL = strSQL & "sum((importe+impuesto+iespecial) * tipo_cam* "
		strSQL = strSQL & "CASE "
        strSQL = strSQL & "WHEN tipo_doc = 'DEV' THEN -1 "
        strSQL = strSQL & "ELSE 1 "      
        strSQL = strSQL & "END "
		strSQL = strSQL & ") as Total "
		strSQL = strSQL & "FROM ventas "
		strSQL = strSQL & "WHERE (estado = 'CO' Or estado = 'CA') "
		strSQL = strSQL & cCondicion
        strSQL = strSQL & " group by f_emision "
        strSQL = strSQL & "order by f_emision;"
        Reporte.SQL = strSQL
		Reporte.Titulo = "Ventas en Caja por Dia"
        Reporte.RetrieveColumns  
        
		Reporte.Columns("fecha").Ancho = 13
		Reporte.Columns("fecha").Formato = "dd-MM-yyyy"	  

          Reporte.Columns("SubTotal").Acumulado = True
          Reporte.Columns("SubTotal").Formato = Ambiente.FDinero
          Reporte.Columns("SubTotal").Font = "Courier New"           
          Reporte.Columns("SubTotal").FontSize = 7           
          Reporte.Columns("SubTotal").Anchocelda = 13
          Reporte.Columns("SubTotal").Ancho = 13
          Reporte.Columns("SubTotal").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

          Reporte.Columns("IVA").Acumulado = True
          Reporte.Columns("IVA").Formato = Ambiente.FDinero
          Reporte.Columns("IVA").Font = "Courier New"
          Reporte.Columns("IVA").FontSize = 7
          Reporte.Columns("IVA").Anchocelda = 13
          Reporte.Columns("IVA").Ancho = 13
          Reporte.Columns("IVA").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

          Reporte.Columns("IEPS").Acumulado = True
          Reporte.Columns("IEPS").Formato = Ambiente.FDinero
          Reporte.Columns("IEPS").Font = "Courier New"
          Reporte.Columns("IEPS").FontSize = 7
          Reporte.Columns("IEPS").Anchocelda = 13
          Reporte.Columns("IEPS").Ancho = 13
          Reporte.Columns("IEPS").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

          Reporte.Columns("total").Acumulado = True
          Reporte.Columns("total").Formato = Ambiente.FDinero
          Reporte.Columns("total").Font = "Courier New"
          Reporte.Columns("total").FontSize = 7      
          Reporte.Columns("total").Anchocelda = 13     
          Reporte.Columns("total").Ancho = 13
          Reporte.Columns("total").Econdition = Prepara("ResultSet('ESTATUS') = 'CO'")

        Reporte.ImprimeReporte

	FinDocumento()
End Sub   