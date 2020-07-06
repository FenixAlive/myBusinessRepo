Sub Main()

    ' Colocamos los datos del rango
    ParamData.ParametrosRequeridos ,,,,,,,,True

    ' Mostramos la ventana de rangos
    Rangos Ambiente, False

    ' Si se presiono el boton cancelado detenemos la operación
    if Cancelado Then
       Exit Sub
    end if

    cCondicion = ""
     
    if Not ParamData.TodasLasFechas Then
       cCondicion = cCondicion & " asistencia.fecha >= " & FechaSQL(ParamData.FechaInicial, Ambiente.Connection) & " AND asistencia.fecha <= " & FechaSQL(ParamData.FechaFinal, Ambiente.Connection)
       Reporte.Titulo2 = Reporte.Titulo2 & "Del día " & Formato(ParamData.FechaInicial, "dd-MM-yyyy") & " al día " & Formato(ParamData.FechaFinal, "dd-MM-yyyy")
    End if    

    IniciaDocumento
    Reporte.Titulo = "REPORTE DE ASISTENCIA DE EMPLEADOS"
   
    If clEmpty( (cCondicion) ) Then
       Reporte.SQL = "SELECT asistencia.empleado, usuarios.nombre, asistencia.fechahora FROM asistencia INNER JOIN usuarios ON asistencia.empleado = usuarios.usuario ORDER BY asistencia.fecha, usuarios.nombre "
    Else
       Reporte.SQL = "SELECT asistencia.empleado, usuarios.nombre, asistencia.fechahora FROM asistencia INNER JOIN usuarios ON asistencia.empleado = usuarios.usuario WHERE " & cCondicion & " ORDER BY asistencia.fecha, usuarios.nombre "
    End If

    Reporte.RetrieveColumns

    Reporte.Columns("nombre").Ancho = 25

    Reporte.Columns("fechahora").Titulo = "Entrada"
    Reporte.Columns("fechahora").Ancho = 13
 


    Reporte.ImprimeReporte 

    FinDocumento

End Sub