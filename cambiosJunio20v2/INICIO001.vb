Sub Main()                                              
    Dim adXactReadSerializable
    Dim adXactReadUncommitted
    Dim adXactReadCommitted
    Dim adXactRepeatableRead

    adXactReadSerializable = 1048576
    adXactReadUncommitted = 256
    adXactReadCommitted = 4096
    adXactRepeatableRead = 65536                                    

    ' Establecemos el modo de Aislamiento
    'Ambiente.Connection.Execute "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"
    Ambiente.Connection.IsolationLevel = adXactReadSerializable 

    Parent.Timer.Enabled = True
    Ambiente.eventRowReports = True 
    Ambiente.ModoDeDepuracion = True                           
       
 
    ' Cargamos el listado de Iconos                                
    ModoAvanzado                
                                   
    ' Vemos si el usuario tiene pendientes                   
    'Call buscaPendientes


    'Set NodX = bmForm.tvTreeView.Nodes.Add("Business", 4, "CarpetaDeUsuario1", "Carpeta Personalizada", 1, 2)
     
    Set rstEConfig = CreaRecordSet( "SELECT * FROM econfig", Ambiente.Connection )
    
    Select Case Trim( Ucase( rstEConfig("pais") ) )
           Case "MEXICO"
                SetSessionValue Ambiente, "IMPUESTO_DEFAULT", "IVA" 
                SetSessionValue Ambiente, "PRODUCTOS_IMPUESTO_DEFAULT", "IVA" 
                SetSessionValue Ambiente, "PAIS", "Mexico" 
           Case "PANAMA"
                SetSessionValue Ambiente, "IMPUESTO_DEFAULT", "ITBMS"
                SetSessionValue Ambiente, "PRODUCTOS_IMPUESTO_DEFAULT", "ITO"
                SetSessionValue Ambiente, "PAIS", "Panama" 
           Case Else
                SetSessionValue Ambiente, "IMPUESTO_DEFAULT", "IVA" 
                SetSessionValue Ambiente, "PRODUCTOS_IMPUESTO_DEFAULT", "IVA" 
                SetSessionValue Ambiente, "PAIS", "Mexico" 
    End Select  
  
    '//--------------------- comienza impresión de ticket por entrada a administracion
If Ambiente.var9 = True Then

cLineaNueva = Chr(13) & Chr(10)
cSalida = ""                                              

'Encabezado
cSalida = cSalida & cLineaNueva & cLineaNueva 
'"CreaRecordSet()" consulta si el usuario colocó algun mensaje en el encabezado y pie del ticket a traves de la utileria "editar texto del ticket" y lo almacena en "rstTextTicket"
Set rstTextTicket = CreaRecordSet( "SELECT * FROM tickettext", Ambiente.Connection ) 

If rstTextTicket.EOF Then 
	cSalida = cSalida & "" & Ambiente.Empresa & cLineaNueva 
	cSalida = cSalida & "" & Trim( Ambiente.Direccion1 ) & cLineaNueva 
	cSalida = cSalida & "" & Trim( Ambiente.Direccion2 ) & cLineaNueva 
	cSalida = cSalida & "" & Trim( Ambiente.Telefonos ) & cLineaNueva 
Else
	cSalida = cSalida & rstTextTicket("textheader") 
End If

cSalida = cSalida & cLineaNueva & "  --------------------------------------------------" & cLineaNueva 

Set fecha = Rst("select CURRENT_TIMESTAMP as fecha", Ambiente.Connection)   
cSalida = cSalida & cLineaNueva & "        Ingreso de Cajero: "
cSalida = cSalida & cLineaNueva & cLineaNueva
cSalida = cSalida & "  Fecha y hora: " & Trim( fecha("fecha") )        
cSalida = cSalida & cLineaNueva & cLineaNueva     
cSalida = cSalida & "  Usuario: " & Trim( Ambiente.Uid )
cSalida = cSalida & cLineaNueva        

cSalida = cSalida & cLineaNueva & "  --------------------------------------------------" & cLineaNueva  

	Script.sendToPrinter Ambiente, (cSalida), prn.Pantalla
	'msgbox cSalida    
	'PrintText cSalida
    
	'guardar en base de datos
	Set rstId = CreaRecordSet( "SELECT TOP 1 id FROM asistencia ORDER BY id DESC", Ambiente.Connection )
	if rstId.EOF then
		execSQL "INSERT into asistencia values(cast(1 as int),CAST(CURRENT_TIMESTAMP AS nvarchar(19)), '"&Trim( Ambiente.Uid )&"', CURRENT_TIMESTAMP, null, null)"
 	else
		execSQL "INSERT into asistencia values(cast("&Val2(rstId(0))+1&" as int),CAST(CURRENT_TIMESTAMP AS nvarchar(19)), '"&Trim( Ambiente.Uid )&"', CURRENT_TIMESTAMP, null, null)"
	end if	
	Ambiente.var9 = False
End If                                  

'//-----------------------termina impresión de ticket por entrar a administracion

End Sub


Sub ModoAvanzado()

    Dim rstUsuario 

    On Error Resume Next    
    vbalListBar1.Width = Me.Width / 4
    On Error Goto 0

    ' Vemos si el usuario es supervisor
    Set rstUsuario = CreaRecordSet( "SELECT supervisor FROM usuarios WHERE usuario = '" & Ambiente.Uid & "'", Ambiente.Connection )

    'Me.vbalImageList1.AddFromFile Ambiente.Path & "\images\recostear.ico", 1, "recostear"
    'Eventos

    vbalListBar1.Bars.Clear       


    With vbalListBar1
                                                            
		 If parent.bIconosNuevos = True Then
		  	.ImageList = imgBarra
		 Else   
         	.ImageList = vbalImageList1
		 End If

         .Bars.Add ,"Ventas", "Ventas"
         '.Bars.Add ,"Restaurante", "Restaurante"

         .Bars.Add ,"Compras", "Compras"
         .Bars.Add ,"Inventario", "Inventario"        
		 .Bars.Add ,"Farmacia", "Farmacia"	

         .Bars.Add ,"Utilerias", "Utilerias" 
		 .Bars.Add ,"Mobile Business", "Mobile Business"

         Set rstEConfig = CreaRecordSet( "SELECT * FROM econfig", Ambiente.Connection )
    
         Select Case Trim( Ucase( rstEConfig("pais") ) )
                Case "MEXICO"
 		             .Bars.Add ,"FacElectronica", "Factura Electrónica" 
   		             .Bars.Add ,"TiempoAire", "Tiempo Aire"     
                Case "PANAMA"
 
                Case Else
 		             .Bars.Add ,"FacElectronica", "Factura Electrónica" 
   		             .Bars.Add ,"TiempoAire", "Tiempo Aire"     
         End Select  

         If rstUsuario("supervisor") <> 0 Then
            ' Le damos acceso a la parte de programación solo si es supervisor
            .Bars.Add ,"Configuracion", "Configuración"
            .Bars.Add ,"Programacion", "Programación"

            '.Bars("Configuracion").Items.Add ,"DatosGenerales", "Conexión a la base de datos", 100            
            .Bars("Configuracion").Items.Add ,"ConfigGeneral", "Datos generales de la empresa", 109
         End If
 
         For n = 1 To .Bars.Count
             .Bars(n).State = 0
         Next                  

         .Bars("ventas").State = 1
                         

         '.Bars.Add ,"Produccion", "Producción" 
 

         ' Bars.Add agrega una pestaña a la barra de tareas
         ' Opcional 1.- El indice de la barra, 2.- La llave 
         ' de la barra, 3.- El titulo de la barra
         '.Bars.Add ,"cursoveracruz", "Curso Veracruz"
         
         ' Items.Add agrega un icono a la pestaña indicada
         ' Opcional 1.- El indice, 2.- La llave del icono,
         ' 3.- Titulo del icono, 4.- Un indice que representa
         ' un icono entre 1 y 140 
         '.Bars("cursoveracruz").Items.Add ,"existenciaremota", _
         '"Existencias en sucursales", 110


         .Bars("Inventario").Items.Add ,"Articulos", "Artículos", 2
         .Bars("Inventario").Items.Add ,"AltaRapida", "Alta Rápida (Artículos)", 45                           
         .Bars("Inventario").Items.Add ,"Tallas", "Tallas y Colores por Modelo", 19                           
         .Bars("Inventario").Items.Add ,"InvExcel", "Importar artículos y existencia desde excel(tm)", 98
         .Bars("Inventario").Items.Add ,"Entradas", "Entradas al Inventario", 25
         .Bars("Inventario").Items.Add ,"Salidas", "Salidas al Inventario", 26
         .Bars("Inventario").Items.Add ,"Fisico", "Inventario físico", 27
         .Bars("Inventario").Items.Add ,"Recalcular", "Recalcular Inventario (Promedio)", 29

         .Bars("Inventario").Items.Add ,"UEPS", "Recalcular Costeo (UEPS)", 130
         .Bars("Inventario").Items.Add ,"PEPS", "Recalcular Costeo (PEPS)", 128

         .Bars("Inventario").Items.Add ,"HojaInventario", "Resumen de operaciones", 65

         '.Bars("Inventario").Items.Add ,"Recostear", "Recostear Inventario", 99

         '.Bars("Inventario").Items.Add ,"PedidosOrdenes", "Recalcular por surtir y por recibir", 35
         '.Bars("Inventario").Items.Add ,"Cerrar", "Cerrar inventario", 30 

         'Rosy 08-Noviembre-2006 
         'If clAt("MySQL",Ambiente.Connection.ConnectionString) Then       
	         .Bars("Inventario").Items.Add ,"CalidadInventario", "Informe de Calidad de Inventario", 34
	         '.Bars("Inventario").Items.Add ,"CalidadInventario2", "Informe de Calidad de Inventario (Solo nuevos)", 62 
         'End If                 
         'Rosy

         .Bars("Inventario").Items.Add ,"series", "Números de serie", 101
         .Bars("Inventario").Items.Add ,"lotes", "Lotes", 51 
         .Bars("Inventario").Items.Add ,"Analisis", "Analisis de Inventario", 121
		' Cambia la clave de un articulo en todas sus tablas
		 .Bars("Inventario").Items.Add ,"CambiaClaveProd", "Cambiar Clave Producto", 3
		' Alta, modificación y baja de descuentos especiales
		 .Bars("Inventario").Items.Add ,"DescuentosEspeciales", "Descuentos Especiales", 4

         .Bars("Ventas").Items.Add ,"PuntoVenta", "Punto de venta", 7
         .Bars("Ventas").Items.Add ,"CierreTienda", "Resumen general de operaciones", 136         
         .Bars("Ventas").Items.Add ,"RemisionFactura", "Convertir remisiones a factura", 137         

         .Bars("Ventas").Items.Add ,"Pedidos", "Pedidos de clientes", 4
         .Bars("Ventas").Items.Add ,"Facturas", "Facturas", 5
         .Bars("Ventas").Items.Add ,"Remisiones", "Remisiones", 6
         .Bars("Ventas").Items.Add ,"Clientes2", "Clientes", 16      
		 .Bars("Ventas").Items.Add ,"ImpClientes", "Importar catálogo de clientes desde Excel(tm)", 130
         .Bars("Ventas").Items.Add ,"Cobranza", "Cobranza", 10
         .Bars("Ventas").Items.Add ,"Devoluciones", "Devoluciones / Notas de crédito", 22
         
         .Bars("Compras").Items.Add ,"OrdenesDeCompra", "Ordenes de Compra", 8
         .Bars("Compras").Items.Add ,"Compras", "Compra", 9
         .Bars("Compras").Items.Add ,"DevComp", "Devolución de compra", 24
         .Bars("Compras").Items.Add ,"Proveedores2", "Proveedores", 17
		 .Bars("Compras").Items.Add ,"ImpProveedores", "Importar catálogo de proveedores desde Excel(tm)", 130
         .Bars("Compras").Items.Add ,"Cxp", "Cuentas por pagar", 11

		 .Bars("Farmacia").Items.Add ,"FrmDoctores", "Mantenimiento de Doctores", 135
		 .Bars("Farmacia").Items.Add ,"FrmAntibioticos", "Mantenimiento de Antibioticos", 136
         .Bars("Farmacia").Items.Add ,"VentasXDia", "Reporte de Ventas por Día", 137
		 .Bars("Farmacia").Items.Add ,"ListaAntibioticos", "Reporte Lista de Antibioticos", 138
		 .Bars("Farmacia").Items.Add ,"KardexAntibioticos", "Reporte Kardex de Antibioticos", 139
         .Bars("Farmacia").Items.Add ,"Existencia", "Reporte Existencias", 140
         
         .Bars("Utilerias").Items.Add ,"EditarTicket", "Editar texto del ticket", 96 

         .Bars("Utilerias").Items.Add ,"Etiquetas", "Etiquetas de código de barras", 39
         .Bars("Utilerias").Items.Add ,"FPersonalizados", "Formatos personalizados (Requiere Microsoft Word)", 140

         .Bars("Utilerias").Items.Add ,"Ofertas", "Ofertas", 134
         .Bars("Utilerias").Items.Add ,"CambioPrecio", "Cambios de precios", 65
         .Bars("Utilerias").Items.Add ,"Bitacoras", "Bitacoras", 41
         .Bars("Utilerias").Items.Add ,"SubirBitacora", "Subir Bitacoras", 98
         '.Bars("Utilerias").Items.Add ,"Telemercadeo", "Telemercadeo", 13
         '.Bars("Utilerias").Items.Add ,"Pendientes", "Pendientes", 14
         
         Select Case Trim( Ucase( rstEConfig("pais") ) )
                Case "MEXICO"  
         			 .Bars("FacElectronica").Items.Add ,"ConfigSucursal", "Configuración de Sucursales", 102
         			 .Bars("FacElectronica").Items.Add ,"FacturaElectronica", "Emisión de Facturas Electrónicas", 102
                     '.Bars("FacElectronica").Items.Add ,"CFD", "Datos para factura Electronica (Medios Propios o CFDI para validar)", 131
                     '.Bars("FacElectronica").Items.Add ,"CFDGEN", "Facturas electrónicas", 141
                     '.Bars("FacElectronica").Items.Add ,"CFDPAPEL1", "Diseñar formato de facturas, carta", 143
                     '.Bars("FacElectronica").Items.Add ,"CFDPAPEL2", "Diseñar formato de facturas, ticket", 142  

 		             .Bars("TiempoAire").Items.Add ,"ActualizarMontos", "Actualizar montos para venta de tiempo aire", 69


                Case "PANAMA"
                   .Bars("Utilerias").Items.Add ,"EditarFacturas", "Editar formato de factura", 7
                   .Bars("Utilerias").Items.Add ,"EditarRemisiones", "Editar formato de remisiones", 7
                   .Bars("Utilerias").Items.Add ,"EditarNc", "Editar formato de Notas de Crédito", 7
               Case Else 

                     .Bars("FacElectronica").Items.Add ,"ConfigSucursal", "Configuración de Sucursales", 102
         			 .Bars("FacElectronica").Items.Add ,"FacturaElectronica", "Emisión de Facturas Electrónicas", 102
                     '.Bars("FacElectronica").Items.Add ,"CFD", "Datos para factura Electronica (Medios Propios o CFDI para validar)", 131
                     '.Bars("FacElectronica").Items.Add ,"CFDGEN", "Facturas electrónicas", 141
                     '.Bars("FacElectronica").Items.Add ,"CFDPAPEL1", "Diseñar formato de facturas, carta", 143
                     '.Bars("FacElectronica").Items.Add ,"CFDPAPEL2", "Diseñar formato de facturas, ticket", 142  
                      
 		             .Bars("TiempoAire").Items.Add ,"ActualizarMontos", "Actualizar montos para venta de tiempo aire", 69



         End Select  
          
 
         '.Bars("Utilerias").Items.Add ,"EditarClientes", "Cartas personalizadas", 7

         '.Bars("Utilerias").Items.Add ,"Pls2001", "Generar PLU Para Torrey PLS 2001", 94
         .Bars("Utilerias").Items.Add ,"Verificador", "Verificador", 36
         .Bars("Utilerias").Items.Add ,"Backup", "Respaldo", 142
         .Bars("Utilerias").Items.Add ,"Huella", "Registro de entradas/Salidas de personal", 144
         .Bars("Utilerias").Items.Add ,"MBInventario", "MyBusiness Inventario", 129

         '.Bars("Produccion").Items.Add ,"ordenproduccion", "Orden de producción", 119
         '.Bars("Produccion").Items.Add ,"capturaetiquetas", "Captura de etiquetas", 120
         '.Bars("Produccion").Items.Add ,"seguimiento", "Seguimiento del proceso", 121 
         '.Bars("Produccion").Items.Add ,"destajo", "Pago del destajo", 122                              

         '.Bars("FacElectronica").Items.Add ,"DatosAdicionales", "Datos adicionales de la empresa", 109


		 .Bars("Mobile Business").Items.Add ,"configpocket", "Configurar Pocket PC", 48
		 .Bars("Mobile Business").Items.Add ,"inventariopocket", "Capturar Inventario de la Pocket PC", 2
         .Bars("Mobile Business").Items.Add ,"pocket", "Exportar/Importar datos a Pocket PC", 71 
		 '.Bars("Mobile Business").Items.Add ,"confirmaventas", "Confirmar ventas de Pocket PC", 5 
		 .Bars("Mobile Business").Items.Add ,"creanotas", "Asistente de devoluciones en Pocket PC", 6
		 '.Bars("Mobile Business").Items.Add ,"creainventariofisico", "Aplicar Inventario Físico de Pocket PC", 27
		 .Bars("Mobile Business").Items.Add ,"reportepedidospk", "Reporte de Pedidos de la Pocket PC", 121
		 .Bars("Mobile Business").Items.Add ,"cargainicial", "Reporte de productos cargados inicialmente a la Pocket PC", 22
		 .Bars("Mobile Business").Items.Add ,"ventaspendientes", "Reporte de ventas generadas por la Pocket PC", 4
		 .Bars("Mobile Business").Items.Add ,"articulosvendidos", "Reporte de artículos vendidos desde la Pocket PC", 108     

		If rstUsuario("supervisor") <> 0 Then

            If clAt( "MySQL", Ambiente.Connection ) Then
               .Bars("Configuracion").Items.Add ,"Mantenimiento", "Mantenimiento a la base de datos", 101
            End If
                                                                                              
            .Bars("Configuracion").Items.Add ,"FormatosImpresion", "Establecer Formatos de Impresión", 102
            .Bars("Configuracion").Items.Add ,"Consecutivos", "Consecutivos de impresión", 103
			.Bars("Configuracion").Items.Add ,"GeneralConsecutivos", "General de consecutivos de impresión", 117
            .Bars("Configuracion").Items.Add ,"BorrarBase", "Borrar Base de Datos", 104
            .Bars("Configuracion").Items.Add ,"AltaEmpresa", "Manejo de conexiones a base de datos", 105
            .Bars("Configuracion").Items.Add ,"CambiarEmpresa", "Cambiar de conexión", 106
            .Bars("Configuracion").Items.Add ,"SincronizarProcedimientos", "Sincronizar Procedimientos", 110
            .Bars("Configuracion").Items.Add ,"Rangos", "Rangos de vista del Business Manager", 111
            .Bars("Configuracion").Items.Add ,"RegistroDeLicencia", "Acerca de...", 107
            .Bars("Programacion").Items.Add ,"EditorFormas", "Ambiente de desarrollo", 42
            .Bars("Programacion").Items.Add ,"Busquedas", "Editor de busquedas", 18
            .Bars("Programacion").Items.Add ,"EditorSQL", "Editor de SQL", 19
            .Bars("Programacion").Items.Add ,"ImportarMySQL", "Importar de MySQL", 101 
         End If         

         '.Bars("Restaurante").Items.Add ,"Secciones", "Secciones", 137
         '.Bars("Restaurante").Items.Add ,"r_impresoras", "Impresoras", 143

         '.Bars("Restaurante").Items.Add ,"Touch", "Caja", 53
         '.Bars("Restaurante").Items.Add ,"Comandas", "Comandas", 136
         '.Bars("Restaurante").Items.Add ,"Reservaciones", "Reservaciones", 140
         '.Bars("Restaurante").Items.Add ,"Menu", "Definición de menú", 141


         If rstUsuario("supervisor") <> 0 Then 
            If clAt( "MySQL", Ambiente.Connection ) > 0 Then         
               .Bars("Utilerias").Items.Add ,"MySQLPass", "Establecer password de MySQL", 46
            End If
         End If  
          
    End With

End Sub


Sub buscaPendientes()
       
    Set rstPendientes = CreaRecordSet( "SELECT * FROM pendient WHERE para = '" & Ambiente.Uid & "' AND estado = 'PE' AND fecha = " & fechaSQL( fecha ), _
    Ambiente.Connection )  
                                                                                                                       
    If Not rstPendientes.EOF Then
       Set Clientes = CreateObject( "MyBClientes.Clientes" )
       Set Clientes.Ambiente = Me.Ambiente
       Clientes.MuestraPendientes
    End If

End Sub                      

Sub execSQL( strSQL )
 
    On Error Resume Next                              
    Ambiente.Connection.Execute (strSQL)
 
    If Err.Number <> 0 Then
       MyMessage (strSQL) & " --- " & Err.Description & vbCrLf
    End If

End Sub