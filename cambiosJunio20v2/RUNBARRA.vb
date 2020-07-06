Sub Main()
    Dim Ventas    
                                       
    Select Case Ambiente.Tag    
           Case "Series"         
                Script.RunForm "SERIES", Parent, Ambiente,, False
           Case "Verificador"

                If Question( "Desea activar el verificador de precios (recuerde desactiva la ventana con ...1 y enter)", 1 ) Then
                   MuestraVerificador  
                End If
 
           Case "Reconectar"
                CreaConexiones                 
           Case "FrmDoctores"
                Script.RunForm "frmdoctores", Parent, Ambiente,, False
           Case "FrmAntibioticos"
                Script.RunForm "frmantibioticos", Parent, Ambiente,, False
		   Case "VentasXDia"
				EjecutaReporte "VENTASXDIA"
           Case "ListaAntibioticos"
                EjecutaReporte "ANTIBIOTICOS"
           Case "KardexAntibioticos"
                EjecutaReporte "KARDEXANTIBIOTICOS"
           Case "Existencia"
                EjecutaReporte "EXISTENCIA"
           Case "AltaRapida"
                Script.RunForm "ALTARAPIDA", Parent, Ambiente,, False
           Case "AltaHotel"
                Script.RunForm "ALTASERVICIOS", Parent, Ambiente,, False
           Case "MySQLPass"
                Script.RunForm "PASSMYSQL", Parent, Ambiente,, False
           Case "HotSync"
                Script.RunProcess "HOTSYNC", Parent, Ambiente
           Case "Ofertas"
                Script.RunForm "OFERTAS", Parent, Ambiente,, False
           Case "CambioPrecio"
                Script.RunForm "CAMBIOPRECIO", Parent, Ambiente,, False
           Case "Procedimientos"
                MuestraNodo "Procedimientos" 
           Case "Formatos"
                MuestraNodo "formatos" 
           Case "Reportes"
                MuestraNodo "Reportes" 
           Case "Hermes"
                Script.RunForm "OFERTAS2", Parent, Ambiente,, False
           Case "Etiquetas2"
                Script.RunForm "ETIQUETA2", Parent, Ambiente,, False
           Case "InvExcel"
                Script.RunForm "PRODSEXCEL", Parent, Ambiente,, False
           Case "estados"
                Script.RunForm "ESTADOS", Parent, Ambiente,, False
           Case "EditarCB" 
                MuestraDEtiquetas
           Case "EditarTicket"
                Script.RunForm "EDITORTICKET", Me, Ambiente,, False 
           Case "EditarGuion"
                Script.RunForm "EDITORMENSAJE", Me, Ambiente,, False 
           Case "Pls2001"
                Script.RunForm "EPLS2000", Me, Ambiente,, False 
           Case "Repartidor"
                Script.RunForm "ALTAREPARTIDOR", Me, Ambiente,, False 
           Case "ImpVentas"
                Script.RunForm "IMPVENTAS2", Me, Ambiente,, False 
           Case "Pacientes"
                Script.RunForm "PACIENTES", Me, Ambiente,, False       
           Case "Bitacoras"
                Script.RunForm "BITACORAS", Me, Ambiente,, False            
           Case "SubirBitacora"
                Script.RunForm "SUBIRBITACORAS", Me, Ambiente,, False            
           Case "prospectos"    
                ' Script puede ejecutar un proceso, una forma de captura, un reporte
                ' o un formato impreso
                Script.RunForm "PROSPECTOS", Me, Ambiente,, False 
           Case "PreciosEspeciales"
                Script.RunForm "PRECIOSESPECIALES", Me, Ambiente,, False

           Case "Recostear"
                Script.RunForm "COSTEO", Me, Ambiente,, False

           Case "Touch"

                Set Ventas = CreateObject("MyBVentas.ventas")
                Set Ventas.Ambiente = Ambiente
                Ventas.ActivaTouch

           Case "PuntoVenta2"

                Set Ventas = CreateObject("MyBVentas.ventas")
                Set Ventas.Ambiente = Ambiente
                Ventas.ActivaPuntoDeVenta2

           Case "Recalcular"

                Script.RunProcess "RECALCULAINVENTARIO", Me, Ambiente

           Case "Servicios"

                ejecuta "CD010"                          

           Case "DatosGenerales" 

                Me.DatosEmpresa

           Case "Mantenimiento"

                RepararBaseDedatos

           Case "FormatosImpresion"

                Formatos

           Case "Consecutivos"

                'Consecutivos
				SetSessionValue Ambiente, "FORMPADRECONSEC", "BUSINESSMANAGER"
				Script.RunForm "CONSECUTIVOS", Me, Ambiente,, True

		   Case "GeneralConsecutivos"

				Script.RunForm "TODOSLOSCONSECUTIVOS", Me, Ambiente,, True

           Case "BorrarBase"

                EliminaDatos

           Case "AltaEmpresa"

                Script.RunForm "DATOSEMPRESA", Me, Ambiente,, False

           Case "CambiarEmpresa"

                CambiaEmpresa

           Case "RegistroDeLicencia"

                 AcercaDe                                   

           Case "ConfigGeneral"

                 ConfigGeneral                      

           Case "SincronizarProcedimientos"

                 SincronizarProcedimientos
 
          Case "Rangos"

                 establecerPeriodosDeCarpetas

         Case "configpocket"                     
		      Script.RunForm "CONFIGPOCKETPC", Me, Ambiente,, True
         Case "inventariopocket"
			  Script.RunForm "ENTRADAPOCKET", Me, Ambiente,, True
         Case "pocket"
              Script.RunForm "EXPORTAIMPORTAPOCKET", Me, Ambiente,, True
			  'Script.RunProcess "EXPORTAINFOTOPOCKET", Parent, Ambiente
		 Case "confirmaventas"
			  Script.RunForm "CONFIRMAVENTAS", Me, Ambiente,, True
		 Case "creanotas"
			  Script.RunForm "AsistenteNotasPocket", Me, Ambiente,, True
		 'Case "creainventariofisico"
			  'Script.RunForm "INVFISICOPOCKET", Me, Ambiente,, True
		 Case "reportepedidospk"
         	  EjecutaReporte "PEDIDOSMOBILE" 
		 Case "cargainicial"
         	  EjecutaReporte "PRODSPOCKET"
		 Case "ventaspendientes"
        	  EjecutaReporte "VENTASMOBILE"
		 Case "articulosvendidos"
   	           EjecutaReporte "VENTAPRODSMOBILE"

         Case "HojaInventario"
                       
              Script.RunForm "HOJACOSTOS", Me, Ambiente,, False
              'muestraHojaDeInventario   
                                           
         Case "CobroCaja"

              Set ventas = CreateObject( "MyBVentas.ventas" )
              Set ventas.Ambiente = Me.Ambiente
              Ventas.MuestraCobroEncaja                              

         Case "Consulta"

              Script.RunForm "CONSULTA", Me, Ambiente,, False           

         Case "etenvio"

              Script.RunForm "ETENVIO", Me, Ambiente,, False

         Case "FacturaCierre"

              Script.Runform "FACTURADECIERRE", Me, Ambiente,, False

         Case "series"

              activaSeriesSalida       

         Case "lotes"

              activaFormaDeLotes

         Case "CambioCosto"

              activaFormaDeCambiosDeCosto

         Case "ReportesFinancieros"

              ShellRun Me.hWnd, "Open", "http://localhost:4410/index.myweb?newsession"

         Case "seriesseguimiento"

              Script.RunForm "RASTREOSERIE", Me, Ambiente,, False

         Case "MBInventario"

              Set MBInventario = CreateObject( "MyBArticulos.Articulos" )
              Set MBInventario.Ambiente = Ambiente
              MBInventario.muestraMBInventario
                                                            
         Case "ordenproduccion"   

              Set p = CreateObject( "MyBProduccion.produccion" )
              Set p.Ambiente = Ambiente
              p.NuevaOp

         Case "capturaetiquetas"   

              Set p = CreateObject( "MyBProduccion.produccion" )
              Set p.Ambiente = Ambiente
              p.MuestraCapturaDeFracciones    

         Case "seguimiento"

              Set p = CreateObject( "MyBProduccion.produccion" )
              Set p.Ambiente = Ambiente
              p.muestraSeguimiento   

         Case "destajo"

              Script.RunForm "DESTAJO", Me, Ambiente,, False 

         Case "ordenservicio"

              Script.RunForm "ORDENES", Me, Ambiente,, False

         Case "BuscarTicket"

              Script.RunForm "BUSCATICKET", Me, Ambiente,, False
               
         Case "existenciaremota"

              Script.RunForm "EXISTENCIAREMOTA", Me, _
              Ambiente, , False                    

         Case "Abonos"

             Set Cobranza = CreateObject( "MyBCobranza.Cobranza" )
             Set Cobranza.Ambiente = Ambiente
             Cobranza.NuevoAbono  

                                          
         Case "ImportarMySQL"
 
             Script.RunForm "IMPORTAMYSQL", Me, Ambiente,, False
 
         Case "Analisis"

             Script.RunForm "PRODUCTOSSUGERIDO", Me, Ambiente,, False

         Case "UEPS"

             Script.RunForm "COSTEOUEPS", Me, Ambiente,, False
        
         Case "PEPS"

             Script.RunForm "COSTEOPEPS", Me, Ambiente,, False

         Case "Secciones"

             Script.RunForm "BUSQUEDASECCIONES", Me, Ambiente,, False
                        
         Case "Comandas"

             Script.RunProcess "AAMESAS", Me, Ambiente

         Case "Reservaciones"

             Script.RunProcess "AARESERVACIONES", Me, Ambiente         

         Case "Backup"

             Script.RunForm "RESPALDOBASEDEDATOS", Me, Ambiente,, False

         Case "Menu"

             Script.RunForm "CATEGORIASMENU", Me, Ambiente,, True

         Case "r_impresoras"

             Script.RunForm "IMPRESORAS", Me, Ambiente,, True

         Case "CierreTienda"

             Script.RunForm "CORTETIENDA", Me, Ambiente,, True
  
         Case "Huella"

             Script.RunHuellaForm "REGISTROACCESO", Me, Ambiente,, True

         Case "Clientes2"
 
             Script.RunForm "CLIENTES", Me, Ambiente,, False

         Case "Proveedores2"
                                                     
             Script.RunForm "PROVEEDORES", Me, Ambiente,, False

         Case "FPersonalizados"

             Script.RunForm "GENERADORDOCUMENTOS", Me, Ambiente,, False

         Case "materiales"

             Script.RunForm "HOJADEMATERIALES", Me, Ambiente,, False
                            
         Case "imc"

             Script.RunForm "IMC", Me, Ambiente,, True

         Case "RemisionFactura"

             Script.RunForm "REMISIONFACTURA", Me, Ambiente,, True

		 Case "ImpClientes"
                      
			 Script.RunForm "IMPORTACLIENTES", Me, Ambiente,, True

		 Case "ImpProveedores"
                                                         
			 Script.RunForm "IMPORTAPROVEEDORES", Me, Ambiente,, True

		 Case "CambioPrecio"

			 Script.RunForm "CAMBIOPRECIO", Me, Ambiente,, True                  
		 case "ConfigSucursal"

			 Script.RunForm "CONEXIONES", Me, Ambiente,, True

		 Case "FacturaElectronica"        
             ConfigFile = Ambiente.Path + "\econfig.txt"
             BatchFile = Ambiente.Path + "\FElectronica.bat"

             Set rstConfiguracion = CreaRecordSet("select * from FEConfig ", Ambiente.Connection)
             if not rstConfiguracion.EOF then                                         
                Set fso=CreateObject("Scripting.FileSystemObject")
                If fso.FileExists(ConfigFile) Then
                   fso.DeleteFile ConfigFile 
                End if                                                        
                If fso.FileExists(BatchFile) Then
                   fso.DeleteFile BatchFile 
                End if                                                        
                If not fso.FolderExists(rstConfiguracion("FileLocation")) Then
                   fso.CreateFolder(rstConfiguracion("FileLocation"))
                End if 

                outline BatchFile,"@echo off" + vbCrLf
                outline BatchFile,"cd " + Ambiente.Path + vbCrLf
                outline BatchFile,"start /wait FElectronica.exe" + vbCrLf
                outline BatchFile,"exit" + vbCrLf

                ambiente.connection.execute "exec FEInicializaSerie '"+ trim(ambiente.estacion) + "'"
                outline ConfigFile,ambiente.connection + vbCrLf                                 
                ShellRun Me.hWnd, "Open",Ambiente.Path + "\FElectronica.bat" 
             end if
		 


		 Case "Tallas"

			 Script.RunForm "TallaColModelos", Me, Ambiente,, True

         Case "CFD"

   		     Script.RunForm "DATOSCFD", Me, Ambiente,, True

         Case "CFDGEN"

   		     Script.RunForm "FACTURAELECTRONICA", Me, Ambiente,, True

         Case "CFDPAPEL1"

             Set report = CreateObject( "Reportes.Reportes" ) 
             Call report.LoadReport( Ambiente.Path & "\Formatos\FormatoFacturaCarta.mrt" )

             Call report.ReportQuery( "datosemisor", "SELECT * FROM cfd_datos", Ambiente.Connection )
             Call report.ReportQuery( "domicilioEmision", "SELECT * FROM cfd_domicilioexpedicion", Ambiente.Connection )
             Call report.ReportQuery( "encabezado", "SELECT * FROM ventas WHERE venta = 1", Ambiente.Connection )
             Call report.ReportQuery( "partidas", "SELECT * FROM partvta WHERE venta = 1 ORDER BY id_salida", Ambiente.Connection )
             Call report.ReportQuery( "cliente", "SELECT * FROM clients WHERE cliente = 'SYS'", Ambiente.Connection )
             Call report.ReportQuery( "moneda", "SELECT * FROM monedas WHERE moneda = 'MN'", Ambiente.Connection )
                                                                         

             Call report.DesignReport()
             'Report.ShowReport()
             Set Report = Nothing        
                                                                                                
         Case "CFDPAPEL2"

             Set report = CreateObject( "Reportes.Reportes" ) 
             Call report.LoadReport( Ambiente.Path & "\Formatos\FormatoFacturaTicket.mrt" )

             Call report.ReportQuery( "datosemisor", "SELECT * FROM cfd_datos", Ambiente.Connection )
             Call report.ReportQuery( "domicilioEmision", "SELECT * FROM cfd_domicilioexpedicion", Ambiente.Connection )
             Call report.ReportQuery( "encabezado", "SELECT * FROM ventas WHERE venta = 1", Ambiente.Connection )
             Call report.ReportQuery( "partidas", "SELECT * FROM partvta WHERE venta = 1 ORDER BY id_salida", Ambiente.Connection )
             Call report.ReportQuery( "cliente", "SELECT * FROM clients WHERE cliente = 'SYS'", Ambiente.Connection )
             Call report.ReportQuery( "moneda", "SELECT * FROM monedas WHERE moneda = 'MN'", Ambiente.Connection )

             report.DesignReport()
             'Report.ShowReport()
             Set Report = Nothing
		 ' cambiar claves de articulos
		 Case "CambiaClaveProd"
			Script.RunForm "CAMBIACLAVEIND", Me, Ambiente,, False
		' cambiar descuentos especiales
		Case "DescuentosEspeciales"
		Script.RunForm "descuentoEspecial", Me, Ambiente,, False	

    End Select    

End Sub                                    


Sub ejecuta( nombreProcedimiento )
    Dim rstFormato

    Set rstFormato = CreaRecordSet( _
    "SELECT * FROM formatos WHERE formato = '" & nombreProcedimiento & "'",_
    Ambiente.Connection )

    If rstFormato.EOF Then
       Exit Sub
    End If
    
    ' 1.- Tipo de programa           
    ' 2.- Codigo de programacion
    ' 3.- El objeto Ambiente
    ' 4.- El objeto padre
    Script.Preview rstFormato("tipo"), rstFormato("codigo"), Ambiente, Me

End Sub 
                  

Sub EjecutaReporte(reporte)

    Dim rstReporte

    Set rstReporte = Rst("SELECT * FROM formatosdelta WHERE formato = '" & reporte & "'", Ambiente.Connection )

    If rstReporte.EOF Then
       Exit Sub
    End If

	Script.Preview rstReporte("tipo"), rstReporte("codigo"), Ambiente, Me, rstReporte("formato"), ""


End Sub




