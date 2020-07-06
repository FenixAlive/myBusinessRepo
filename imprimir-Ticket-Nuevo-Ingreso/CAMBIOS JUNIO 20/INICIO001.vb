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
	execSQL "INSERT into asistencia values((SELECT TOP 1 id FROM asistencia ORDER BY id DESC)+1,CAST(CURRENT_TIMESTAMP AS nvarchar(19)), '"&Trim( Ambiente.Uid )&"', CURRENT_TIMESTAMP, null, null)"
 	Ambiente.var9 = False
End If                                  

'//-----------------------termina impresión de ticket por entrar a administracion

End Sub