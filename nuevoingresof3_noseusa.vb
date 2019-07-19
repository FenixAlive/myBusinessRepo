' al presionar f3 imprimir ticket de nuevo ingreso
	If Ambiente.Tag = 114 Then
       If Question( "¿Quieres imprimir tu ticket de ingreso al sistema?" ) Then
          '// comienza impresión de ticket por entrada a punto de venta 
             
cLineaNueva = Chr(13) & Chr(10)
cSalida = ""                                              

'Encabezado
cSalida = cSalida & cLineaNueva & cLineaNueva
Set rstTextTicket = CreaRecordSet( "SELECT * FROM tickettext", Ambiente.Connection ) '"CreaRecordSet()" consulta si el usuario colocó algun mensaje en el encabezado y pie del ticket a traves de la utileria "editar texto del ticket" y lo almacena en "rstTextTicket"

If rstTextTicket.EOF Then 'en caso de que el usuario haya dejado vacio el encabezado y pie de ticket (rstTextTicket Nulo), entrar al bloque para colocar los datos de la Empresa que se colocan con la Configuración "Datos Generales de la Empresa"
cSalida = cSalida & "" & Ambiente.Empresa & cLineaNueva '"Ambiente.Empresa" es el nombre de la Empresa que se coloca con la Configuración "Datos Generales de la Empresa". "cLineaNueva" es un salto de linea en el ticket. en caso de que necesitara un salto mas solo basta agregar " & cLineaNueva" al final de esta liea
cSalida = cSalida & "" & Trim( Ambiente.Direccion1 ) & cLineaNueva ''"Trim( Ambiente.Direccion1 )" es la dirección que se encuentra en la primera linea, que se coloca con la Configuración "Datos Generales de la Empresa". la función "Trim()" elimina los espacios en blanco antes y al final de la dirección1
cSalida = cSalida & "" & Trim( Ambiente.Direccion2 ) & cLineaNueva ''"Trim( Ambiente.Direccion2 )" es la dirección que se encuentra en la segunda linea, que se coloca con la Configuración "Datos Generales de la Empresa". la función "Trim()" elimina los espacios en blanco antes y al final de la dirección2
cSalida = cSalida & "" & Trim( Ambiente.Telefonos ) & cLineaNueva ''"Trim( Ambiente.telefonos )" es el espacio para mostrar algun dato adicional que se coloca con la Configuración "Datos Generales de la Empresa". la función "Trim()" elimina los espacios en blanco antes y al final de la dirección1
Else

cSalida = cSalida & rstTextTicket("textheader") 'en caso de que el usuario haya colocado algun dato en el encabezado y pie de ticket, tomamos el Encabezado de este "rstTextTicket("textheader")" y lo agregamos a la variable "cSalida" con "cSalida = cSalida & "
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


msgbox cSalida                                            

'PrintText cSalida

'//termina impresión de ticket por entrada a punto de venta
       End If
    End If  