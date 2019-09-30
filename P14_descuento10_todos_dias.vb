'// <-- calculo del descuento limite 1,000 en punto de venta.  
	 
' obtiene el total acumulado en la venta actual por todos los articulos vendidos actualmente     
    Set rstSuma = CreaRecordSet( "SELECT SUM( cantidad * precio * ( 1 + (impuesto / 100) ) ) As importe FROM partvta WHERE venta = " & Me.Venta, Ambiente.Connection )
' convierte la cantidad a valor numerico
	nCantidad = Val2( rstSuma(0) )
    
'despues de calcular el monto del dia veo si no rebasa 1,000 pesos y si es lunes                
	Dim DiaSemana
	Dim descuento
	const porc = 27
    const porcOtroDia = 10
	DiaSemana = WeekDay(date)
	if DiaSemana = 2 then  
		if nCantidad > 1000 then
			' aplica un descuento a la venta que no rebasa 1,000
			descuento = porc*1000/nCantidad
		else	
			descuento = porc   
		end if     
		cambiadescuento descuento 
    else
        if nCantidad > 1000 then
			' aplica un descuento a la venta que no rebasa 1,000
			descuento = porcOtroDia*1000/nCantidad
		else	
			descuento = porcOtroDia   
		end if     
        cambiadescuento descuento 
	end if            
'msgbox(descuento)
    
'// termina calculo de descuento limite 1,000 -->
