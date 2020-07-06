Sub Main()
	'dar alta base de datos
	execSQL "CREATE TABLE dbo.descuentoEspecial (  articulo nvarchar(30) NOT NULL , CONSTRAINT AK_TransactionID UNIQUE(articulo), pocodesc bit NULL, descuento float NULL );"
	Dim porcPoco
	Dim prodSinDes
	Dim prodPocoDes
	' 				    ALCOHOL 108, GEL GUKOL 250ML,  GEL GUKOL 250ML, GEL AG 1LT,      GEL AG 250ML,      GEL AG 500ML,   GEL AG 60ML,     GEL BRITZ,      GEL MUNCHKIN 500ML,  GEL MUNCHKIN 900ML, GEL JALOMA 60ML, GEL JALOMA 120ML, CUBRE DOBLE, CUBRE TRIPLE, CUBRE MASC,      GUANTE 121,  GUANTE ESTERIL,  TOALLAS DESINF,  SEDALMERK,  SINUBERASE,      HIDROXICLOROQUINA , OMEGA RX 30 
	prodSinDes = Array( "108",       "153",            "154",           "7502281191710", "7502281190119", "7502281190126", "7502281190102", "7503009088664", "7508006012100",     "7508006012117",    "759684154232",  "759684154140",   "118",       "157",        "7506329100214", "121",       "7502224240659", "7503006503474", "126",      "7501159580182", "7502009747458",   "7502227420362")
	'                   VENDA 30 B-CARE
	For i = LBound(prodSinDes) To Ubound(prodSinDes)
   		execSQL "INSERT into descuentoEspecial values( cast("& prodSinDes(i) &" AS nvarchar(30)),0,0)"
 		'execSQL "DELETE from descuentoEspecial where articulo = cast("& prodSinDes(i) &" AS nvarchar(30))"   
	Next
	prodPocoDes = Array("7503003707301")
 	porcPoco =    Array(23)
	For i = LBound(prodPocoDes) To Ubound(prodPocoDes)
   		execSQL "INSERT into descuentoEspecial values( cast("& prodPocoDes(i) &" AS nvarchar(30)),1,"& porcPoco(i) &")"
    Next
	msgbox "Termin�"
End Sub  

Sub execSQL( strSQL )
 
    On Error Resume Next                              
    Ambiente.Connection.Execute (strSQL)
 
    If Err.Number <> 0 Then
       MyMessage (strSQL) & " --- " & Err.Description & vbCrLf
    End If

End Sub
