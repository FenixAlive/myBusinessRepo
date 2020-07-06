Sub Main()
	'dar alta base de datos
	execSQL "CREATE TABLE dbo.descuentoEspecial (  articulo nvarchar(30) NOT NULL , CONSTRAINT AK_TransactionID UNIQUE(articulo), pocodesc bit NULL, descuento float NULL );"
	Dim porcPoco
	Dim prodSinDes
	Dim prodPocoDes
	' 				   VALSARTAN AMLODIPINO, ALCOHOL 108,  GEL 155, GEL 156, GEL BRITZ,       GEL JALOMA,     CUBRE 118, CUBRE 159, CUBRE MASC,      GUANTE 121,  GUANTE ESTERIL,  TOALLAS DESINF,  SEDALMERK 126 
	prodSinDes = Array("7501124819521",      "108",        "155",   "156",   "7503009088664", "759684154140", "118",     "159",     "7506329100214", "121",       "7502224240659", "7503006503474", "126" )
	'                   PANTOPRAZOL 40,   LOSARTAN 50,    ATORVASTATINA 40 10, ATORVASTATINA 40 10, PAROXETINA 10,   PAROXETINA 10,   PAROXETINA 10,   TIRAS REACTIVAS 50+10, TIRAS REACTIVAS 50+10, TIRAS REACTIVAS 50, BAUMANOMETRO,   TRIBEDOCE COMP,  METOPROLOL 100,  METOPROLOL 100,  METOPROLOL 100   
	prodPocoDes = Array("7501349028364", "7502240450070", "785120754858",      "7501349028913",     "7501349024939", "7501075718041", "7502227872628", "701822717250",        "7501554500259",       "353885771504",     "073796712020", "7501537164713", "7501075714173", "7501075718881", "7501493888944")
 	porcPoco =    Array(2.67,            8.75,            11.36,               11.36,               8.75,            8.75,            8.75,            25,                    25,                    25,                 23.73,           18.89,           23,              23,              23)  
	For i = LBound(prodSinDes) To Ubound(prodSinDes)
   		execSQL "INSERT into descuentoEspecial values( cast("& prodSinDes(i) &" AS nvarchar(30)),0,0)"
 		'execSQL "DELETE from descuentoEspecial where articulo = cast("& prodSinDes(i) &" AS nvarchar(30))"   
	Next
	For i = LBound(prodPocoDes) To Ubound(prodPocoDes)
   		execSQL "INSERT into descuentoEspecial values( cast("& prodPocoDes(i) &" AS nvarchar(30)),1,"& porcPoco(i) &")"
    Next
	msgbox "Termino"
End Sub  

Sub execSQL( strSQL )
 
    On Error Resume Next                              
    Ambiente.Connection.Execute (strSQL)
 
    If Err.Number <> 0 Then
       MyMessage (strSQL) & " --- " & Err.Description & vbCrLf
    End If

End Sub
