Sub Main()
	'dar alta base de datos
	execSQL "CREATE TABLE dbo.descuentoEspecial (  articulo nvarchar(30) NOT NULL , CONSTRAINT AK_TransactionID UNIQUE(articulo), pocodesc bit NULL, descuento float NULL );"
	Dim porcPoco
	Dim prodSinDes
	Dim prodPocoDes
	'Productos a incluir por primera vez:
	' 				   GELMICIN,       B.PINAVERIO,    CAPTOPRIL,       CAPTOPRIL,       DICLOFENACO,     DICLOFENACO,     DICLOFENACO,     DICLOFENACO,      ENALAPRIL,      ENALAPRIL,       ENALAPRIL,       GABAPENTINA 15, KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      KETOROLACO,      LORATADINA 10,   LORATADINA 20,   LORATADINA 20,   LOSARTAN 50MG,   LOSARTAN 100MG   METFORMINA,      METFORMINA,      METOPROLOL,      METOPROLOL,      OMEPRAZOL 14,    OMEPRAZOL 14,    OMEPRAZOL 14,    OMEPRAZOL 30,    OMEPRAZOL 30,    PANTOPRAZOL 20 7, PANTOPRAZOL 20 14, PANTOPRAZOL 40 14, SILDENAFIL 50 1, SILDENAFIL 50 1, SILDENAFIL 50 1, SILDENAFIL 100 1, SILDENAFIL 100 1, PIROXICAM,       ATORVASTATINA ,  ALCOHOL 108, CUBRE 118, CUBRE 158, GEL 156, GEL MUN,         GEL BRITZ,       GEL JALOMA,      GUANTE 121, GUANTE ESTERIL,  TOALLAS DESINF,  GEL 157, MASCARA,        SEDALMERK 126, TERRAMICINA 155
	prodSinDes = Array("780083148966", "821998000434", "7503000422405", "7501537102142", "7503000422368", "7502216802964", "7502009746017", "7501482201952", "7502216792845", "7502216791299", "7502227872246", "785120754919", "7501493888760", "7501557140308", "7501349024847", "7501573900559", "7502216791312", "7502009740244", "7501075716924", "7502216790339", "7502216793415", "7502009742828", "7502240450070", "7502240450711", "7501075717020", "7503004908776", "7501075714173", "7501075718881", "7501573902584", "7501571201221", "7501493888746", "7501342803562", "7502216792760", "7501349022485",  "7501277094318",   "7501349028364",   "7502216796812", "7501258210133", "7502009744457", "7502009744914", "7502216796836",   "7501537102371", "7501349024526", "108",       "118",     "158",     "156",   "7508006012117", "7503009088664", "759684154140", "121",       "7502224240659", "7503006503474", "157",   "7506329100214", "126",        "155")
	'                   CARBAMAZEPINA,   BUTILHIOSINA,   VENGESIC,       ITRACONAZOL,     ATORVASTATINA 40 10, ATORVASTATINA 40 10, PAROXETINA 10,   PAROXETINA 10,   DULOXETINA,      ADIOLOL,        ENFO-KI,         RMFLEX,          INFACAR ET 1,    INFACAR ET 2,    ACENOCUMAROL,    TRIBEDOCE COMP  
	prodPocoDes = Array("7501075710724", "821998000601", "780083140731", "7501075717174", "785120754858",      "7501349028913",     "7501349024939", "7501075718041", "7502009745478", "725742762145", "7501590285608", "7508304309513", "7502253072191", "7502253072207", "7501471889611", "7501537164713")
 	porcPoco =    Array(17.5,            10,             4.55,           6.25,            6.25,                6.25,                6.25,            6.25,            5,               2.5,            10,              11.22,           6.25,            6.25,            19.7,            16.67)   
	
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

