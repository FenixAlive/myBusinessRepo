@echo off
set dirBase=C:\Respaldos
set dbname=MyBusiness
set ext=bak
set end=5
FOR /L %%G IN (1,1,%end%) DO (
    set iter=%%G
    if NOT exist N"%dirBase%\%dbname%%iter%.%ext%" (
        if %%G GEQ %end%(
            set temp=1
        )else(
            set temp=%%G+1
        )
        if exist N"%dirBase%\%dbname%%temp%.%ext%" (
            del N"%dirBase%\%dbname%%temp%.%ext%"
        )
        BREAK
    )
  SqlCmd -E -S .\SQLEXPRESS -Q "BACKUP DATABASE [C:\MyBusinessDatabase\MyBusinessPOS2011.mdf] TO  DISK = N'%dirBase%\%dbname%%%G.%ext%' WITH NOFORMAT, NOINIT,  NAME = N'C:\MyBusinessDatabase\MyBusinessPOS2011.mdf-Completa Base de datos Copia de seguridad', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
) 

