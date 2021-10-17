ECHO OFF
set dirBase=C:\Respaldos
set dbname=MyBusiness
set ext=bak
set end=5
setlocal enabledelayedexpansion
FOR /L %%G in (1,1,%end%) DO (
    SET /a "I=%%G"
    SET /a "J=%%G+1"
    ECHO %dirBase%!B!.%ext%
    ECHO %dirBase%\%dbname%!I!.%ext%
    if NOT exist N"%dirBase%\%dbname%!I!.%ext%" (
        if %%G GEQ %end%(
            echo "ya existe el 5"
        )else(
            echo "no existe el numero "!I!
        )
        BREAK
    )
)