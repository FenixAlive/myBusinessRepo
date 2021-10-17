@echo off
set dirBd=C:\MyBusinessDatabase
set dirBase=C:%HOMEPATH%\drive
set dir1=%dirBase%\MyBusinessDatabase1
set dir2=%dirBase%\MyBusinessDatabase2
set dir3=%dirBase%\MyBusinessDatabase3
set dirUsb=E:\
set dirzip=C:\Program Files\7-Zip\7z.exe
set dirzip=D:\Archivos de Programas\7-Zip\7z.exe
if exist "%dir1%" (
    if exist "%dir2%" (
      if exist "%dir3%" (
        rmdir /s /q "%dir3%"
        del "%dir3%.zip"
        del "%dirUsb%MyBusinessDatabase3.zip"
      )
        rmdir /s /q "%dir1%"
        del "%dir1%.zip"
        del "%dirUsb%MyBusinessDatabase1.zip"
        xcopy "%dirBd%" "%dir3%" /E/H/C/i
        "%dirzip%" a -tzip "%dir3%.zip" "%dir3%"
        xcopy "%dir3%.zip" "%dirUsb%"

    ) else (
        if exist "%dir3%" (
            rmdir /s /q "%dir3%"
            del "%dir3%.zip"
            del "%dirUsb%MyBusinessDatabase3.zip"
        )
        xcopy "%dirBd%" "%dir2%" /E/H/C/I
        "%dirzip%" a -tzip "%dir2%.zip" "%dir2%"
        xcopy "%dir2%.zip" "%dirUsb%"
    )
   
) else (
    if exist "%dir3%" (
        rmdir /s /q "%dir2%"
        del "%dir2%.zip"
        del "%dirUsb%MyBusinessDatabase2.zip"
    )
    xcopy "%dirBd%" "%dir1%" /E/H/C/I
    "%dirzip%" a -tzip "%dir1%.zip" "%dir1%"
    xcopy "%dir1%.zip" "%dirUsb%" 
)