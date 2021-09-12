@echo off
set dirBd=C:\MyBusinessDatabase
set dirBase=C:%HOMEPATH%\drive
set dir1=%dirBase%\MyBusinessDatabase1
set dir2=%dirBase%\MyBusinessDatabase2
set dir3=%dirBase%\MyBusinessDatabase3
if exist %dir1% (
    if exist %dir2% (
      if exist %dir3% (
        rmdir /s /q %dir3%
      )
        rmdir /s /q %dir1%
        xcopy %dirBd% %dir3% /E/H/C/I
    ) else (
        if exist %dir3% (
            rmdir /s /q %dir3%
        )
        xcopy %dirBd% %dir2% /E/H/C/I
    )
   
) else (
    if exist %dir3% (
        rmdir /s /q %dir2%
    )
    xcopy %dirBd% %dir1% /E/H/C/I
)