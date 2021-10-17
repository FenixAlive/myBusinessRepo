const ServerName = "localhost\SQLEXPRESS"
const BackupDir = "C:\Respaldos"
const DBName = "[C:\MyBusinessDatabase\MyBusinessPOS2011.mdf]"
const numRes = 3

Dim backupFileName

SET conn = CREATEOBJECT("ADODB.Connection")
SET cmd = CREATEOBJECT("ADODB.Command")
SET rs = CREATEOBJECT("ADODB.RecordSet")

DIM fso    
Set fso = CreateObject("Scripting.FileSystemObject")

if not (fso.FolderExists(BackupDir)) Then
    fso.CreateFolder(BackupDir)
End If

Dim i
For i=1 to numRes
    If NOT(fso.FileExists("C:\Respaldos\" & i & ".bak")) Then
        backupFileName = i & ".bak"
        Exit For
    End If
    
Next
if (i = numRes) Then
    fso.DeleteFile "C:\Respaldos\1.bak"
else
    If (fso.FileExists("C:\Respaldos\"& i+1 & ".bak")) Then
        fso.DeleteFile "C:\Respaldos\"& i+1 & ".bak"
    End If
End If
conn.open "Provider=SQLOLEDB.1;Data Source=" & ServerName & "; Integrated Security=SSPI;InitialCatalog=" & DBName

call backupDB(backupFileName)

conn.close

SUB backupDB(name)
    backupFilePath = BackupDir & "\" & name
    SET cmdbackup = CREATEOBJECT("ADODB.Command")
    cmdbackup.activeconnection = conn
    cmdbackup.commandtext = "backup database " & DBName & " to disk='" & backupFilePath & "' WITH NOFORMAT, NOINIT,  NAME = N'C:\MyBusinessDatabase\MyBusinessPOS2011.mdf-Completa Base de datos Copia de seguridad', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
    cmdbackup.EXECUTE
END SUB
