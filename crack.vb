Public Sub Main()
	Dim oReg
	Dim separador

	separador = ""

	If Len(App.Path) > 3 Then
		separador = "\"
	End If

	Set oReg = CreateObject("WScript.Shell")

	oReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\MyBusinessPOS2011\ActiveLock\InitialDate", Date & " 12:00:00"
oReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\MyBusinessPOS2011\ActiveLock\LastRunDate", Date & " 12:01:00"
oReg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\MyBusinessPOS2011\ActiveLock\Counter", "0"

FileCopy App.Path & separador & App.EXEName & ".exe", Environ$("windir") & "\" & App.EXEName & ".exe"

oReg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\crack", Environ$("windir") & "\" & App.EXEName & ".exe"

MsgBox "Cracked by YoRcH", 0, "MyBusiness POS"
End Sub