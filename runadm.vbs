
' #=======================================================#
' #  RUNADM.VBS                                           #
' #=======================================================#
' #  Run a specified program with admin rights.           #
' #                                                       #
' #     Copyright(C) ZulNs, Yogyakarta, June 29'th, 2013  #
' #=======================================================#

Option Explicit

Const TITLE = "ZulNs: Run Admin", ADMIN = "adm"

If Wscript.Arguments.Count = 0 Then
	MsgBox "Run a program with administrative priveleges.", _
			vbInformation, TITLE
	Wscript.Quit
End If

Dim params, quote, arg, shell
quote = Chr(34)

For Each arg In Wscript.Arguments
	If arg <> ADMIN Then
		If InStr(Trim(arg), " ") = 0 Then
			params = params & Trim(arg) & " "
		Else
			params = params & quote & Trim(arg) & quote & " "
		End If
	End If
Next
Set arg = Nothing
params = Trim(params)

If Wscript.Arguments.Item(0) <> ADMIN Then
	Set shell = CreateObject("Shell.Application")
	shell.ShellExecute "wscript.exe", _
			quote & WScript.ScriptFullName & quote & _
			" " & ADMIN & " " & params, "", "runas", 1
Else
	Dim retCode
	Set shell = CreateObject("Wscript.Shell")
	retCode = shell.Run(params, 1, True)
	
	If retCode = 0 Then
		MsgBox "DONE.", vbInformation, TITLE
	Else
		MsgBox "FAILED!!!", vbCritical, TITLE
	End If
End If

Set shell = Nothing
