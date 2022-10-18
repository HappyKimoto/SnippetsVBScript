Option Explicit

Const DEBUGGING = False
Const FP7Z = "C:\Program Files\7-Zip\7z.exe"

' Objective:
' 	- Automate folder compression.
' 
' Preparation:
' 	- Place a shortcut of this VBScript (or this file itself) on the SendTo Folder.
'	- 7zip is installed in the system.
'     - https://www.7-zip.org/download.html
'     - To run this VBScript, necessary files are are 7z.dll and 7z.exe at "C:\Program Files\7-Zip\".
' How to use:
' 	- Right click a folder on explorer.
'	- Select SendTo.
' 	- Select this VBScript.
'	- Choose compression option by bit flag.
'     - Add date/time to make each compression unique.
'     - Add encryption for security.
' 
' How to enforce certain encoding:
' 	- File names may be prone to regional settings.
' 	- If you want to enforce certain code page / character encoding, run this VBScript from batch.
' 	- Before cscript execution, change the code page to the desired code page.
' 	- If shift-jis is desired, change it to code page 932.
'   - Place this batch script to the SendTo folder.
' ------------------------------------------------------------------------------------------
' @echo off
' echo "%1"
' chcp 932
' cscript "FilePathTo\7zAdd.vbs" %1
' pause
' -----------------------------------------------------------------------------------------

' Purpose: Print current directory
Sub PrintCurrentDirectory()
	Wscript.Echo "objShell.CurrentDirectory=" & objShell.CurrentDirectory
End Sub

' Purpose: Pad 2 zeros
Function strPadZero(ByVal num)
	Dim strNum
	strNum = CStr(num)
	If len(strNum) = 1 Then
		strPadZero = "0" & strNum
	Elseif len(strNum) = 2 Then
		strPadZero = strNum
	End If
End Function

'# Current DateTime
Function strCurrentDateTime()
	Dim cur
	cur = Now()
	strCurrentDateTime = _
	Year(cur) & strPadZero(Month(cur)) & strPadZero(Day(cur)) & _
	"-" & strPadZero(Hour(cur)) & strPadZero(Minute(cur))
End Function

'# Get the passed folder path.
Function ArgFolder()

	Dim strPath
	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'# Check argument count is 1
	If WScript.Arguments.Count = 1 Then
		'# Check if argument is a valid folder.
		strPath = WScript.Arguments(0)
		If objFSO.FolderExists(strPath) Then
			'# Return folder path
			ArgFolder = strPath
		Else
			'# Error
			MsgBox "strPath=" & strPath, 48, "Not Folder Error"
			WScript.Quit
		End If
	Else
		' Error
		MsgBox "WScript.Arguments.Count=" & WScript.Arguments.Count, 48, "Argument Count Error"
		WScript.Quit
	End If
	
	' Clean object
	Set objFSO = Nothing	
End Function

' Masure sure 7zip exists in the system.
Sub Check7zApp()
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'# Confirm that 7z.exe exists in the system.
	If Not objFSO.FileExists(FP7Z) Then
		MsgBox "7z.exe is not found in Program Files", 48, "File Not Found Error"
		WScript.Quit
	End If	

	' Clean object
	Set objFSO = Nothing
End Sub

' The main sub procedure.
Sub Main()
	'# Declare Variables
	Dim strPath
	Dim strFolderName
	Dim strParent
	Dim strChoiceNum
	Dim intChoiceNum
	Dim objFSO
	Dim objShell
	Dim strStatement1
	Dim strStatement2
	Dim strStatement3
	Dim strStatementResult

	'# Set Objects
	Set objShell = CreateObject("Wscript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	'# Confirm that passed argument is a folder path.
	strPath = ArgFolder()

	'# Confirm that 7z.exe exists in the system.
	Call Check7zApp

	'# Get folder's base name
	strFolderName = objFSO.GetFolder(strPath).Name
	If DEBUGGING Then MsgBox "strFolderName=" & strFolderName
	
	'# Get parent folder path
	strParent = objFSO.GetParentFolderName(strPath)
	If DEBUGGING Then MsgBox "strParent=" & strParent

	'# Let the user to choose how to archive the folder
	'# Because VBScript does not support check box, user input is in the form of bit flag.
	strChoiceNum = InputBox(Join(Array("Bit field=1: Add DateTime to folder name", "Bit field=2: Encrypt", _
	"Example:", "0: Bare compression", "1: With date/time", "2: With encryptipn", "3: With date/time and encryption (default)", _
	"Press [Cancel] if not compressing"), vbCrlf), _
	"Select Compression Options", _
	"3")
	If DEBUGGING Then MsgBox "strChoiceNum=" & strChoiceNum

	'# By default, type is 7z and recursive.
	strStatement3 = " -t7z -r"

	If strChoiceNum = "" Then
		MsgBox "Cancel was selected"
	'# If user input is a valid integer, then proceed
	Elseif IsNumeric(strChoiceNum) Then
		'# Convert string to integer
		'# Integer in VBScript is 16-bit (0xFFFF)
		intChoiceNum = CInt(strChoiceNum)
		If DEBUGGING Then MsgBox "intChoiceNum=" & intChoiceNum
		
		'# 0x0001: Add Date/Time to the folder name
		If (intChoiceNum And &h0001) > 0 Then
			If DEBUGGING Then MsgBox "&h0001=ON"
			strFolderName = strFolderName & "(" & strCurrentDateTime & ")"
		Else
			If DEBUGGING Then MsgBox "&h0001=OFF"
		End If

		'# 0x0002: Add password
		If (intChoiceNum And &h0002) > 0 Then
			If DEBUGGING Then MsgBox "&h0002=ON"
			strStatement3 = strStatement3 & " -p""" & strFolderName & """ -mhe"
		Else
			If DEBUGGING Then MsgBox "&h0002=OFF"
		End If
		
		'# Statement: Change directory to the selected folder
		strStatement1 = "cd """ & strPath & """"
		If DEBUGGING Then MsgBox "strStatement1=" & strStatement1
		
		'# Statement: Add all files
		strStatement2 = """" & FP7Z & """ a """ & strParent & "\" & strFolderName & ".7z"" *.*"
		If DEBUGGING Then MsgBox "strStatement2=" & strStatement2
		
		If DEBUGGING Then MsgBox "strStatement3=" & strStatement3
		
		'# Constructed Statement
		strStatementResult = "%comspec% /k " & strStatement1 & _
		" & " & strStatement2 & strStatement3
		If DEBUGGING Then MsgBox "strStatementResult=" & strStatementResult

		'# Execute statement
		objShell.Run strStatementResult, 3, False
		
	'# If user input is not a valid integer, then quit.
	Else
		MsgBox "Choice is not a number", 48, "ERROR"
		WScript.Quit
	End If


	'# Clean objects
	Set objShell = Nothing
	Set objFSO = Nothing

End Sub

Call Main
