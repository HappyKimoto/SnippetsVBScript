Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' https://ss64.com/vb/browseforfolder.html
'
' .BrowseForFolder
' Prompt the user to select a folder.
'
' Syntax 
'       .BrowseForFolder(WINDOW_HANDLE, "Window Title", WINDOW_OPTIONS, StartPath)
'
' Key
'    WINDOW_HANDLE  This should always be 0
'
'    WINDOW_OPTIONS
'      Const BIF_RETURNONLYFSDIRS   = &H0001  (The default)
'      Const BIF_DONTGOBELOWDOMAIN  = &H0002
'      Const BIF_STATUSTEXT         = &H0004
'      Const BIF_RETURNFSANCESTORS  = &H0008
'      Const BIF_EDITBOX            = &H0010
'      Const BIF_VALIDATE           = &H0020
'      Const BIF_NONEWFOLDER        = &H0200
'      Const BIF_BROWSEFORCOMPUTER  = &H1000
'      Const BIF_BROWSEFORPRINTER   = &H2000
'      Const BIF_BROWSEINCLUDEFILES = &H4000
'      ' These can be combined e.g. BIF_EDITBOX + BIF_NONEWFOLDER
'
'    StartPath     A drive/folder path or one of the following numeric constants: 
'      DESKTOP = 0
'      PROGRAMS = 2
'      DRIVES = 17
'      NETWORK = 18
'      NETHOOD = 19
'      PROGRAMFILES = 38
'      PROGRAMFILESx86 = 48
'      WINDOWS = 36
' Although you can display files with .BrowseForFolder, the method will only return a folder, hence the name.
'
' Examples
'
' Dim objFolder, objShell
' Set objShell = CreateObject("Shell.Application")
' Set objFolder = objShell.BrowseForFolder(0, "Please select the folder.", 1, "")
' If Not (objFolder Is Nothing) Then
'    wscript.echo "Folder: " & objFolder.title
'    wscript.echo "Full Path: " & objFolder.Self.path 
' End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Purpose: Prompt folder input with Folder Browser
' Inputs: The user to select a folder
' Return: File path
Function strPromptFolder()
	Dim strFolder, objShell, objFolder
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0, "Select a folder", 1, "")
	If objFolder Is Nothing Then
		wscript.echo "Error - a folder did not get selected."
		wscript.quit
	Else
		strPromptFolder = objFolder.Self.path
	End If
	Set objShell = Nothing
	Set objFolder = Nothing
End Function

Sub Test()
    Dim strFolder
    strFolder = strPromptFolder
    WScript.Echo strFolder, TypeName(strFolder)
End Sub
Test

