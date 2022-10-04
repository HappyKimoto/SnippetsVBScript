Option Explicit

' Purpose: Print each table name
' Inputs:
' - Argument(0): File Path
' How to run:
' >cscript //nologo AccessPrintTableNames.vbs <AccessFilePath>

Dim strAccFp
Dim objApp, objDB, objTable

strAccFp = WScript.Arguments(0)

Set objApp = CreateObject("Access.Application")
WScript.Echo "TypeName(objApp)=" & TypeName(objApp)

objApp.OpenCurrentDatabase strAccFp

Set objDB = objApp.CurrentDB
WScript.Echo "TypeName(objDB)=" & TypeName(objDB)

For Each objTable In objDB.TableDefs
	WScript.Echo "TypeName(objTable)=" & TypeName(objTable) & " " & _
	"objTable.Name=" & objTable.Name
Next
objApp.Quit
Set objApp = Nothing