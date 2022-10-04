Option Explicit

' Purpose: Print each field value
' Inputs:
' - Argument(0): File Path
' - Argument(1): Table Name
' - Argument(2): Field Name
' How to run:
' >cscript //nologo AccessPrintField.vbs <AccessFilePath> <TableName> <Field Name>

Dim strAccFp, strTable, strField
Dim objApp, objWS, objDB, objRS

strAccFp = WScript.Arguments(0)
strTable = WScript.Arguments(1)
strField = WScript.Arguments(2)

Set objApp = CreateObject("Access.Application")
Set objWS = objApp.DBEngine(0)
Set objDB = objWS.OpenDatabase(strAccFp)
Set objRS = objDB.OpenRecordset(strTable)

Do Until objRS.EOF
	WScript.Echo objRS(strField)
	objRS.MoveNext
Loop

objRS.Close
Set objRS = Nothing
Set objDB = Nothing
Set objWS = Nothing
objApp.Quit
Set objApp = Nothing
