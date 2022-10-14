Option Explicit

' Read bytes from the source file and dump the hex string with the specified format.

' VarType
Const sconVerTypeNull = 1
Const sconVarTypeBytes = 8209

' ADODB Stream Type
Const adTypeText   = 2
Const adTypeBinary = 1

' File System Object
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

' Purpose: Convert Long to hex string with zero padding
' Inputs: lngNum as Long Number
' Return: Hex string with zero padding
Function strLongToHex(ByVal lngNum)
	strLongToHex = Right("00000000" & Hex(lngNum), 8)
End Function

' Purpose: Write ASCII text to file
' Inputs: 
' 	ByVal strFp as absolute file path
' 	ByVal strTxt as ASCII string to be saved on the file
' Effects: strTxt will be written to strFp.
Sub WriteTextASCII(ByVal strFp, ByVal strTxt)
	WScript.Echo "WriteTextASCII strFp=" & strFp
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTS: Set objTS = objFSO.OpenTextFile(strFp, ForWriting, True)
	Call objTS.Write(strTxt)
	objTS.Close
	Set objTS = Nothing
	Set objFSO = Nothing
End Sub

' Purpose: Delete files in the specified folder
' Inputs: 
' 	ByVal argDir as folder path
' Effects: All the files in the folder will be deleted.
Sub DeleteFiles(ByVal argDir)
	Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFile: For Each objFile In objFSO.GetFolder(argDir).Files
		WScript.Echo "Deleting: " & objFile.Name
		objFile.Delete
	Next
	Set objFile = Nothing
	Set objFSO = Nothing
End Sub

' Purpose: Converts Bytes to Hex String

Class BytesToHexString
	' output folder path to be initialized
	Private strDirOut
	Public Property Let DirOut(ByVal argDirOut)
		strDirOut = argDirOut
	End Property

	' column count to be initialized
	Private intColumnCount
	Public Property Let ColumnCount(ByVal argColumnCount)
		intColumnCount = argColumnCount
	End Property

	' start position of the bytes
	' used for row header and file name
	Private lngStartPosition
	Private Sub Class_Initialize()
		lngStartPosition = 0
	End Sub
	
	' Convert byte array to hex string and save on file
	Public Sub BytesToFile(ByVal argBytes)
		Dim i, j, k, varRow, strResult, intColumnUBound
		
		' Create column header
		ReDim varRow(intColumnCount - 1)
		For i = 0 To intColumnCount - 1
			varRow(i) = Right("00" & Hex(i), 2)
		Next
		strResult = "-------- " & Join(varRow, " ") & vbCrlf
		
		' Loop through rows
		intColumnUBound = intColumnCount - 1
		For i = LBound(argBytes) To UBound(argBytes) Step intColumnCount
			' Adjust the upper bound in case byte count is smaller than coulumn length
			If i + (intColumnCount - 1) > UBound(argBytes) Then
				intColumnUBound = UBound(argBytes) - i
			End If

			' Reset array buffer
			ReDim varRow(intColumnUBound)

			' Fill the array
			For j = 0 To intColumnUBound
				varRow(j) = Right("00" & Hex(AscB(MidB(argBytes, i+j+1, 1))), 2)
			Next
			' Update resultant string
			strResult = strResult & strLongToHex(lngStartPosition + i) & " " & Join(varRow, " ") & vbCrlf			
		Next
		
		' Write the result to file
		Call WriteTextASCII(strDirOut & "\" & strLongToHex(lngStartPosition) & ".txt", strResult)
		
		' Increment start position for the next round
		lngStartPosition = lngStartPosition + UBound(argBytes) + 1
	End Sub

End Class


Sub Main()
	' get parameters from the caller
	Dim strFpIn: strFpIn = WScript.Arguments(0)
	Dim strDirOut: strDirOut = WScript.Arguments(1)
	Dim intRowCount: intRowCount = CInt(WScript.Arguments(2))
	Dim intColumnCount: intColumnCount = CInt(WScript.Arguments(3))
	
	' clean the output folder
	DeleteFiles(strDirOut)
	
	' BytesToHexString class object
	Dim lngBlockSize: lngBlockSize = intRowCount * intColumnCount
	Dim objBytesToHexString: Set objBytesToHexString = New BytesToHexString
	objBytesToHexString.DirOut = strDirOut
	objBytesToHexString.ColumnCount = intColumnCount

	Rem - Stream
	Dim varChunk
	Dim objStream: Set objStream = CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = adTypeBinary
		.LoadFromFile(strFpIn)
		Do While True
			Rem - Read Bytes
			varChunk = .Read(lngBlockSize)
			WScript.Echo "TypeName(varChunk)=" & TypeName(varChunk), "VarType(varChunk)=" & VarType(varChunk)
			Rem - Process Bytes
			Select Case VarType(varChunk)
				Case sconVarTypeBytes
					WScript.Echo "UBound(varChunk)=" & UBound(varChunk)
					objBytesToHexString.BytesToFile(varChunk)
				Case sconVerTypeNull
					WScript.Echo "No more data"
					Exit Do
				Case Else
					WScript.Echo "Unexpected VerType", TypeName(varChunk), VarType(varChunk)
					WScript.Quit 1
			End Select
		Loop
		.Close
	End With
	Set objStream = Nothing
	
	Set objBytesToHexString = Nothing
	
End Sub
Call Main
