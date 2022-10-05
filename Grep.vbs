Option Explicit

' Purpose: Grep match with group globally
' Inputs: 
'   ByVal strData as search string
'   ByVal strPattern as search pattern
'   ByVal intGroupUBound as upper bound of group within the search pattern
' Effects:
'   If called multiple times, creating RegExp object will be created multiple times.
' Return: Array of submatches arrays
Function varMatchesSingleGlobal(ByVal strData, ByVal strPattern, ByVal intGroupUBound)
    Dim objMC, intMcIdx, intGrpCnt, varReturn, intGrpIdx, varRecord
    Dim objRgx: Set objRgx = New RegExp
    With objRgx
        .Pattern = strPattern
        .Global = True
        Set objMC = .Execute(strData)
    End With
    ' WScript.Echo "objMC.Count=" & objMC.Count
    ReDim varReturn(objMC.Count - 1)
    If objMC.Count > 0 Then
        For intMcIdx = 0 To objMC.Count - 1
            ' WScript.Echo "intMcIdx=" & intMcIdx
            ReDim varRecord(intGroupUBound) ' Reset the row/record array
            For intGrpIdx = 0 To intGroupUBound
                ' WScript.Echo "intGrpIdx=" & intGrpIdx
                varRecord(intGrpIdx) = objMC(intMcIdx).SubMatches(intGrpIdx)
            Next
            varReturn(intMcIdx) = varRecord ' Append the record
        Next
    End If
    varMatchesSingleGlobal = varReturn  ' Return the array
    Set objRgx = Nothing    ' Destroy the objet
End Function

' Sub Test()
'     Dim x, i, j
'     x = varMatchesSingleGlobal("1 + 2 = 3, 2 + 3 = 5, 5 - 2 = 3", "(\d) \+ (\d)", 1)
'     For i = LBound(x) to UBound(x)
'         WScript.Echo Join(x(i), vbTab)
'     Next
' End Sub
' Test

' Purpose: Grep match with group expression globally
' Inputs: 
'   ByVal strData as search string
'   ByRef objRgx as regular expression object. 
'       - Global should be set True. 
'       - objRgx is passed by reference and this function will not be destroyed.
'   ByVal intGroupUBound as upper bound of group within the search pattern
' Effects:
'   Even if this function is called multiple times, this function will not waste resource.
' Return: Array of submatches arrays
Function varMatchesSingleGlobal2(ByVal strData, ByRef objRgx, ByVal intGroupUBound)
    Dim objMC, intMcIdx, intGrpCnt, varReturn, intGrpIdx, varRecord
    Set objMC = objRgx.Execute(strData)    
    ReDim varReturn(objMC.Count - 1)
    If objMC.Count > 0 Then
        For intMcIdx = 0 To objMC.Count - 1
            ReDim varRecord(intGroupUBound)
            For intGrpIdx = 0 To intGroupUBound
                varRecord(intGrpIdx) = objMC(intMcIdx).SubMatches(intGrpIdx)
            Next
            varReturn(intMcIdx) = varRecord ' Append record
        Next
    End If
    varMatchesSingleGlobal2 = varReturn ' Return array
End Function

' Sub Test()
'     Dim objRgx: Set objRgx = New RegExp
'     With objRgx
'         .Pattern = "(\d) \+ (\d)"
'         .Global = True
'     End With
'     Dim x, i, j
'     x = varMatchesSingleGlobal2("1 + 2 = 3, 2 + 3 = 5, 5 - 2 = 3", objRgx, 1)
'     For i = LBound(x) to UBound(x)
'         WScript.Echo Join(x(i), vbTab)
'     Next
' End Sub
' Test

' Purpose: Create RegExp object
' Inputs: strPattern - search pattern with group expressions
' Return: RegExp object to be executed from the caller
Function objRgxGlobalGroup(ByVal strPattern)
    Dim objRgx: Set objRgx = New RegExp
    With objRgx
        .Pattern = strPattern
        .Global = True
    End With
    Set objRgxGlobalGroup = objRgx
    Set objRgx = Nothing
End Function

' Sub Test()
'     Dim objRgx
'     Set objRgx = objRgxGlobalGroup("(\d) \+ (\d)")
'     Dim x, i, j
'     x = varMatchesSingleGlobal2("1 + 2 = 3, 2 + 3 = 5, 5 - 2 = 3", objRgx, 1)
'     For i = LBound(x) to UBound(x)
'         WScript.Echo Join(x(i), vbTab)
'     Next
' End Sub
' Test

Function strArrayTableToString(ByRef varTbl)
	Dim intRec, varLines
	ReDim varLines(UBound(varTbl))
	For intRec = LBound(varTbl) To UBound(varTbl)
		varLines(intRec) = Join(varTbl(intRec), vbTab)
	Next
	strArrayTableToString = Join(varLines, vbCrlf)
End Function

' | Constant      | Value                        | Description                                                                      | 
' | ------------- | ---------------------------- | -------------------------------------------------------------------------------- | 
' | vbCr          | Chr(13)                      | Carriage return.                                                                 | 
' | VbCrLf        | Chr(13) & Chr(10)            | Carriage return–linefeed combination.                                            | 
' | vbFormFeed    | Chr(12)                      | Form feed; not useful in Microsoft Windows.                                      | 
' | vbLf          | Chr(10)                      | Line feed.                                                                       | 
' | vbNewLine     | Chr(13) & Chr(10) or Chr(10) | Platform-specific newline character; whatever is appropriate for the platform.   | 
' | vbNullChar    | Chr(0)                       | Character having the value 0.                                                    | 
' | vbNullString  | String having value 0        | Not the same as a zero-length string (""); used for calling external procedures. | 
' | vbTab         | Chr(9)                       | Horizontal tab.                                                                  | 
' | vbVerticalTab | Chr(11)                      | Vertical tab; not useful in Microsoft Windows.                                   | 

Function strArrayTableToString2(ByRef varTbl, ByVal strSepCol, ByVal strSepRow)
	Dim intRec, varLines
	ReDim varLines(UBound(varTbl))
	For intRec = LBound(varTbl) To UBound(varTbl)
		varLines(intRec) = Join(varTbl(intRec), strSepCol)
	Next
	strArrayTableToString2 = Join(varLines, strSepRow)
End Function

Function strArrayTableToString3(ByRef varTbl, ByVal strSepCol)
	Dim intRec, varLines
	ReDim varLines(UBound(varTbl))
	For intRec = LBound(varTbl) To UBound(varTbl)
		varLines(intRec) = Join(varTbl(intRec), strSepCol)
	Next
	strArrayTableToString3 = Join(varLines, vbCrlf)
End Function

Sub Test()
    Dim objRgx
    Set objRgx = objRgxGlobalGroup("(\d) \+ (\d)")
    Dim x, i, j
    x = varMatchesSingleGlobal2("1 + 2 = 3, 2 + 3 = 5, 5 - 2 = 3", objRgx, 1)
    WScript.Echo strArrayTableToString(x)
    WScript.Echo strArrayTableToString2(x, Chr(9), Chr(13) & Chr(10))
    WScript.Echo strArrayTableToString3(x, Chr(44))
End Sub
Test