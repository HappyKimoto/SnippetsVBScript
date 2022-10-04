Option Explicit

' Purpose: Simulate Selection Box
' Inputs: Type in numbers 0 to 2
' Effect: Associated sub procedure will be called

Sub Sub1()
    MsgBox "A"
End Sub

Sub Sub2()
    MsgBox "B"
End Sub

Sub Sub3()
    MsgBox "C"
End Sub

Dim varOptions, intSelect, strOptions, i
ReDim varOptions(2)
Const conNumber = 0
Const conName = 1
Const conTitle = "Select by Number"
strOptions = ""
varOptions(0) = Array(0, "Sub1")
varOptions(1) = Array(1, "Sub2")
varOptions(2) = Array(2, "Sub3")

For i = LBound(varOptions) To UBound(varOptions)
    strOptions = strOptions &  varOptions(i)(conNumber) & vbTab & varOptions(i)(conName) & vbCrlf
Next

intSelect = CInt(InputBox(strOptions, conTitle))
Select Case intSelect
Case 0
    Sub1
Case 1
    Sub2
Case 2
    Sub3
Case Else
    MsgBox "Invalid Entry"
End Select
