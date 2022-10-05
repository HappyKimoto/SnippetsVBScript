Option Explicit

' Purpose: Pad 8 zeros to the passed number
' Inputs: intNum - an integer number
' Assumptions: the number of digit is 8 or less in the passed number.
' Returns: string expression of an integer with zero padding
Function strPadZero8(ByVal intNum)
	strPadZero8 = Right("00000000" & CStr(intNum), 8)
End Function


