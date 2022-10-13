Option Explicit

' Purpose: Pad 8 zeros to the passed number
' Inputs: intNum - an integer number
' Assumptions: the number of digit is 8 or less in the passed number.
' Returns: string expression of an integer with zero padding
Function strPadZero8(ByVal intNum)
	strPadZero8 = Right("00000000" & CStr(intNum), 8)
End Function

' Purpose: Pad 2 zeros to the passed number
' Assumptions: Can be used for integer value for month, day, hour, minute, second
Function strPadZero2(ByVal intNum)
	strPadZero2 = Right("00" & CStr(intNum), 2)
End Function

' Purpose: Pad 4 zeros to the passed number
' Assumptions: Can be used for integer value for year
Function strPadZero4(ByVal intNum)
	strPadZero4 = Right("0000" & CStr(intNum), 4)
End Function