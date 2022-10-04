Option Explicit

' Date Functions
' | Function       | Description                                                                                             | 
' | -------------- | ------------------------------------------------------------------------------------------------------- | 
' | Date           | A Function, which returns the current system date                                                       | 
' | CDate          | A Function, which converts a given input to Date                                                        | 
' | DateAdd        | A Function, which returns a date to which a specified time interval has been added                      | 
' | DateDiff       | A Function, which returns the difference between two time period                                        | 
' | DatePart       | A Function, which returns a specified part of the given input date value                                | 
' | DateSerial     | A Function, which returns a valid date for the given year, month and date                               | 
' | FormatDateTime | A Function, which formats the date based on the supplied parameters                                     | 
' | IsDate         | A Function, which returns a Boolean Value whether or not the supplied parameter is a date               | 
' | Day            | A Function, which returns an integer between 1 and 31 that represents the day of the specified Date     | 
' | Month          | A Function, which returns an integer between 1 and 12 that represents the month of the specified Date   | 
' | Year           | A Function, which returns an integer that represents the year of the specified Date                     | 
' | MonthName      | A Function, which returns Name of the particular month for the specified date                           | 
' | WeekDay        | A Function, which returns an integer(1 to 7) that represents the day of the week for the specified day. | 
' | WeekDayName    | A Function, which returns the weekday name for the specified day.                                       | 

' Time Functions
' | Function   | Description                                                                                                   | 
' | ---------- | ------------------------------------------------------------------------------------------------------------- | 
' | Now        | A Function, which returns the current system date and Time                                                    | 
' | Hour       | A Function, which returns and integer between 0 and 23 that represents the Hour part of the the given time    | 
' | Minute     | A Function, which returns and integer between 0 and 59 that represents the Minutes part of the the given time | 
' | Second     | A Function, which returns and integer between 0 and 59 that represents the Seconds part of the the given time | 
' | Time       | A Function, which returns the current system time                                                             | 
' | Timer      | A Function, which returns the number of seconds and milliseconds since 12:00 AM                               | 
' | TimeSerial | A Function, which returns the time for the specific input of hour,minute and second                           | 
' | TimeValue  | A Function, which converts the input string to a time format                                                  | 

Sub Test1()
    Dim dtmSample1: dtmSample1 = #2022/10/04 17:13:12#
    WScript.Echo dtmSample1             ' 10/4/2022 5:13:12 PM
    WScript.Echo Second(dtmSample1)     ' 12
    WScript.Echo TypeName(dtmSample1)   ' Date
    Dim dblSample1: dblSample1 = CDbl(dtmSample1)
    WScript.Echo dblSample1     ' 44838.7175
    Dim dtmSample2: dtmSample2 = DateAdd("d", -44837, dtmSample1)
    WScript.Echo dtmSample2     ' 12/31/1899 5:13:12 PM
    Dim dtmSample3: dtmSample3 = #1800/01/01 00:00:01#
    WScript.Echo dtmSample3     ' 1/1/1800 12:00:01 AM
    Dim dblSample3: dblSample3 = CDbl(dtmSample3)
    WScript.Echo dblSample3     ' -36522.0000115741
    Dim dtmSample4: dtmSample4 = #1899/12/30 00:00:00#
    WScript.Echo dtmSample4     ' 12:00:00 AM
    Dim dblSample4: dblSample4 = CDbl(dtmSample4)
    WScript.Echo dblSample4     '0
    WScript.Echo CDbl(#1899/12/30 00:00:01#)    ' 1.15740740740741E-05
    WScript.Echo CDbl(#1899/12/30 00:01:00#)    ' 6.94444444444444E-04
    WScript.Echo CDbl(#1899/12/30 01:00:00#)    ' 4.16666666666667E-02
    WScript.Echo 1/24/60/60 ' 1.15740740740741E-05
    WScript.Echo 1/24/60    ' 6.94444444444444E-04
    WScript.Echo 1/24       ' 4.16666666666667E-02

    ' Learning
    ' Date is the elapsed days from 1899/12/30 00:00:00 in double.
    ' Date older than 0 can be expressed in negative value
    ' Time is a fraction of 1 day.
    ' 1 / 24 is equivalent to 1 hour
    ' 1 / (24 * 60 * 60) is equivalent to 1 second
End Sub
' Test1

Sub Test2()
    WScript.echo CDate("2022/10/04 17:13:12")       ' 10/4/2022 5:13:12 PM
    WScript.echo CDate("10/04/2022 17:13:12")       ' 10/4/2022 5:13:12 PM
    WScript.echo CDate("04/10/2022 17:13:12")       ' 4/10/2022 5:13:12 PM
    WScript.echo CDate("04/10/2022 17:13:12.891")   ' Type mismatch: 'CDate'
    ' Learning
    ' Using CDate, String can be converted to Date.
    ' ****/**/** is interpreted as yyyy/mm/dd
    ' **/**/**** is interpreted as mm/dd/yyyy
    ' fraction of a second is not accepted with Microsoft VBScript runtime error

End Sub
' Test2

' DateDiff(interval,date1,date2[,firstdayofweek[,firstweekofyear]])
' | Parameter       | Description                                                                                 | 
' | --------------- | ------------------------------------------------------------------------------------------- | 
' | interval        | Required. The interval you want to use to calculate the differences between date1 and date2 | 
' |                 | Can take the following values:                                                              | 
' |                 | yyyy - Year                                                                                 | 
' |                 | q - Quarter                                                                                 | 
' |                 | m - Month                                                                                   | 
' |                 | y - Day of year                                                                             | 
' |                 | d - Day                                                                                     | 
' |                 | w - Weekday                                                                                 | 
' |                 | ww - Week of year                                                                           | 
' |                 | h - Hour                                                                                    | 
' |                 | n - Minute                                                                                  | 
' |                 | s - Second                                                                                  | 
' | date1,date2     | Required. Date expressions. Two dates you want to use in the calculation                    | 
' | firstdayofweek  | Optional. Specifies the day of the week.                                                    | 
' |                 | Can take the following values:                                                              | 
' |                 | 0 = vbUseSystemDayOfWeek - Use National Language Support (NLS) API setting                  | 
' |                 | 1 = vbSunday - Sunday (default)                                                             | 
' |                 | 2 = vbMonday - Monday                                                                       | 
' |                 | 3 = vbTuesday - Tuesday                                                                     | 
' |                 | 4 = vbWednesday - Wednesday                                                                 | 
' |                 | 5 = vbThursday - Thursday                                                                   | 
' |                 | 6 = vbFriday - Friday                                                                       | 
' |                 | 7 = vbSaturday - Saturday                                                                   | 
' | firstweekofyear | Optional. Specifies the first week of the year.                                             | 
' |                 | Can take the following values:                                                              | 
' |                 | 0 = vbUseSystem - Use National Language Support (NLS) API setting                           | 
' |                 | 1 = vbFirstJan1 - Start with the week in which January 1 occurs (default)                   | 
' |                 | 2 = vbFirstFourDays - Start with the week that has at least four days in the new year       | 
' |                 | 3 = vbFirstFullWeek - Start with the first full week of the new year                        | 

Sub Test3()
    Dim dtmA: dtmA = #2022/10/04 00:00:00#
    Dim dtmB: dtmB = #2022/10/04 00:00:01#
    Dim dtmC: dtmC = #2022/10/04 00:01:00#
    Dim dtmD: dtmD = #2022/10/04 01:00:00#
    Dim dtmE: dtmE = #2022/10/05 00:00:00#
    Dim dtmF: dtmF = #2022/10/04 00:59:59#
    WScript.Echo dtmE - dtmA, TypeName(dtmE - dtmA) ' 1    Double
    Dim lngSec: lngSec = DateDiff("s", dtmA, dtmB)
    WScript.Echo lngSec, TypeName(lngSec)   ' 1 Long
    lngSec = DateDiff("s", dtmB, dtmA)
    WScript.Echo lngSec, TypeName(lngSec)   ' -1 Long
    lngSec = DateDiff("s", dtmA, dtmC)
    WScript.Echo lngSec, TypeName(lngSec)   ' 60 Long
    lngSec = DateDiff("s", dtmB, dtmE)
    WScript.Echo lngSec, TypeName(lngSec)   ' 86399 Long
    Dim lngHour: lngHour = DateDiff("h", dtmD, dtmE)
    WScript.Echo lngHour, TypeName(lngHour)   ' 23 Long
    lngHour = DateDiff("h", dtmC, dtmE)
    WScript.Echo lngHour, TypeName(lngHour)   ' 24 Long
    lngHour = DateDiff("h", dtmB, dtmE)
    WScript.Echo lngHour, TypeName(lngHour)   ' 24 Long
    lngHour = DateDiff("h", dtmF, dtmE)
    WScript.Echo lngHour, TypeName(lngHour)   ' 24 Long
    WScript.Echo DateDiff("s", "2022/10/04", "1800/10/04")  ' -7005657600
    WScript.Echo DateDiff("s", "2022/10/04", "1000/10/04")  ' -32251219200
    WScript.Echo DateDiff("s", "3000/10/04", "1000/10/04")  ' -63113904000
    WScript.Echo DateDiff("s", "3000/10/04", "100/10/04")   ' -91515139200
    ' Learning
    ' Date - Date = Double
    ' DateDiff for second is Long (either positive or negative). This is the most accurate.
    ' DateDiff for hour is Long, but it gets rounded up. Hour can only be used for estimation.
End Sub
' Test3

' Purpose: Create Date based on date string and time string
' Inputs:
'   ByVal strDate such as 20221004 as in 2022/10/04, expected length is 8
'   ByVal strTime such as 141203 as in 14:12:03, expected length is 6
' Assumptions;
'   Input values are valid numbers in string. For performance, value error is not handled.
' Return: Date
Function dtmFromString(ByVal strDate, ByVal strTime)
    dtmFromString = CDate(Left(strDate, 4) & "/" & Mid(strDate, 5, 2) & "/" & Right(strDate, 2) & " " & _
    Left(strTime, 2) & ":" & Mid(strDate, 3, 2) & ":" & Right(strTime, 2))
End Function

Sub Test4()
    Dim dtmReturn, strDate, strTime
    strDate = "20221004"
    strTime = "170059"
    dtmReturn = dtmFromString(strDate, strTime)
    WScript.Echo dtmReturn, TypeName(dtmReturn) ' 10/4/2022 5:22:59 PM Date
End Sub
Test4

