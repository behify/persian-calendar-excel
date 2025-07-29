Attribute VB_Name = "PersianCalendarHelpers"
' Persian Calendar Conversion Package for Excel
' Module 4: Helper Functions and Additional Utilities

Option Explicit

' Excel worksheet function: Check if Persian year is leap year
Public Function IS_PERSIAN_LEAP_YEAR(persianYear As Long) As Boolean
    IS_PERSIAN_LEAP_YEAR = IsPersianLeapYear(persianYear)
End Function

' Excel worksheet function: Get days in Persian month
Public Function PERSIAN_DAYS_IN_MONTH(persianYear As Long, persianMonth As Long) As Long
    PERSIAN_DAYS_IN_MONTH = GetDaysInPersianMonth(persianYear, persianMonth)
End Function

' Excel worksheet function: Get Persian month name by number
Public Function PERSIAN_MONTH_NAME_BY_NUMBER(monthNumber As Long) As String
    PERSIAN_MONTH_NAME_BY_NUMBER = GetPersianMonthName(monthNumber)
End Function

' Excel worksheet function: Add days to Persian date
Public Function ADD_DAYS_TO_PERSIAN(persianDateStr As String, daysToAdd As Long) As String
    Dim dateParts As Variant
    Dim persianYear As Long, persianMonth As Long, persianDay As Long
    Dim gregorianDate As Date
    Dim newGregorianDate As Date
    Dim newPersianDate As persianDate
    
    ' Parse Persian date
    persianDateStr = Replace(persianDateStr, "-", "/")
    dateParts = Split(persianDateStr, "/")
    
    If UBound(dateParts) <> 2 Then
        Err.Raise 5, "ADD_DAYS_TO_PERSIAN", "Invalid date format. Use YYYY/MM/DD"
    End If
    
    persianYear = CLng(dateParts(0))
    persianMonth = CLng(dateParts(1))
    persianDay = CLng(dateParts(2))
    
    ' Convert to Gregorian, add days, convert back
    gregorianDate = PersianToGregorian(persianYear, persianMonth, persianDay)
    newGregorianDate = gregorianDate + daysToAdd
    newPersianDate = GregorianToPersian(newGregorianDate)
    
    ADD_DAYS_TO_PERSIAN = Format(newPersianDate.Year, "0000") & "/" & _
                         Format(newPersianDate.Month, "00") & "/" & _
                         Format(newPersianDate.Day, "00")
End Function

' Excel worksheet function: Calculate difference between two Persian dates in days
Public Function PERSIAN_DATE_DIFF(startPersianDate As String, endPersianDate As String) As Long
    Dim startGregorian As Date, endGregorian As Date
    
    startGregorian = PERSIAN_TO_GREGORIAN(startPersianDate)
    endGregorian = PERSIAN_TO_GREGORIAN(endPersianDate)
    
    PERSIAN_DATE_DIFF = endGregorian - startGregorian
End Function

' Excel worksheet function: Get current Persian date
Public Function TODAY_PERSIAN() As String
    Dim currentDate As Date
    Dim persianDate As persianDate
    
    currentDate = Date
    persianDate = GregorianToPersian(currentDate)
    
    TODAY_PERSIAN = Format(persianDate.Year, "0000") & "/" & _
                   Format(persianDate.Month, "00") & "/" & _
                   Format(persianDate.Day, "00")
End Function

' Excel worksheet function: Get current Persian date formatted
Public Function TODAY_PERSIAN_FORMATTED() As String
    Dim currentDate As Date
    
    currentDate = Date
    TODAY_PERSIAN_FORMATTED = PERSIAN_DATE_FORMATTED(currentDate)
End Function

' Excel worksheet function: Get Persian weekday name
Public Function PERSIAN_WEEKDAY_NAME(gregorianDate As Date) As String
    Dim weekdayNumber As Long
    Dim persianWeekdays As Variant
    
    ' Initialize Persian weekday names (Saturday = 1)
    persianWeekdays = Array("‘‰»Â", "Ìò‘‰»Â", "œÊ‘‰»Â", "”Âù‘‰»Â", _
                           "çÂ«—‘‰»Â", "Å‰Ãù‘‰»Â", "Ã„⁄Â")
    
    ' Convert VBA weekday (Sunday = 1) to Persian weekday (Saturday = 1)
    weekdayNumber = Weekday(gregorianDate, vbSunday)
    Select Case weekdayNumber
        Case 1: weekdayNumber = 2 ' Sunday -> Ìò‘‰»Â
        Case 2: weekdayNumber = 3 ' Monday -> œÊ‘‰»Â
        Case 3: weekdayNumber = 4 ' Tuesday -> ”Âù‘‰»Â
        Case 4: weekdayNumber = 5 ' Wednesday -> çÂ«—‘‰»Â
        Case 5: weekdayNumber = 6 ' Thursday -> Å‰Ãù‘‰»Â
        Case 6: weekdayNumber = 7 ' Friday -> Ã„⁄Â
        Case 7: weekdayNumber = 1 ' Saturday -> ‘‰»Â
    End Select
    
    PERSIAN_WEEKDAY_NAME = persianWeekdays(weekdayNumber - 1)
End Function

' Excel worksheet function: Validate Persian date string
Public Function IS_VALID_PERSIAN_DATE(persianDateStr As String) As Boolean
    Dim dateParts As Variant
    Dim persianYear As Long, persianMonth As Long, persianDay As Long
    
    On Error GoTo ErrorHandler
    
    ' Parse date string
    persianDateStr = Replace(persianDateStr, "-", "/")
    dateParts = Split(persianDateStr, "/")
    
    If UBound(dateParts) <> 2 Then
        IS_VALID_PERSIAN_DATE = False
        Exit Function
    End If
    
    persianYear = CLng(dateParts(0))
    persianMonth = CLng(dateParts(1))
    persianDay = CLng(dateParts(2))
    
    IS_VALID_PERSIAN_DATE = IsValidPersianDate(persianYear, persianMonth, persianDay)
    Exit Function
    
ErrorHandler:
    IS_VALID_PERSIAN_DATE = False
End Function

