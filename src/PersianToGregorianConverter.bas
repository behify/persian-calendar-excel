Attribute VB_Name = "PersianToGregorianConverter"
' Persian Calendar Conversion Package for Excel
' Module 2: Persian to Gregorian Date Conversion

Option Explicit

' Convert Persian date to Gregorian date
Public Function PersianToGregorian(persianYear As Long, persianMonth As Long, persianDay As Long) As Date
    Dim totalDays As Long
    Dim gregorianDate As Date
    
    ' Validate input
    If Not IsValidPersianDate(persianYear, persianMonth, persianDay) Then
        Err.Raise 5, "PersianToGregorian", "Invalid Persian date"
    End If
    
    ' Calculate total days from Persian epoch
    totalDays = GetTotalDaysFromPersianEpoch(persianYear, persianMonth, persianDay)
    
    ' Convert to Gregorian date (Persian epoch: March 21, 622 AD)
    gregorianDate = DateSerial(PERSIAN_EPOCH_YEAR, 3, 21) + totalDays - 1
    
    PersianToGregorian = gregorianDate
End Function

' Validate Persian date
Public Function IsValidPersianDate(persianYear As Long, persianMonth As Long, persianDay As Long) As Boolean
    If persianYear < MIN_PERSIAN_YEAR Or persianYear > MAX_PERSIAN_YEAR Then
        IsValidPersianDate = False
        Exit Function
    End If
    
    If persianMonth < 1 Or persianMonth > 12 Then
        IsValidPersianDate = False
        Exit Function
    End If
    
    If persianDay < 1 Or persianDay > GetDaysInPersianMonth(persianYear, persianMonth) Then
        IsValidPersianDate = False
        Exit Function
    End If
    
    IsValidPersianDate = True
End Function

' Calculate total days from Persian epoch (1/1/1 Persian = 22/3/622 Gregorian)
Private Function GetTotalDaysFromPersianEpoch(persianYear As Long, persianMonth As Long, persianDay As Long) As Long
    Dim totalDays As Long
    Dim yearIndex As Long
    Dim monthIndex As Long
    
    ' Add days for complete years
    For yearIndex = 1 To persianYear - 1
        If IsPersianLeapYear(yearIndex) Then
            totalDays = totalDays + 366
        Else
            totalDays = totalDays + 365
        End If
    Next yearIndex
    
    ' Add days for complete months in current year
    For monthIndex = 1 To persianMonth - 1
        totalDays = totalDays + GetDaysInPersianMonth(persianYear, monthIndex)
    Next monthIndex
    
    ' Add remaining days
    totalDays = totalDays + persianDay
    
    GetTotalDaysFromPersianEpoch = totalDays
End Function

' Excel worksheet function: Convert Persian date string to Gregorian
Public Function PERSIAN_TO_GREGORIAN(persianDateStr As String) As Date
    Dim dateParts As Variant
    Dim persianYear As Long, persianMonth As Long, persianDay As Long
    
    ' Parse date string (format: YYYY/MM/DD or YYYY-MM-DD)
    persianDateStr = Replace(persianDateStr, "-", "/")
    dateParts = Split(persianDateStr, "/")
    
    If UBound(dateParts) <> 2 Then
        Err.Raise 5, "PERSIAN_TO_GREGORIAN", "Invalid date format. Use YYYY/MM/DD"
    End If
    
    persianYear = CLng(dateParts(0))
    persianMonth = CLng(dateParts(1))
    persianDay = CLng(dateParts(2))
    
    PERSIAN_TO_GREGORIAN = PersianToGregorian(persianYear, persianMonth, persianDay)
End Function

' Excel worksheet function: Convert Persian date components to Gregorian
Public Function PERSIAN_DATE_TO_GREGORIAN(persianYear As Long, persianMonth As Long, persianDay As Long) As Date
    PERSIAN_DATE_TO_GREGORIAN = PersianToGregorian(persianYear, persianMonth, persianDay)
End Function
