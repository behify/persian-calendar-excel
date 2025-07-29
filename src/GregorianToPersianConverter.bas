Attribute VB_Name = "GregorianToPersianConverter"
' Persian Calendar Conversion Package for Excel
' Module 3: Gregorian to Persian Date Conversion

Option Explicit

' Type for Persian date components
Public Type persianDate
    Year As Long
    Month As Long
    Day As Long
End Type

' Convert Gregorian date to Persian date
Public Function GregorianToPersian(gregorianDate As Date) As persianDate
    Dim persianEpochDate As Date
    Dim daysSinceEpoch As Long
    Dim persianResult As persianDate
    
    ' Persian epoch: March 21, 622 AD
    persianEpochDate = DateSerial(PERSIAN_EPOCH_YEAR, 3, 21)
    
    ' Calculate days since Persian epoch
    daysSinceEpoch = gregorianDate - persianEpochDate + 1
    
    ' Convert days to Persian date components
    persianResult = DaysToPersianDate(daysSinceEpoch)
    
    GregorianToPersian = persianResult
End Function

' Convert total days to Persian date components
Private Function DaysToPersianDate(totalDays As Long) As persianDate
    Dim persianDate As persianDate
    Dim remainingDays As Long
    Dim currentYear As Long
    Dim currentMonth As Long
    Dim daysInCurrentYear As Long
    Dim daysInCurrentMonth As Long
    
    remainingDays = totalDays
    currentYear = 1
    
    ' Find the year
    Do While remainingDays > 0
        If IsPersianLeapYear(currentYear) Then
            daysInCurrentYear = 366
        Else
            daysInCurrentYear = 365
        End If
        
        If remainingDays > daysInCurrentYear Then
            remainingDays = remainingDays - daysInCurrentYear
            currentYear = currentYear + 1
        Else
            Exit Do
        End If
    Loop
    
    persianDate.Year = currentYear
    currentMonth = 1
    
    ' Find the month
    Do While remainingDays > 0 And currentMonth <= 12
        daysInCurrentMonth = GetDaysInPersianMonth(currentYear, currentMonth)
        
        If remainingDays > daysInCurrentMonth Then
            remainingDays = remainingDays - daysInCurrentMonth
            currentMonth = currentMonth + 1
        Else
            Exit Do
        End If
    Loop
    
    persianDate.Month = currentMonth
    persianDate.Day = remainingDays
    
    DaysToPersianDate = persianDate
End Function

' Excel worksheet function: Convert Gregorian date to Persian date string
Public Function GREGORIAN_TO_PERSIAN(gregorianDate As Date) As String
    Dim persianDate As persianDate
    
    persianDate = GregorianToPersian(gregorianDate)
    
    GREGORIAN_TO_PERSIAN = Format(persianDate.Year, "0000") & "/" & _
                          Format(persianDate.Month, "00") & "/" & _
                          Format(persianDate.Day, "00")
End Function

' Excel worksheet function: Get Persian year from Gregorian date
Public Function PERSIAN_YEAR(gregorianDate As Date) As Long
    Dim persianDate As persianDate
    persianDate = GregorianToPersian(gregorianDate)
    PERSIAN_YEAR = persianDate.Year
End Function

' Excel worksheet function: Get Persian month from Gregorian date
Public Function PERSIAN_MONTH(gregorianDate As Date) As Long
    Dim persianDate As persianDate
    persianDate = GregorianToPersian(gregorianDate)
    PERSIAN_MONTH = persianDate.Month
End Function

' Excel worksheet function: Get Persian day from Gregorian date
Public Function PERSIAN_DAY(gregorianDate As Date) As Long
    Dim persianDate As persianDate
    persianDate = GregorianToPersian(gregorianDate)
    PERSIAN_DAY = persianDate.Day
End Function

' Excel worksheet function: Get Persian month name from Gregorian date
Public Function PERSIAN_MONTH_NAME(gregorianDate As Date) As String
    Dim persianDate As persianDate
    persianDate = GregorianToPersian(gregorianDate)
    PERSIAN_MONTH_NAME = GetPersianMonthName(persianDate.Month)
End Function

' Excel worksheet function: Format Persian date with month name
Public Function PERSIAN_DATE_FORMATTED(gregorianDate As Date) As String
    Dim persianDate As persianDate
    persianDate = GregorianToPersian(gregorianDate)
    
    PERSIAN_DATE_FORMATTED = persianDate.Day & " " & _
                            GetPersianMonthName(persianDate.Month) & " " & _
                            persianDate.Year
End Function

