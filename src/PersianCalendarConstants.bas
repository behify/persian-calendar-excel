Attribute VB_Name = "PersianCalendarConstants"
' Persian Calendar Conversion Package for Excel
' Based on .NET PersianCalendar algorithm
' Module 1: Constants and Core Variables

Option Explicit

' Persian calendar constants
Public Const PERSIAN_ERA As Long = 1
Public Const PERSIAN_EPOCH_YEAR As Long = 622 ' Gregorian year of Persian calendar start
Public Const MIN_PERSIAN_YEAR As Long = 1
Public Const MAX_PERSIAN_YEAR As Long = 9378

' Days in each month for Persian calendar
Private Const DAYS_IN_MONTH_1_TO_6 As Long = 31 ' Farvardin to Shahrivar
Private Const DAYS_IN_MONTH_7_TO_11 As Long = 30 ' Mehr to Bahman
Private Const DAYS_IN_MONTH_12_NORMAL As Long = 29 ' Esfand in normal year
Private Const DAYS_IN_MONTH_12_LEAP As Long = 30 ' Esfand in leap year

' Persian month names array
Private PersianMonthNames As Variant

' Initialize Persian month names
Private Sub InitializePersianMonthNames()
    PersianMonthNames = Array("ÝÑæÑÏíä", "ÇÑÏíÈåÔÊ", "ÎÑÏÇÏ", "ÊíÑ", _
                             "ãÑÏÇÏ", "ÔåÑíæÑ", "ãåÑ", "ÂÈÇä", _
                             "ÂÐÑ", "Ïí", "Èåãä", "ÇÓÝäÏ")
End Sub

' Get Persian month name by index (1-12)
Public Function GetPersianMonthName(monthIndex As Long) As String
    If IsEmpty(PersianMonthNames) Then InitializePersianMonthNames
    
    If monthIndex >= 1 And monthIndex <= 12 Then
        GetPersianMonthName = PersianMonthNames(monthIndex - 1)
    Else
        GetPersianMonthName = ""
    End If
End Function

' Check if Persian year is leap year
Public Function IsPersianLeapYear(persianYear As Long) As Boolean
    ' 33-year cycle pattern implementation
    Dim cycle33 As Long
    Dim yearInCycle As Long
    
    cycle33 = (persianYear - 1) \ 33
    yearInCycle = (persianYear - 1) Mod 33
    
    ' Leap years in 33-year cycle: 1,5,9,13,17,22,26,30
    Select Case yearInCycle
        Case 0, 4, 8, 12, 16, 21, 25, 29
            IsPersianLeapYear = True
        Case Else
            IsPersianLeapYear = False
    End Select
End Function

' Get days in Persian month
Public Function GetDaysInPersianMonth(persianYear As Long, persianMonth As Long) As Long
    Select Case persianMonth
        Case 1 To 6
            GetDaysInPersianMonth = DAYS_IN_MONTH_1_TO_6
        Case 7 To 11
            GetDaysInPersianMonth = DAYS_IN_MONTH_7_TO_11
        Case 12
            If IsPersianLeapYear(persianYear) Then
                GetDaysInPersianMonth = DAYS_IN_MONTH_12_LEAP
            Else
                GetDaysInPersianMonth = DAYS_IN_MONTH_12_NORMAL
            End If
        Case Else
            GetDaysInPersianMonth = 0 ' Invalid month
    End Select
End Function

