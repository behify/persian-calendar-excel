Attribute VB_Name = "PersianCalendarTests"
' Persian Calendar Conversion Package for Excel
' Module 5: Test Functions and Setup

Option Explicit

' Test all conversion functions
Public Sub TestPersianCalendarFunctions()
    Dim testResults As String
    Dim passed As Long, failed As Long
    
    testResults = "Persian Calendar Package Test Results" & vbCrLf & _
                 String(50, "=") & vbCrLf & vbCrLf
    
    ' Test 1: Basic Persian to Gregorian conversion
    If TestPersianToGregorian() Then
        testResults = testResults & "? Persian to Gregorian conversion: PASSED" & vbCrLf
        passed = passed + 1
    Else
        testResults = testResults & "? Persian to Gregorian conversion: FAILED" & vbCrLf
        failed = failed + 1
    End If
    
    ' Test 2: Basic Gregorian to Persian conversion
    If TestGregorianToPersian() Then
        testResults = testResults & "? Gregorian to Persian conversion: PASSED" & vbCrLf
        passed = passed + 1
    Else
        testResults = testResults & "? Gregorian to Persian conversion: FAILED" & vbCrLf
        failed = failed + 1
    End If
    
    ' Test 3: Leap year detection
    If TestLeapYearDetection() Then
        testResults = testResults & "? Leap year detection: PASSED" & vbCrLf
        passed = passed + 1
    Else
        testResults = testResults & "? Leap year detection: FAILED" & vbCrLf
        failed = failed + 1
    End If
    
    ' Test 4: Date validation
    If TestDateValidation() Then
        testResults = testResults & "? Date validation: PASSED" & vbCrLf
        passed = passed + 1
    Else
        testResults = testResults & "? Date validation: FAILED" & vbCrLf
        failed = failed + 1
    End If
    
    ' Test 5: Helper functions
    If TestHelperFunctions() Then
        testResults = testResults & "? Helper functions: PASSED" & vbCrLf
        passed = passed + 1
    Else
        testResults = testResults & "? Helper functions: FAILED" & vbCrLf
        failed = failed + 1
    End If
    
    testResults = testResults & vbCrLf & String(50, "=") & vbCrLf & _
                 "Total Tests: " & (passed + failed) & vbCrLf & _
                 "Passed: " & passed & vbCrLf & _
                 "Failed: " & failed & vbCrLf
    
    ' Display results
    MsgBox testResults, vbInformation, "Test Results"
    
    ' Write to immediate window if available
    Debug.Print testResults
End Sub

' Test Persian to Gregorian conversion
Private Function TestPersianToGregorian() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testDate As Date
    Dim expectedDate As Date
    
    ' Test known conversion: 1400/01/01 should be 2021/03/21
    testDate = PersianToGregorian(1400, 1, 1)
    expectedDate = DateSerial(2021, 3, 21)
    
    TestPersianToGregorian = (testDate = expectedDate)
    Exit Function
    
ErrorHandler:
    TestPersianToGregorian = False
End Function

' Test Gregorian to Persian conversion
Private Function TestGregorianToPersian() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testDate As Date
    Dim persianResult As persianDate
    
    ' Test known conversion: 2021/03/21 should be 1400/01/01
    testDate = DateSerial(2021, 3, 21)
    persianResult = GregorianToPersian(testDate)
    
    TestGregorianToPersian = (persianResult.Year = 1400 And _
                             persianResult.Month = 1 And _
                             persianResult.Day = 1)
    Exit Function
    
ErrorHandler:
    TestGregorianToPersian = False
End Function

' Test leap year detection
Private Function TestLeapYearDetection() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test known leap years and non-leap years
    Dim testPassed As Boolean
    testPassed = True
    
    ' 1399 is leap year
    If Not IsPersianLeapYear(1399) Then testPassed = False
    
    ' 1400 is not leap year
    If IsPersianLeapYear(1400) Then testPassed = False
    
    ' 1403 is leap year
    If Not IsPersianLeapYear(1403) Then testPassed = False
    
    TestLeapYearDetection = testPassed
    Exit Function
    
ErrorHandler:
    TestLeapYearDetection = False
End Function

' Test date validation
Private Function TestDateValidation() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testPassed As Boolean
    testPassed = True
    
    ' Valid date should pass
    If Not IsValidPersianDate(1400, 1, 1) Then testPassed = False
    
    ' Invalid month should fail
    If IsValidPersianDate(1400, 13, 1) Then testPassed = False
    
    ' Invalid day should fail
    If IsValidPersianDate(1400, 12, 31) Then testPassed = False
    
    TestDateValidation = testPassed
    Exit Function
    
ErrorHandler:
    TestDateValidation = False
End Function

' Test helper functions
Private Function TestHelperFunctions() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testPassed As Boolean
    testPassed = True
    
    ' Test month name function (check for Persian month name)
    If GetPersianMonthName(1) <> "›—Ê—œÌ‰" Then
        Debug.Print "Month name test failed. Got: " & GetPersianMonthName(1) & " Expected: ›—Ê—œÌ‰"
        testPassed = False
    End If
    
    ' Test days in month function
    If GetDaysInPersianMonth(1400, 1) <> 31 Then
        Debug.Print "Days in month 1 test failed. Got: " & GetDaysInPersianMonth(1400, 1)
        testPassed = False
    End If
    If GetDaysInPersianMonth(1400, 12) <> 29 Then
        Debug.Print "Days in month 12 (non-leap) test failed. Got: " & GetDaysInPersianMonth(1400, 12)
        testPassed = False
    End If
    If GetDaysInPersianMonth(1399, 12) <> 30 Then
        Debug.Print "Days in month 12 (leap) test failed. Got: " & GetDaysInPersianMonth(1399, 12)
        testPassed = False
    End If
    
    TestHelperFunctions = testPassed
    Exit Function
    
ErrorHandler:
    TestHelperFunctions = False
End Function

' Generate sample data for testing Excel functions
Public Sub GenerateSampleData()
    Dim ws As Worksheet
    Dim i As Long
    
    ' Create new worksheet for samples
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "Persian Calendar Sample"
    
    ' Headers
    ws.Cells(1, 1).Value = "Gregorian Date"
    ws.Cells(1, 2).Value = "Persian Date"
    ws.Cells(1, 3).Value = "Persian Formatted"
    ws.Cells(1, 4).Value = "Persian Weekday"
    ws.Cells(1, 5).Value = "Is Leap Year"
    
    ' Sample data for current month
    For i = 1 To 30
        Dim sampleDate As Date
        sampleDate = DateSerial(2024, 1, i)
        
        ws.Cells(i + 1, 1).Value = sampleDate
        ws.Cells(i + 1, 2).Value = "=GREGORIAN_TO_PERSIAN(A" & (i + 1) & ")"
        ws.Cells(i + 1, 3).Value = "=PERSIAN_DATE_FORMATTED(A" & (i + 1) & ")"
        ws.Cells(i + 1, 4).Value = "=PERSIAN_WEEKDAY_NAME(A" & (i + 1) & ")"
        ws.Cells(i + 1, 5).Value = "=IS_PERSIAN_LEAP_YEAR(PERSIAN_YEAR(A" & (i + 1) & "))"
    Next i
    
    ' Format columns
    ws.Columns("A:E").AutoFit
    ws.Range("A1:E1").Font.Bold = True
    
    MsgBox "Sample data generated in 'Persian Calendar Sample' worksheet", vbInformation
End Sub

' Setup instructions as comment block
'
' INSTALLATION INSTRUCTIONS:
' =========================
' 1. Open Excel and press Alt+F11 to open VBA Editor
' 2. Insert -> Module (repeat 5 times to create 5 modules)
' 3. Copy each module code into separate modules
' 4. Save the workbook as .xlsm (macro-enabled) format
' 5. Run TestPersianCalendarFunctions() to verify installation
' 6. Run GenerateSampleData() to see usage examples
'
' USAGE IN EXCEL WORKSHEETS:
' ==========================
' Available Functions:
' - PERSIAN_TO_GREGORIAN("1400/01/01") - Convert Persian to Gregorian
' - GREGORIAN_TO_PERSIAN(A1) - Convert Gregorian to Persian
' - PERSIAN_YEAR(A1) - Get Persian year from Gregorian date
' - PERSIAN_MONTH(A1) - Get Persian month from Gregorian date
' - PERSIAN_DAY(A1) - Get Persian day from Gregorian date
' - PERSIAN_MONTH_NAME(A1) - Get Persian month name
' - PERSIAN_DATE_FORMATTED(A1) - Get formatted Persian date
' - TODAY_PERSIAN() - Get current date in Persian
' - IS_PERSIAN_LEAP_YEAR(1400) - Check if year is leap
' - ADD_DAYS_TO_PERSIAN("1400/01/01", 10) - Add days to Persian date
' - PERSIAN_DATE_DIFF("1400/01/01", "1400/01/10") - Calculate date difference
'

