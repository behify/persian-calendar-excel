# API Documentation

<div align="center">
  <img src="https://raw.githubusercontent.com/behify/persian-calendar-excel/main/assets/logo.png" alt="Persian Calendar Excel Logo" width="150">
</div>

## Core Functions

### Date Conversion Functions

#### `PersianToGregorian(persianYear, persianMonth, persianDay)`
Convert Persian date to Gregorian date

**Parameters:**
- `persianYear` (Long): Persian year (1-9378)
- `persianMonth` (Long): Persian month (1-12)
- `persianDay` (Long): Persian day (1-31)

**Returns:** Date - Gregorian date

**Example:**
```vba
Dim gregorianDate As Date
gregorianDate = PersianToGregorian(1403, 5, 8)
```

#### `GregorianToPersian(gregorianDate)`
Convert Gregorian date to Persian date

**Parameters:**
- `gregorianDate` (Date): Gregorian date

**Returns:** PersianDate - Contains Year, Month, Day

**Example:**
```vba
Dim persianDate As PersianDate
persianDate = GregorianToPersian(Date)
```

### Excel Worksheet Functions

#### Conversion Functions

| Function | Description | Syntax | Example |
|----------|-------------|---------|---------|
| `PERSIAN_TO_GREGORIAN` | Convert Persian string to Gregorian | `PERSIAN_TO_GREGORIAN(dateString)` | `=PERSIAN_TO_GREGORIAN("1403/05/08")` |
| `PERSIAN_DATE_TO_GREGORIAN` | Convert Persian components to Gregorian | `PERSIAN_DATE_TO_GREGORIAN(year, month, day)` | `=PERSIAN_DATE_TO_GREGORIAN(1403, 5, 8)` |
| `GREGORIAN_TO_PERSIAN` | Convert Gregorian to Persian string | `GREGORIAN_TO_PERSIAN(date)` | `=GREGORIAN_TO_PERSIAN(TODAY())` |

#### Component Extraction Functions

| Function | Description | Returns | Example |
|----------|-------------|---------|---------|
| `PERSIAN_YEAR` | Extract Persian year | Long | `=PERSIAN_YEAR(A1)` |
| `PERSIAN_MONTH` | Extract Persian month | Long | `=PERSIAN_MONTH(A1)` |
| `PERSIAN_DAY` | Extract Persian day | Long | `=PERSIAN_DAY(A1)` |

#### Formatting Functions

| Function | Description | Format | Example |
|----------|-------------|---------|---------|
| `PERSIAN_DATE_FORMATTED` | Formatted Persian date | "day month year" | `=PERSIAN_DATE_FORMATTED(TODAY())` |
| `PERSIAN_MONTH_NAME` | Persian month name | String | `=PERSIAN_MONTH_NAME(TODAY())` |
| `PERSIAN_MONTH_NAME_BY_NUMBER` | Month name by number | String | `=PERSIAN_MONTH_NAME_BY_NUMBER(5)` |
| `PERSIAN_WEEKDAY_NAME` | Persian weekday name | String | `=PERSIAN_WEEKDAY_NAME(TODAY())` |

#### Utility Functions

| Function | Description | Returns | Example |
|----------|-------------|---------|---------|
| `TODAY_PERSIAN` | Current Persian date | String | `=TODAY_PERSIAN()` |
| `TODAY_PERSIAN_FORMATTED` | Current formatted Persian date | String | `=TODAY_PERSIAN_FORMATTED()` |
| `IS_PERSIAN_LEAP_YEAR` | Check leap year | Boolean | `=IS_PERSIAN_LEAP_YEAR(1403)` |
| `PERSIAN_DAYS_IN_MONTH` | Days in Persian month | Long | `=PERSIAN_DAYS_IN_MONTH(1403, 12)` |
| `IS_VALID_PERSIAN_DATE` | Validate Persian date | Boolean | `=IS_VALID_PERSIAN_DATE("1403/05/08")` |

#### Calculation Functions

| Function | Description | Syntax | Example |
|----------|-------------|---------|---------|
| `ADD_DAYS_TO_PERSIAN` | Add days to Persian date | `ADD_DAYS_TO_PERSIAN(dateString, days)` | `=ADD_DAYS_TO_PERSIAN("1403/05/08", 10)` |
| `PERSIAN_DATE_DIFF` | Calculate days difference | `PERSIAN_DATE_DIFF(startDate, endDate)` | `=PERSIAN_DATE_DIFF("1403/05/01", "1403/05/08")` |

## Data Types

### PersianDate Type
```vba
Public Type PersianDate
    Year As Long    ' Persian year
    Month As Long   ' Persian month (1-12)
    Day As Long     ' Persian day (1-31)
End Type
```

## Constants

```vba
Public Const PERSIAN_ERA As Long = 1
Public Const MIN_PERSIAN_YEAR As Long = 1
Public Const MAX_PERSIAN_YEAR As Long = 9378
```

## Error Handling

All functions raise Runtime Error 5 (Invalid Procedure Call) for invalid input.

### Common Errors:
- **"Invalid Persian date"**: Invalid Persian date values
- **"Invalid date format"**: Wrong date string format (must be YYYY/MM/DD)

## Persian Month Names

| Number | Persian Name | English Name |
|--------|-------------|-------------|
| 1 | فروردین | Farvardin |
| 2 | اردیبهشت | Ordibehesht |
| 3 | خرداد | Khordad |
| 4 | تیر | Tir |
| 5 | مرداد | Mordad |
| 6 | شهریور | Shahrivar |
| 7 | مهر | Mehr |
| 8 | آبان | Aban |
| 9 | آذر | Azar |
| 10 | دی | Dey |
| 11 | بهمن | Bahman |
| 12 | اسفند | Esfand |

## Persian Weekday Names

| Persian | English |
|---------|---------|
| شنبه | Saturday |
| یکشنبه | Sunday |
| دوشنبه | Monday |
| سه‌شنبه | Tuesday |
| چهارشنبه | Wednesday |
| پنج‌شنبه | Thursday |
| جمعه | Friday |