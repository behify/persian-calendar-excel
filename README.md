# Persian Calendar for Microsoft Excel

<div align="center">
  <img src="https://raw.githubusercontent.com/behify/persian-calendar-excel/main/assets/logo.png" alt="Persian Calendar Excel Logo" width="200">
</div>

<div align="center">

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/behify/persian-calendar-excel?style=flat-square&color=brightgreen)](https://github.com/behify/persian-calendar-excel/releases/latest)
[![GitHub Workflow Status](https://img.shields.io/github/actions/workflow/status/behify/persian-calendar-excel/test.yml?branch=main&style=flat-square)](https://github.com/behify/persian-calendar-excel/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)
[![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346?style=flat-square&logo=microsoft-excel&logoColor=white)](https://docs.microsoft.com/en-us/office/vba/api/overview/)
[![Persian Calendar](https://img.shields.io/badge/Calendar-Persian%20%7C%20Jalali-blue.svg?style=flat-square)](https://en.wikipedia.org/wiki/Solar_Hijri_calendar)
[![GitHub stars](https://img.shields.io/github/stars/behify/persian-calendar-excel?style=flat-square&color=gold)](https://github.com/behify/persian-calendar-excel/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/behify/persian-calendar-excel?style=flat-square&color=blue)](https://github.com/behify/persian-calendar-excel/network)
[![GitHub issues](https://img.shields.io/github/issues/behify/persian-calendar-excel?style=flat-square&color=red)](https://github.com/behify/persian-calendar-excel/issues)

A comprehensive VBA package for Persian (Jalali) to Gregorian calendar conversion in Microsoft Excel, based on the .NET PersianCalendar algorithm.

**[فارسی](README.fa.md)** | English

</div>

## Features

- Bidirectional conversion: Persian ↔ Gregorian
- Leap year detection based on 33-year cycle
- Persian month and weekday names
- Date calculations: add days, calculate differences
- Date validation and verification
- Comprehensive test suite included

## Installation

### Step 1: Download Files
```bash
git clone https://github.com/behify/persian-calendar-excel.git
```

### Step 2: Install in Excel
1. Open Microsoft Excel
2. Press `Alt+F11` to open VBA Editor
3. Repeat `Insert > Module` 5 times to create 5 modules
4. Copy each `.bas` file content to corresponding module:
   - `Module1` ← `PersianCalendarConstants.bas`
   - `Module2` ← `PersianToGregorianConverter.bas`
   - `Module3` ← `GregorianToPersianConverter.bas`
   - `Module4` ← `PersianCalendarHelpers.bas`
   - `Module5` ← `PersianCalendarTests.bas`
5. Save file as `.xlsm` format

### Step 3: Security Warning Resolution

When opening the Excel file with macros, you may see a security warning. This is normal for Excel files containing VBA code.

**To enable the Persian Calendar functions:**

**Quick Method (Recommended):**
1. Click "Enable Content" when the security warning appears
2. The functions will be immediately available

**Permanent Method:**
1. Go to `File > Options > Trust Center > Trust Center Settings`
2. Select `Macro Settings`
3. Choose "Disable all macros with notification"
4. Restart Excel and reopen the file
5. Click "Enable Content" when prompted

**For Trusted Location (Most Convenient):**
1. Go to `File > Options > Trust Center > Trust Center Settings`
2. Select `Trusted Locations`
3. Click "Add new location"
4. Browse and select your project folder
5. Check "Subfolders of this location are also trusted"
6. Files in this location will automatically run macros

**Important:** Only enable macros from trusted sources. This package is open source and safe to use.

### Step 4: Test Installation
```vba
' Run in VBA Editor
TestPersianCalendarFunctions()
```

## Usage

### Core Functions

#### Persian to Gregorian Conversion
```excel
=PERSIAN_TO_GREGORIAN("1403/05/08")
=PERSIAN_DATE_TO_GREGORIAN(1403, 5, 8)
```

#### Gregorian to Persian Conversion
```excel
=GREGORIAN_TO_PERSIAN(TODAY())
=PERSIAN_YEAR(A1)
=PERSIAN_MONTH(A1)
=PERSIAN_DAY(A1)
```

#### Formatting and Display
```excel
=PERSIAN_DATE_FORMATTED(TODAY())          // "8 Mordad 1403"
=PERSIAN_MONTH_NAME(TODAY())              // "Mordad"
=PERSIAN_WEEKDAY_NAME(TODAY())            // "Panjshanbeh"
```

#### Calculations
```excel
=TODAY_PERSIAN()                          // Current Persian date
=ADD_DAYS_TO_PERSIAN("1403/05/08", 10)    // Add 10 days
=PERSIAN_DATE_DIFF("1403/05/01", "1403/05/08")  // Days difference
```

#### Utilities
```excel
=IS_PERSIAN_LEAP_YEAR(1403)              // Check leap year
=PERSIAN_DAYS_IN_MONTH(1403, 12)         // Days in month
=IS_VALID_PERSIAN_DATE("1403/05/08")     // Date validation
```

## Complete Function List

| Function | Description | Example |
|----------|-------------|---------|
| `PERSIAN_TO_GREGORIAN` | Convert Persian to Gregorian | `=PERSIAN_TO_GREGORIAN("1403/05/08")` |
| `GREGORIAN_TO_PERSIAN` | Convert Gregorian to Persian | `=GREGORIAN_TO_PERSIAN(TODAY())` |
| `PERSIAN_YEAR` | Extract Persian year | `=PERSIAN_YEAR(A1)` |
| `PERSIAN_MONTH` | Extract Persian month | `=PERSIAN_MONTH(A1)` |
| `PERSIAN_DAY` | Extract Persian day | `=PERSIAN_DAY(A1)` |
| `PERSIAN_DATE_FORMATTED` | Formatted Persian date | `=PERSIAN_DATE_FORMATTED(A1)` |
| `PERSIAN_MONTH_NAME` | Persian month name | `=PERSIAN_MONTH_NAME(A1)` |
| `PERSIAN_WEEKDAY_NAME` | Persian weekday name | `=PERSIAN_WEEKDAY_NAME(A1)` |
| `TODAY_PERSIAN` | Current Persian date | `=TODAY_PERSIAN()` |
| `IS_PERSIAN_LEAP_YEAR` | Check leap year | `=IS_PERSIAN_LEAP_YEAR(1403)` |
| `ADD_DAYS_TO_PERSIAN` | Add days to Persian date | `=ADD_DAYS_TO_PERSIAN("1403/05/08", 10)` |
| `PERSIAN_DATE_DIFF` | Calculate days difference | `=PERSIAN_DATE_DIFF("1403/05/01", "1403/05/08")` |

## Examples

To see practical examples:

```vba
' Run in VBA Editor
GenerateSampleData()
```

This creates a sample worksheet with various usage examples.

## Project Structure

```
persian-calendar-excel/
├── README.md
├── README.fa.md
├── LICENSE
├── assets/
│   └── logo.png
├── src/
│   ├── PersianCalendarConstants.bas
│   ├── PersianToGregorianConverter.bas
│   ├── GregorianToPersianConverter.bas
│   ├── PersianCalendarHelpers.bas
│   └── PersianCalendarTests.bas
├── examples/
│   └── Sample.xlsx
└── docs/
    ├── API.md
    └── TROUBLESHOOTING.md
```

## Algorithm

This package is implemented based on the official .NET Framework PersianCalendar algorithm:
- 33-year cycle for leap year detection
- Reference date: March 21, 622 AD (1 Farvardin 1 Persian)
- High precision calculations

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) file for details.

## Support

If you encounter any issues:
- Check [Troubleshooting Guide](docs/TROUBLESHOOTING.md) first
- Review [Issues](https://github.com/behify/persian-calendar-excel/issues)
- Join [Discussions](https://github.com/behify/persian-calendar-excel/discussions)

## Acknowledgments

This project is based on the [.NET PersianCalendar](https://github.com/dotnet/runtime/blob/main/src/libraries/System.Private.CoreLib/src/System/Globalization/PersianCalendar.cs) algorithm.

## Author

Created by [Behify](https://github.com/behify)