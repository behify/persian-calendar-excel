# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2025-07-29

### Added
- **Core Conversion Functions**
  - `PersianToGregorian()` - Convert Persian date to Gregorian
  - `GregorianToPersian()` - Convert Gregorian date to Persian
  - Based on .NET PersianCalendar algorithm

- **Excel Worksheet Functions**
  - `PERSIAN_TO_GREGORIAN()` - Convert Persian string to Gregorian
  - `GREGORIAN_TO_PERSIAN()` - Convert Gregorian to Persian string
  - `PERSIAN_YEAR()`, `PERSIAN_MONTH()`, `PERSIAN_DAY()` - Extract date components
  - `PERSIAN_DATE_FORMATTED()` - Formatted Persian date with month name
  - `PERSIAN_MONTH_NAME()` - Persian month name
  - `PERSIAN_WEEKDAY_NAME()` - Persian weekday name

- **Utility Functions**
  - `TODAY_PERSIAN()` - Current Persian date
  - `IS_PERSIAN_LEAP_YEAR()` - Check leap year
  - `PERSIAN_DAYS_IN_MONTH()` - Days in Persian month
  - `ADD_DAYS_TO_PERSIAN()` - Add days to Persian date
  - `PERSIAN_DATE_DIFF()` - Calculate days difference
  - `IS_VALID_PERSIAN_DATE()` - Date validation

- **Calendar Features**
  - 33-year cycle for leap year detection
  - Support for Persian years 1 to 9378
  - Persian month and weekday names
  - Accurate calculations based on March 21, 622 AD reference

- **Testing & Quality**
  - Comprehensive automatic test suite
  - `TestPersianCalendarFunctions()` - Run all tests
  - `GenerateSampleData()` - Generate sample data for testing
  - Input and output validation

- **Documentation**
  - Complete installation and usage guide
  - API documentation with practical examples
  - Contributing guidelines
  - Various usage examples

### Technical Details
- **Modules**: 5 organized and independent VBA modules
- **Compatibility**: Microsoft Excel 2010 and later
- **Language**: English code, Persian names and outputs
- **Algorithm**: Based on .NET Framework PersianCalendar
- **Testing**: 5 main test suites with 10+ detailed tests

### Performance
- Fast and optimized calculations
- No external dependencies
- Minimal memory usage
- Compatible with large Excel files

---

## Template for Future Releases

### [X.Y.Z] - YYYY-MM-DD

### Added
- New features

### Changed 
- Changes in existing functionality

### Deprecated
- Soon-to-be removed features

### Removed
- Now removed features

### Fixed
- Bug fixes

### Security
- Vulnerability fixes