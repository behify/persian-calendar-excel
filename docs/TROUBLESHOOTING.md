# Troubleshooting Guide

<div align="center">
  <img src="https://raw.githubusercontent.com/behify/persian-calendar-excel/main/assets/logo.png" alt="Persian Calendar Excel Logo" width="150">
</div>

## Common Issues and Solutions

### 1. Security Warning: "Microsoft has blocked macros from running"

**Problem:** Excel shows a security warning and blocks macros when opening files containing VBA code.

**Solution Options:**

#### Option A: Enable Content (Quick Fix)
1. Click the **"Enable Content"** button in the security warning bar
2. Functions will be immediately available
3. You'll need to do this each time you open the file

#### Option B: Change Macro Security Settings (Permanent)
1. Go to `File > Options`
2. Select `Trust Center` from the left menu
3. Click `Trust Center Settings`
4. Select `Macro Settings`
5. Choose **"Disable all macros with notification"**
6. Click `OK` twice
7. Restart Excel and reopen the file
8. Click `Enable Content` when prompted

#### Option C: Add to Trusted Locations (Most Convenient)
1. Go to `File > Options > Trust Center > Trust Center Settings`
2. Select `Trusted Locations`
3. Click **"Add new location"**
4. Browse to your project folder
5. Check **"Subfolders of this location are also trusted"**
6. Click `OK`
7. Files in this location will automatically run macros

### 2. Functions Showing #NAME? Error

**Problem:** Excel functions show `#NAME?` error instead of results.

**Possible Causes & Solutions:**

#### Cause A: Macros are disabled
- **Solution:** Follow security warning resolution above

#### Cause B: VBA modules not properly installed
- **Solution:** 
  1. Press `Alt+F11` to open VBA Editor
  2. Verify all 5 modules are present with correct code
  3. Check for compilation errors (`Debug > Compile VBAProject`)

#### Cause C: Function name misspelled
- **Solution:** Check function spelling against documentation
- Common mistakes:
  - `PERSIAN_TO_GREGORIAN` (correct)
  - `PERSIANTOGREGORIAN` (incorrect)

### 3. Functions Return Wrong Dates

**Problem:** Date conversion functions return incorrect results.

**Troubleshooting Steps:**

#### Check Input Format
- Persian dates must be in format: `"YYYY/MM/DD"`
- Example: `"1403/05/08"` (correct)
- Example: `"1403-5-8"` (may cause issues)

#### Verify Date Validity
```excel
=IS_VALID_PERSIAN_DATE("1403/05/08")  ' Should return TRUE
```

#### Test with Known Dates
```excel
=PERSIAN_TO_GREGORIAN("1400/01/01")   ' Should return 2021/03/21
=GREGORIAN_TO_PERSIAN(DATE(2021,3,21)) ' Should return "1400/01/01"
```

### 4. Persian Text Displays as Question Marks

**Problem:** Persian month names and weekdays show as `????` or `???`.

**Solutions:**

#### Solution A: Font Settings
1. Select cells with Persian text
2. Right-click > `Format Cells`
3. Go to `Font` tab
4. Change font to **Tahoma** or **Arial Unicode MS**
5. Click `OK`

#### Solution B: Windows Language Settings
1. Open Windows Settings
2. Go to `Time & Language > Language`
3. Add Persian (Farsi) as a language
4. Restart Excel

#### Solution C: Cell Alignment
1. Select Persian text cells
2. Right-click > `Format Cells`
3. Go to `Alignment` tab
4. Set `Text direction` to **Right-to-left**

### 5. Performance Issues with Large Datasets

**Problem:** Functions run slowly with many calculations.

**Solutions:**

#### Solution A: Disable Automatic Calculation
1. Go to `Formulas > Calculation Options`
2. Select **Manual**
3. Press `F9` to recalculate when needed

#### Solution B: Use Array Formulas Sparingly
- Instead of applying function to each cell individually
- Consider copying results and pasting as values

### 6. Functions Not Available in Formula Builder

**Problem:** Persian Calendar functions don't appear in Excel's function list.

**Explanation:** Custom VBA functions don't appear in Excel's built-in function wizard, but they work when typed directly.

**Workaround:**
- Type function names directly in cells
- Use IntelliSense when available
- Refer to documentation for function signatures

### 7. Error: "Invalid Persian Date"

**Problem:** Functions raise "Invalid Persian date" error.

**Common Causes:**
- Day 31 in months 7-11 (only have 30 days)
- Day 30 in Esfand of non-leap years (only has 29 days)
- Month numbers outside 1-12 range
- Day numbers outside valid range

**Example Fixes:**
```excel
' Wrong:
=PERSIAN_TO_GREGORIAN("1403/07/31")  ' Mehr only has 30 days

' Correct:
=PERSIAN_TO_GREGORIAN("1403/07/30")
```

### 8. Excel Crashes When Using Functions

**Problem:** Excel becomes unresponsive or crashes.

**Troubleshooting:**

#### Check for Circular References
- Ensure functions don't reference cells that depend on themselves

#### Reduce Calculation Load
- Limit number of simultaneous calculations
- Use manual calculation mode for large datasets

#### Update Excel
- Ensure you're using a supported Excel version (2010 or later)

## Getting Help

If none of these solutions work:

1. **Check Issues:** Visit [GitHub Issues](https://github.com/behify/persian-calendar-excel/issues)
2. **Search Discussions:** Check [GitHub Discussions](https://github.com/behify/persian-calendar-excel/discussions)
3. **Create New Issue:** If problem persists, create a detailed issue report

## System Requirements

- **Excel Version:** Microsoft Excel 2010 or later
- **Operating System:** Windows 7 or later, macOS 10.12 or later
- **VBA Support:** Must be enabled (included in standard Excel installations)
- **Memory:** Minimum 4GB RAM recommended for large datasets

## Compatibility Notes

- **Excel Online:** VBA functions are not supported in Excel Online
- **Mobile Excel:** VBA functions are not supported on mobile devices
- **LibreOffice/OpenOffice:** Not compatible (Excel VBA specific)