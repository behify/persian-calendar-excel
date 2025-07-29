# تقویم شمسی برای مایکروسافت اکسل

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

یک پکیج VBA کامل برای تبدیل تاریخ شمسی به میلادی و بالعکس در Microsoft Excel، بر اساس الگوریتم .NET PersianCalendar.

فارسی | **[English](README.md)**

</div>

## ویژگی‌ها

- تبدیل دوطرفه: شمسی ↔ میلادی
- تشخیص سال کبیسه بر اساس چرخه 33 ساله
- نام‌های فارسی ماه‌ها و روزهای هفته
- محاسبات تاریخ: اضافه کردن روز، محاسبه تفاوت
- اعتبارسنجی و بررسی تاریخ‌ها
- شامل مجموعه کامل تست‌ها

## نصب

### مرحله اول: دانلود فایل‌ها
```bash
git clone https://github.com/behify/persian-calendar-excel.git
```

### مرحله دوم: نصب در اکسل
1. Microsoft Excel را باز کنید
2. `Alt+F11` را برای باز کردن VBA Editor فشار دهید
3. `Insert > Module` را 5 بار تکرار کنید تا 5 ماژول ایجاد شود
4. محتوای هر فایل `.bas` را در ماژول مربوطه کپی کنید:
   - `Module1` ← `PersianCalendarConstants.bas`
   - `Module2` ← `PersianToGregorianConverter.bas`
   - `Module3` ← `GregorianToPersianConverter.bas`
   - `Module4` ← `PersianCalendarHelpers.bas`
   - `Module5` ← `PersianCalendarTests.bas`
5. فایل را با فرمت `.xlsm` ذخیره کنید

### مرحله سوم: حل هشدار امنیتی

هنگام باز کردن فایل Excel حاوی ماکرو، ممکن است هشدار امنیتی مشاهده کنید. این برای فایل‌های Excel حاوی کد VBA طبیعی است.

**برای فعال کردن توابع تقویم شمسی:**

**روش سریع (پیشنهادی):**
1. روی "Enable Content" کلیک کنید وقتی هشدار امنیتی ظاهر می‌شود
2. توابع فوراً در دسترس خواهند بود

**روش دائمی:**
1. به `File > Options > Trust Center > Trust Center Settings` بروید
2. `Macro Settings` را انتخاب کنید
3. "Disable all macros with notification" را انتخاب کنید
4. Excel را مجدداً راه‌اندازی کرده و فایل را دوباره باز کنید
5. روی "Enable Content" کلیک کنید

**برای مکان مورد اعتماد (راحت‌ترین روش):**
1. به `File > Options > Trust Center > Trust Center Settings` بروید
2. `Trusted Locations` را انتخاب کنید
3. روی "Add new location" کلیک کنید
4. فولدر پروژه را جستجو و انتخاب کنید
5. "Subfolders of this location are also trusted" را تیک بزنید
6. فایل‌های این مکان به طور خودکار ماکروها را اجرا خواهند کرد

**مهم:** فقط ماکروها را از منابع مورد اعتماد فعال کنید. این پکیج متن‌باز و امن است.

### مرحله چهارم: تست نصب
```vba
' در VBA Editor اجرا کنید
TestPersianCalendarFunctions()
```

## استفاده

### توابع اصلی

#### تبدیل شمسی به میلادی
```excel
=PERSIAN_TO_GREGORIAN("1403/05/08")
=PERSIAN_DATE_TO_GREGORIAN(1403, 5, 8)
```

#### تبدیل میلادی به شمسی
```excel
=GREGORIAN_TO_PERSIAN(TODAY())
=PERSIAN_YEAR(A1)
=PERSIAN_MONTH(A1)
=PERSIAN_DAY(A1)
```

#### فرمت‌دهی و نمایش
```excel
=PERSIAN_DATE_FORMATTED(TODAY())          // "8 مرداد 1403"
=PERSIAN_MONTH_NAME(TODAY())              // "مرداد"
=PERSIAN_WEEKDAY_NAME(TODAY())            // "پنج‌شنبه"
```

#### محاسبات
```excel
=TODAY_PERSIAN()                          // تاریخ امروز شمسی
=ADD_DAYS_TO_PERSIAN("1403/05/08", 10)    // اضافه کردن 10 روز
=PERSIAN_DATE_DIFF("1403/05/01", "1403/05/08")  // تفاوت روزها
```

#### کمکی
```excel
=IS_PERSIAN_LEAP_YEAR(1403)              // بررسی کبیسه
=PERSIAN_DAYS_IN_MONTH(1403, 12)         // تعداد روزهای ماه
=IS_VALID_PERSIAN_DATE("1403/05/08")     // اعتبارسنجی
```

## فهرست کامل توابع

| تابع | توضیح | مثال |
|------|-------|-------|
| `PERSIAN_TO_GREGORIAN` | تبدیل شمسی به میلادی | `=PERSIAN_TO_GREGORIAN("1403/05/08")` |
| `GREGORIAN_TO_PERSIAN` | تبدیل میلادی به شمسی | `=GREGORIAN_TO_PERSIAN(TODAY())` |
| `PERSIAN_YEAR` | استخراج سال شمسی | `=PERSIAN_YEAR(A1)` |
| `PERSIAN_MONTH` | استخراج ماه شمسی | `=PERSIAN_MONTH(A1)` |
| `PERSIAN_DAY` | استخراج روز شمسی | `=PERSIAN_DAY(A1)` |
| `PERSIAN_DATE_FORMATTED` | تاریخ فرمت شده فارسی | `=PERSIAN_DATE_FORMATTED(A1)` |
| `PERSIAN_MONTH_NAME` | نام ماه فارسی | `=PERSIAN_MONTH_NAME(A1)` |
| `PERSIAN_WEEKDAY_NAME` | نام روز هفته فارسی | `=PERSIAN_WEEKDAY_NAME(A1)` |
| `TODAY_PERSIAN` | تاریخ امروز شمسی | `=TODAY_PERSIAN()` |
| `IS_PERSIAN_LEAP_YEAR` | بررسی سال کبیسه | `=IS_PERSIAN_LEAP_YEAR(1403)` |
| `ADD_DAYS_TO_PERSIAN` | اضافه کردن روز | `=ADD_DAYS_TO_PERSIAN("1403/05/08", 10)` |
| `PERSIAN_DATE_DIFF` | محاسبه تفاوت روزها | `=PERSIAN_DATE_DIFF("1403/05/01", "1403/05/08")` |

## نمونه کد

برای مشاهده نمونه‌های کاربردی:

```vba
' در VBA Editor اجرا کنید
GenerateSampleData()
```

این دستور یک شیت نمونه با مثال‌های متنوع ایجاد می‌کند.

## ساختار پروژه

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

## الگوریتم

این پکیج بر اساس الگوریتم رسمی .NET Framework PersianCalendar پیاده‌سازی شده است:
- چرخه 33 ساله برای تشخیص سال کبیسه
- تاریخ مرجع: 21 مارس 622 میلادی (1 فروردین 1 شمسی)
- دقت بالا در محاسبات

## مشارکت

مشارکت‌ها خوشامد است! لطفاً راهنمای [CONTRIBUTING.md](CONTRIBUTING.md) را مطالعه کنید.

1. مخزن را Fork کنید
2. شاخه ویژگی ایجاد کنید (`git checkout -b feature/amazing-feature`)
3. تغییرات را Commit کنید (`git commit -m 'Add amazing feature'`)
4. به شاخه Push کنید (`git push origin feature/amazing-feature`)
5. Pull Request ایجاد کنید

## مجوز

این پروژه تحت مجوز MIT منتشر شده است. فایل [LICENSE](LICENSE) را مطالعه کنید.

## پشتیبانی

اگر مشکلی داشتید:
- ابتدا [راهنمای عیب‌یابی](docs/TROUBLESHOOTING.md) را بررسی کنید
- [Issues](https://github.com/behify/persian-calendar-excel/issues) را مطالعه کنید
- در [Discussions](https://github.com/behify/persian-calendar-excel/discussions) شرکت کنید

## تشکر

این پروژه بر اساس الگوریتم [.NET PersianCalendar](https://github.com/dotnet/runtime/blob/main/src/libraries/System.Private.CoreLib/src/System/Globalization/PersianCalendar.cs) پیاده‌سازی شده است.

## نویسنده

ایجاد شده توسط [Behify](https://github.com/behify)