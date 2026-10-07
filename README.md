# Bale NEW_FIRE automation

این ربات کانال بله را برای فایل‌های روزانه با این الگو بررسی می‌کند:

```text
NEW_FIRE_14050715_IPS.xlsx
NEW_FIRE_14050716_IPS.xlsx
NEW_FIRE_14050717_IPS.xlsx
```

به‌محض رسیدن فایل مطابق الگو:

1. فایل Excel را از بله دانلود می‌کند.
2. ۱۰ IP معتبر اول را پیدا می‌کند.
3. نام/کد کشور همان ردیف را هم برمی‌دارد.
4. هر ۱۰ مورد را در یک پیام به اکانت مقصد در بله می‌فرستد.
5. فایل پردازش‌شده را ثبت می‌کند تا دوباره ارسال نشود.

## پیش‌نیاز

- Python 3.10 یا جدیدتر
- یک Bot در بله
- اضافه‌کردن Bot به کانال مبدا
- اکانت مقصد باید یک بار Bot را باز کند و `/start` بزند

## نصب

### Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows

```powershell
py -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## تنظیمات

فایل نمونه را کپی کن:

### Linux

```bash
cp .env.example .env
```

### Windows PowerShell

```powershell
Copy-Item .env.example .env
```

بعد فایل `.env` را پر کن:

```text
BALE_BOT_TOKEN=توکن_ربات
BALE_SOURCE_CHANNEL_ID=شناسه_کانال
BALE_DESTINATION_CHAT_ID=شناسه_اکانت_مقصد
```

> توکن Bot را داخل GitHub commit نکن. فایل `.env` در `.gitignore` قرار دارد.

## پیدا کردن Chat ID مقصد

برنامه را با توکن Bot اجرا کن و در اکانت مقصد به Bot این دستور را بفرست:

```text
/id
```

Bot جواب می‌دهد:

```text
Chat ID: 123456789
```

این عدد را در `BALE_DESTINATION_CHAT_ID` قرار بده.

## پیدا کردن Channel ID

ابتدا می‌توانی `BALE_SOURCE_CHANNEL_ID` را در `.env` خالی بگذاری و Bot را به کانال اضافه کنی.

برای گرفتن شناسه کانال، از خروجی `getUpdates` بله یا ابزار مدیریت Bot استفاده کن و مقدار `chat.id` مربوط به `channel_post` را در `BALE_SOURCE_CHANNEL_ID` قرار بده.

بعد از تنظیم این مقدار، Bot فایل‌های کانال‌های دیگر را نادیده می‌گیرد.

## اجرا

```bash
python bot.py
```

وقتی اجرا شود، منتظر فایل‌های جدید می‌ماند.

## فرمت Excel

Bot سعی می‌کند ستون‌ها را از روی عنوان پیدا کند.

نام‌های شناخته‌شده برای IP شامل:

```text
IP
IP Address
ip_address
ipaddr
sourceip
destinationip
```

نام‌های شناخته‌شده برای کشور شامل:

```text
Country
Country Name
Location
Country Code
```

اگر عنوان ستون IP قابل تشخیص نباشد، Bot فرض می‌کند:

- ستون A = IP
- ستون B = Country

فقط IPv4 معتبر شمرده می‌شود و وقتی حداقل ۱۰ IP پیدا شود، پیام ارسال می‌شود.

## نمونه پیام خروجی

```text
NEW_FIRE_14050715_IPS

1. 185.10.20.30 - Iran
2. 91.100.20.5 - Germany
3. 45.12.55.20 - France
4. ...
5. ...
6. ...
7. ...
8. ...
9. ...
10. 8.8.8.8 - United States
```

## جلوگیری از ارسال دوباره

بعد از ارسال موفق، شناسه فایل در فایل محلی زیر ذخیره می‌شود:

```text
processed_files.json
```

بنابراین اگر همان فایل دوباره در کانال فرستاده شود، مجدداً برای مقصد ارسال نمی‌شود.

## نکته

Parser فعلی با `openpyxl` است و برای فایل‌های `.xlsx` طراحی شده است. اگر فایل واقعی شما فرمت قدیمی `.xls` باشد باید پشتیبانی جداگانه اضافه شود.
