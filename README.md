# Bale NEW_FIRE Automation — User Account Mode

این نسخه برای حالتی است که کانال خبرگزاری مال شما نیست و اکانت شخصی شما عضو آن کانال است.

برنامه با اکانت واقعی بله شما وارد می‌شود، تاریخ شمسی امروز را می‌سازد و در history کانال دنبال فایل دقیق همان روز می‌گردد.

مثال:

```text
1405/07/17 -> NEW_FIRE_14050717_IPS.xlsx
1405/07/18 -> NEW_FIRE_14050718_IPS.xlsx
```

پس از پیدا کردن فایل:

1. Excel را دانلود می‌کند.
2. ۱۰ IPv4 معتبر اول را پیدا می‌کند.
3. Country همان ردیف را برمی‌دارد.
4. همه را در یک پیام برای مقصد می‌فرستد.
5. نتیجه را در `processed_files.json` ثبت می‌کند تا دوباره ارسال نشود.

> این پروژه از `bale-sdk` استفاده می‌کند که یک کلاینت غیررسمی برای API وب بله است. ممکن است با تغییرات بله نیاز به اصلاح پیدا کند.

## 1) نصب

Python 3.10+ پیشنهاد می‌شود.

### Windows PowerShell

```powershell
git clone https://github.com/Iman-gholami/balebot-.git
cd balebot-
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

### Linux

```bash
git clone https://github.com/Iman-gholami/balebot-.git
cd balebot-
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

## 2) ورود به اکانت بله

```bash
python login.py
```

شماره موبایل اکانتی را وارد کنید که عضو کانال خبرگزاری است.

کد OTP برای همان اکانت ارسال می‌شود. بعد از وارد کردن کد، `BALE_TOKEN` داخل فایل محلی `.env` ذخیره می‌شود.

فایل `.env` در `.gitignore` است و نباید روی GitHub قرار بگیرد.

## 3) پیدا کردن شناسه کانال و مقصد

```bash
python list_dialogs.py
```

خروجی نمونه:

```text
channel:123456789    News Channel
user:987654321       Ali
```

سپس در `.env` تنظیم کنید:

```text
BALE_SOURCE=channel:123456789
BALE_DESTINATION=user:987654321
```

اگر کانال یا کاربر username عمومی دارد، می‌توانید مستقیماً استفاده کنید:

```text
BALE_SOURCE=@news_username
BALE_DESTINATION=@destination_username
```

## 4) اجرای تست

```bash
python bot.py
```

ابتدا باید لاگی شبیه این ببینید:

```text
Logged in as ...
Source: ...
Today's target: NEW_FIRE_14050717_IPS.xlsx
```

اگر فایل امروز قبلاً در history کانال باشد، لازم نیست دوباره در کانال ارسال شود. برنامه تا `HISTORY_LIMIT` پیام آخر را جستجو می‌کند.

اگر هنوز فایل منتشر نشده باشد:

```text
Today's file not found yet: NEW_FIRE_...
```

برنامه هر `CHECK_INTERVAL` ثانیه دوباره بررسی می‌کند.

## Excel

عنوان‌های رایج IP:

```text
IP
IP Address
ip_address
ipaddr
sourceip
destinationip
```

عنوان‌های رایج Country:

```text
Country
Country Name
Location
Country Code
```

اگر header قابل تشخیص نباشد:

- ستون A = IP
- ستون B = Country

## خروجی

```text
NEW_FIRE_14050717_IPS

1. 185.10.20.30 - Iran
2. 91.100.20.5 - Germany
...
10. 8.8.8.8 - United States
```

## تنظیمات مهم

```text
BALE_TIMEZONE=Asia/Tehran
CHECK_INTERVAL=60
HISTORY_LIMIT=100
```

اگر کانال در یک روز بیشتر از ۱۰۰ پیام دارد، `HISTORY_LIMIT` را بیشتر کنید.


## اجرای خودکار روزانه ساعت ۱۲

برنامه اکنون بعد از ارسال موفق خروجی همان روز بسته می‌شود. اگر فایل هنوز منتشر نشده باشد، هر `CHECK_INTERVAL` ثانیه دوباره بررسی می‌کند و حداکثر تا `MAX_WAIT_MINUTES` دقیقه منتظر می‌ماند.

پیشنهاد:

```text
CHECK_INTERVAL=60
MAX_WAIT_MINUTES=360
```

برای اجرای روزانه با cron، مسیر Python داخل virtualenv را مستقیماً استفاده کنید. نمونه:

```cron
0 12 * * * cd /home/thomas/balebot- && /home/thomas/balebot-/.venv/bin/python bot.py >> /home/thomas/balebot-/bot.log 2>&1
```

زمان cron بر اساس timezone سیستم VPS است. برای اجرای ساعت ۱۲ ایران، timezone سیستم/cron باید Asia/Tehran باشد یا زمان معادل آن تنظیم شود.
