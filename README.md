# DC Consulting Accounting System v3.0

نظام محاسبي متكامل لـ Dewan Consulting مبني على Google Apps Script.

## الملفات

| الملف | الوصف |
|-------|-------|
| `01-Config.gs` | القائمة والإعدادات الأساسية |
| `02-Databases.gs` | قواعد البيانات (Settings, Holidays, Items) |
| `03-Parties.gs` | العملاء والموردين والموظفين |
| `04-CashBank.gs` | إدارة الخزائن والحسابات البنكية |
| `05-Transactions.gs` | الحركات المالية |
| `06-Invoicing.gs` | نظام الفواتير |
| `07-Email.gs` | إرسال البريد الإلكتروني |
| `08-Reports.gs` | التقارير وكشوف الحسابات |
| `09-Dashboard.gs` | لوحة التحكم |
| `10-Advances.gs` | نظام العهد المؤقتة |

## النشر التلقائي

هذا المشروع يستخدم GitHub Actions للنشر التلقائي إلى Google Apps Script.

### الإعداد

1. احصل على `SCRIPT_ID` من Google Apps Script
2. أضف الـ Secrets التالية في GitHub:
   - `CLASP_TOKEN`: من `~/.clasprc.json` بعد `clasp login`
   - `SCRIPT_ID`: معرف مشروع Apps Script

### الاستخدام

كل push إلى `main` سيتم نشره تلقائياً إلى Google Apps Script.

## الميزات

- ✅ إدارة العملاء والموردين
- ✅ تتبع الحركات المالية
- ✅ إنشاء الفواتير وإرسالها
- ✅ إدارة الخزائن والبنوك
- ✅ نظام العهد المؤقتة
- ✅ التقارير وكشوف الحسابات
- ✅ لوحة تحكم شاملة
