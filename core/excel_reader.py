import re
from pathlib import Path
import openpyxl

OPERATIONAL_SHEETS = ["وحید جهاندیده", "ابراهیم احدی", "جلال غضنفری", "حسین داودی", "رامین صدری", "سعید احمدی", "مجتبی طاهریان", "حسین یکتا"]
MONTHS = {"فروردین":1,"اردیبهشت":2,"خرداد":3,"تیر":4,"مرداد":5,"شهریور":6,"مهر":7,"آبان":8,"آذر":9,"دی":10,"بهمن":11,"اسفند":12}
TRANS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

def norm(value):
    return str(value).translate(TRANS).replace("\u200c", " ").strip() if value is not None else ""

def read_workbook(path):
    p = Path(path)
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True, keep_vba=p.suffix.lower()==".xlsm")
    sheets = [s for s in OPERATIONAL_SHEETS if s in wb.sheetnames]
    if not sheets:
        raise ValueError("هیچ‌یک از ۸ شیت عملیاتی مورد انتظار پیدا نشد.")
    employees, warnings = [], []
    month_name, month_number, year = "", None, None
    for sn in sheets:
        ws = wb[sn]
        if not month_name:
            title = norm(ws["P1"].value)
            found = next((m for m in MONTHS if m in title), None)
            month_name, month_number = found or title or "نامشخص", MONTHS.get(found)
            ym = re.search(r"(?:13|14)\d{2}", norm(ws["X1"].value))
            year = int(ym.group()) if ym else None
        for r in range(5, ws.max_row + 1):
            first, last, nid = norm(ws.cell(r,2).value), norm(ws.cell(r,3).value), norm(ws.cell(r,4).value)
            if not (first and last and nid):
                continue
            nid = re.sub(r"\D", "", nid)
            if nid and len(nid) < 10:
                nid = nid.zfill(10)
            if len(nid) != 10:
                warnings.append(f"{sn}، ردیف {r}: کد ملی غیرمعمول ({nid})")
            days = {d:norm(ws.cell(r,15+d).value) for d in range(1,32)}
            employees.append({"sheet":sn,"row":r,"first":first,"last":last,"nid":nid,"shift":norm(ws.cell(r,10).value),"days":days})
    wb.close()
    if month_number in (7,8,9,10,11): max_days=30
    elif month_number == 12: max_days=29
    elif month_number: max_days=31
    else: max_days=31
    for e in employees:
        e["days"] = {d:v for d,v in e["days"].items() if d <= max_days}
    if not month_number or not year:
        warnings.append("ماه یا سال از P1 و X1 کامل تشخیص داده نشد؛ قبل از استفاده بررسی کنید.")
    if month_number and year:
        start_date = f"{year:04d}/{month_number:02d}/01"
        next_month, next_year = (month_number + 1, year) if month_number < 12 else (1, year + 1)
        end_date = f"{next_year:04d}/{next_month:02d}/01"
    else:
        start_date = end_date = ""
    return {"month":month_name,"month_number":month_number,"year":year,"start_date":start_date,"end_date":end_date,"employees":employees,"sheets":sheets,"warnings":warnings}
