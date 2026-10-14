# final_file_service.py - نسخه نهایی با قالب پویا، ویرایشگر WYSIWYG و پشتیبانی از دسترسی can_edit_template
import os
import sys
import json
import io
import re
import traceback
import logging
import subprocess
from datetime import datetime
from flask import Flask, request, send_file, jsonify, render_template_string
import pyodbc
import jdatetime
import pandas as pd
from docxtpl import DocxTemplate
from jinja2 import Environment, Template

# ==================== راه‌اندازی ====================
logging.basicConfig(level=logging.DEBUG, filename='service_errors.log', filemode='a',
                    format='%(asctime)s - %(levelname)s - %(message)s')
app = Flask(__name__)

# ==================== مسیرها و تنظیمات ====================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

FINAL_STORAGE = r"C:\Users\myazd\Downloads\Invoice-Bills"   # ریشه اصلی - می‌توانید تغییر دهید
os.makedirs(FINAL_STORAGE, exist_ok=True)

def get_db_connection():
    """اتصال به دیتابیس با خواندن db_config.json از مسیر مناسب"""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.abspath(".")
    local_config = os.path.join(base_dir, "db_config.json")
    if os.path.exists(local_config):
        config_path = local_config
    else:
        config_path = resource_path("db_config.json")
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={config['server']};DATABASE={config['database']};UID={config['username']};PWD={config['password']}"
    return pyodbc.connect(conn_str)

# ==================== توابع کمکی ====================
def to_persian_digits(text):
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    english_digits = "0123456789"
    trans_table = str.maketrans(english_digits, persian_digits)
    return text.translate(trans_table)

def persian_to_english_number_strict(value, field_name="مقدار"):
    if pd.isna(value):
        raise ValueError(f"{field_name}: مقدار None یا NaN است")
    value = str(value).strip()
    if value == "":
        raise ValueError(f"{field_name}: رشته خالی است")
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    english_digits = "0123456789"
    for p, e in zip(persian_digits, english_digits):
        value = value.replace(p, e)
    value = re.sub(r"[^\d\-]", "", value)
    if value == "" or value == "-":
        raise ValueError(f"{field_name}: عدد معتبر یافت نشد ('{value}')")
    return int(value)

def safe_str_strict(value, field_name="فیلد"):
    if pd.isna(value):
        raise ValueError(f"{field_name}: مقدار None یا NaN است")
    s = str(value).strip()
    if s == "":
        raise ValueError(f"{field_name}: رشته خالی است")
    return s

def safe_filename(text):
    text = str(text).strip()
    if not text:
        text = "بدون_نام"
    text = re.sub(r'[\\/*?:"<>|]', '-', text)
    text = text.replace('/', '-').replace('\\', '-')
    return text

def get_company_settings_from_db(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM company_settings WHERE id=1")
    row = cursor.fetchone()
    if row:
        cols = [desc[0] for desc in cursor.description]
        return dict(zip(cols, row))
    return {}

# ==================== توابع شماره‌گذاری اتمیک ====================
def get_next_letter_number(conn, payer_code, year):
    cursor = conn.cursor()
    cursor.execute("SELECT MIN(number) FROM available_letter_numbers WHERE payer_code = ? AND year = ?", (payer_code, year))
    row = cursor.fetchone()
    if row and row[0] is not None:
        return row[0]
    cursor.execute("SELECT last_number FROM letter_sequences WHERE payer_code = ? AND year = ?", (payer_code, year))
    row = cursor.fetchone()
    return row[0] + 1 if row else 1

def update_letter_sequence(conn, payer_code, year, last_number):
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM available_letter_numbers WHERE payer_code = ? AND year = ? AND number = ?", (payer_code, year, last_number))
    if cursor.fetchone():
        cursor.execute("DELETE FROM available_letter_numbers WHERE payer_code = ? AND year = ? AND number = ?", (payer_code, year, last_number))
    else:
        cursor.execute("""
            MERGE INTO letter_sequences AS target
            USING (VALUES (?, ?, ?)) AS source (payer_code, year, last_number)
            ON target.payer_code = source.payer_code AND target.year = source.year
            WHEN MATCHED THEN UPDATE SET last_number = source.last_number
            WHEN NOT MATCHED THEN INSERT (payer_code, year, last_number) VALUES (source.payer_code, source.year, source.last_number);
        """, (payer_code, year, last_number))
    conn.commit()

def get_or_create_final_letter_number(conn, req_id, payer_code):
    cursor = conn.cursor()
    cursor.execute("SELECT final_letter_number FROM approval_requests WHERE id = ?", (req_id,))
    row = cursor.fetchone()
    if row and row[0]:
        return row[0]
    year_two_digit = str(jdatetime.date.today().year)[-2:]
    next_num = get_next_letter_number(conn, payer_code, year_two_digit)
    update_letter_sequence(conn, payer_code, year_two_digit, next_num)
    format_str = "YYYY/CODE/NNNNN"
    today = jdatetime.date.today()
    final_number = format_str.replace("YYYY", str(today.year)).replace("YY", year_two_digit).replace("CODE", payer_code).replace("CCC", payer_code.zfill(3)).replace("NUM", str(next_num)).replace("NNNNN", str(next_num).zfill(5))
    cursor.execute("UPDATE approval_requests SET final_letter_number = ? WHERE id = ?", (final_number, req_id))
    conn.commit()
    return final_number

# ==================== مدیریت قالب پویا ====================
def ensure_template_table():
    """ایجاد جدول invoice_templates در صورت عدم وجود"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='invoice_templates' AND xtype='U')
        CREATE TABLE invoice_templates (
            id INT IDENTITY(1,1) PRIMARY KEY,
            name NVARCHAR(100) NOT NULL,
            description NVARCHAR(500),
            html_content NVARCHAR(MAX),
            css_content NVARCHAR(MAX),
            settings_json NVARCHAR(MAX),
            is_active BIT DEFAULT 1,
            created_by INT NULL,
            created_at DATETIME DEFAULT GETDATE(),
            updated_at DATETIME DEFAULT GETDATE()
        )
    """)
    conn.commit()
    conn.close()

def get_active_template():
    """دریافت قالب فعال از دیتابیس، در صورت نبود، قالب پیش‌فرض را برمی‌گرداند"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT html_content, css_content, settings_json FROM invoice_templates WHERE is_active=1")
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            'html': row[0] or '',
            'css': row[1] or '',
            'settings': json.loads(row[2]) if row[2] else {}
        }
    # قالب پیش‌فرض
    default_html = """
    <div class="invoice-header">صورتحساب {{ service_type }} &nbsp;&nbsp; شماره: {{ invoice_number }} &nbsp;&nbsp; تاریخ: {{ invoice_date }}</div>
    <div class="box-title">مشخصات فروشنده</div>
    <table class="seller-box">
        <tr><td class="label">شناسه ملی:ERCERD<td colspan="3">{{ seller.national_id }}ERCERD</td></tr>
        <tr><td class="label">نام شخص حقوقی:ERCERD<td colspan="3">{{ seller.company_name }}ERCERD</td></tr>
        <tr><td class="label">نشانی:ERCERD<td colspan="3">{{ seller.address }}ERCERD</td></tr>
        <tr><td class="label">کدپستی:ERCERD<td colspan="3">{{ seller.postal_code }}ERCERD<td class="label">تلفن:ERCERD<td colspan="3">{{ seller.phone }}ERCERD</td></tr>
    </table>
    <div class="box-title">مشخصات خریدار</div>
    <table class="buyer-box">
        <tr><td class="label">شناسه ملی:ERCERD<td colspan="3">{{ buyer.national_id }}ERCERD</td></tr>
        <tr><td class="label">نام شخص حقوقی:ERCERD<td colspan="3">{{ buyer.name }}ERCERD</td></tr>
        <tr><td class="label">نشانی:ERCERD<td colspan="3">{{ buyer.address }}ERCERD</td></tr>
        <tr><td class="label">کدپستی:ERCERD<td colspan="3">{{ buyer.postal_code }}ERCERD<td class="label">تلفن:ERCERD<td colspan="3">{{ buyer.phone }}ERCERD</td></tr>
    </table>
    <table class="data-table"><thead><tr><th>ردیف</th><th>شماره قرارداد</th><th>نام مسافر</th><th>شرح خدمات</th><th>تاریخ</th><th>توضیحات</th><th>بدهکار (ریال)</th><th>بستانکار (ریال)</th><th>مانده (ریال)</th></tr></thead>
    <tbody>{% for row in page.rows %}<tr>
        <td>{{ loop.index + (page.page_num-1)*8 }}ERCERD
        <td>{{ row.contract }}ERCERD
        <td>{{ row.passenger }}ERCERD
        <td>{{ row.description }}ERCERD
        <td>{{ row.date }}ERCERD
        <td>{{ row.notes }}ERCERD
        <td>{{ row.debt|persian_number }}ERCERD
        <td>{{ row.credit|persian_number }}ERCERD
        <td>{{ row.balance|persian_number }}ERCERD
    </tr>{% endfor %}</tbody></table>
    <div class="totals-row"><div>جمع بدهکار این صفحه: {{ page.total_debt|persian_number }} ریال</div><div>جمع بستانکار این صفحه: {{ page.total_credit|persian_number }} ریال</div><div>مانده این صفحه: {{ page.total_balance|persian_number }} ریال</div></div>
    {% if page.is_last %}<div class="totals-row"><div>جمع کل نهایی بدهکار: {{ grand_total_debt|persian_number }} ریال</div><div>جمع کل نهایی بستانکار: {{ grand_total_credit|persian_number }} ریال</div><div>مانده نهایی: {{ grand_total_balance|persian_number }} ریال</div></div>
    <div class="footer"><div class="signature">{% if seller.signature_path %}<img src="file:///{{ seller.signature_path }}">{% else %}<p>................................</p>{% endif %}<p>{{ seller.manager_name }}<br>{{ seller.manager_position }}</p></div>
    <div class="stamp">{% if seller.stamp_path %}<img src="file:///{{ seller.stamp_path }}">{% else %}<p>................................</p>{% endif %}<p>مهر شرکت</p></div></div>{% endif %}
    """
    default_css = """
        @page { size: A4 landscape; margin: 1cm; }
        body { font-family: 'B Nazanin', Tahoma, Arial, sans-serif; font-size: 9px; line-height: 1.2; }
        .page { page-break-after: avoid; }
        .seller-box, .buyer-box { border: 1px solid #000; border-collapse: collapse; width: 100%; margin-bottom: 6px; }
        .seller-box td, .buyer-box td { border: 1px solid #000; padding: 4px; vertical-align: top; }
        .seller-box .label, .buyer-box .label { font-weight: bold; background-color: #f0f0f0; width: 25%; }
        .box-title { font-weight: bold; margin: 3px 0; text-align: center; background-color: #e0e0e0; padding: 3px; font-size: 10px; }
        .invoice-header { text-align: center; font-size: 14px; font-weight: bold; margin-bottom: 10px; }
        .data-table { width: 100%; border-collapse: collapse; margin: 8px 0; border: 1px solid #000; }
        .data-table th, .data-table td { border: 1px solid #000; padding: 3px; text-align: center; }
        .data-table th { background-color: #f0f0f0; font-weight: bold; font-size: 9px; }
        .totals-row { margin-top: 6px; display: flex; justify-content: space-between; border-top: 1px solid #000; padding-top: 4px; font-size: 9px; }
        .footer { margin-top: 20px; display: flex; justify-content: space-between; align-items: flex-end; }
    """
    return {'html': default_html, 'css': default_css, 'settings': {}}

def save_active_template(html, css, settings, created_by=1):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE invoice_templates SET is_active=0")
    cursor.execute("""
        INSERT INTO invoice_templates (name, description, html_content, css_content, settings_json, is_active, created_by)
        VALUES (?, ?, ?, ?, ?, 1, ?)
    """, ('قالب اصلی', 'قالب سفارشی کاربر', html, css, json.dumps(settings), created_by))
    conn.commit()
    conn.close()

# ==================== تولید PDF با قالب پویا ====================
def generate_invoice_pdf_in_memory(pdf_context):
    def persian_number_filter(value):
        try:
            if value is None or pd.isna(value):
                return "۰"
            return to_persian_digits(f"{int(value):,}")
        except:
            return to_persian_digits(f"{value:,}") if value else "۰"

    try:
        env = Environment()
        env.filters['persian_number'] = persian_number_filter

        template_data = get_active_template()
        html_template = template_data['html']
        css_style = template_data['css']

        rows = pdf_context.get('rows', [])
        if not rows:
            raise ValueError("هیچ ردیفی برای تولید PDF وجود ندارد")

        chunk_size = 8
        chunks = [rows[i:i+chunk_size] for i in range(0, len(rows), chunk_size)]
        page_data = []
        for idx, chunk in enumerate(chunks):
            page_total_debt = sum(int(r.get('debt', 0)) for r in chunk)
            page_total_credit = sum(int(r.get('credit', 0)) for r in chunk)
            page_total_balance = sum(int(r.get('balance', 0)) for r in chunk)
            page_data.append({
                'page_num': idx+1,
                'rows': chunk,
                'total_debt': page_total_debt,
                'total_credit': page_total_credit,
                'total_balance': page_total_balance,
                'is_last': (idx == len(chunks)-1)
            })

        template_context = {
            'pages': page_data,
            'seller': pdf_context.get('seller', {}),
            'buyer': pdf_context.get('buyer', {}),
            'service_type': pdf_context.get('service_type', ''),
            'invoice_number': pdf_context.get('invoice_number', ''),
            'invoice_date': pdf_context.get('invoice_date', ''),
            'grand_total_debt': pdf_context.get('total_debt', 0),
            'grand_total_credit': pdf_context.get('total_credit', 0),
            'grand_total_balance': pdf_context.get('total_balance', 0)
        }

        full_html = f"""
        <!DOCTYPE html>
        <html dir="rtl" lang="fa">
        <head><meta charset="UTF-8"><style>{css_style}</style></head>
        <body>
        {html_template}
        </body>
        </html>
        """
        rendered_html = Template(full_html).render(**template_context)

        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        html_debug_path = os.path.join(base_dir, "last_invoice.html")
        with open(html_debug_path, 'w', encoding='utf-8') as f:
            f.write(rendered_html)

        edge_paths = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                      r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]
        edge_exe = next((p for p in edge_paths if os.path.exists(p)), None)
        if not edge_exe:
            logging.error("Microsoft Edge یافت نشد")
            return io.BytesIO(b"%PDF-1.4 dummy")

        pdf_output_path = os.path.join(base_dir, "temp_output.pdf")
        subprocess.run([edge_exe, "--headless", f"--print-to-pdf={pdf_output_path}", "--no-pdf-header-footer", html_debug_path], check=True, timeout=60)
        with open(pdf_output_path, 'rb') as pf:
            pdf_bytes = pf.read()
        os.unlink(pdf_output_path)
        return io.BytesIO(pdf_bytes)

    except Exception as e:
        logging.error(f"خطا در generate_invoice_pdf_in_memory: {traceback.format_exc()}")
        return io.BytesIO(b"%PDF-1.4 dummy")

# ==================== توابع اصلی تولید فایل برای یک درخواست ====================
def generate_files_for_request(req_dict, storage_dir):
    import json
    service_data = json.loads(req_dict.get('service_data_json', '[]'))
    flight_data = json.loads(req_dict.get('flight_data_json', '[]'))
    hotel_data = json.loads(req_dict.get('hotel_data_json', '[]'))

    payer_code = req_dict.get('payer_code')
    payer_name = req_dict.get('payer_name')
    letter_date = req_dict.get('letter_date')
    period_range = req_dict.get('period_range')
    req_id = req_dict.get('id')

    if not payer_code or not payer_name:
        raise ValueError("payer_code یا payer_name خالی است")

    conn = get_db_connection()
    try:
        final_letter_base = get_or_create_final_letter_number(conn, req_id, payer_code)
    except Exception as e:
        conn.close()
        raise ValueError(f"خطا در تخصیص شماره نامه: {str(e)}")

    today_shamsi = jdatetime.date.today()
    date_folder = today_shamsi.strftime("%Y-%m-%d")
    safe_payer = safe_filename(payer_name)
    final_dir = os.path.join(storage_dir, date_folder, safe_payer)
    os.makedirs(final_dir, exist_ok=True)

    company_settings = get_company_settings_from_db(conn)
    cursor = conn.cursor()
    cursor.execute("SELECT national_id, economic_code, address, postal_code, phone FROM payers WHERE code=?", (payer_code,))
    payer_row = cursor.fetchone()
    if not payer_row:
        conn.close()
        raise ValueError(f"طرف حساب با کد {payer_code} یافت نشد")
    buyer_info = {
        'name': payer_name,
        'national_id': payer_row[0] or '',
        'economic_code': payer_row[1] or '',
        'address': payer_row[2] or '',
        'postal_code': payer_row[3] or '',
        'phone': payer_row[4] or ''
    }

    service_packages = []

    # پروازها
    if flight_data:
        flight_rows = []
        flight_total = 0
        for idx, flt in enumerate(flight_data):
            try:
                debt = persian_to_english_number_strict(flt.get('debt'), f"بدهکار پرواز ردیف {idx+1}")
                credit = persian_to_english_number_strict(flt.get('credit'), f"بستانکار پرواز ردیف {idx+1}")
                balance = debt - credit
                contract = safe_str_strict(flt.get('contract', ''), f"شماره قرارداد پرواز ردیف {idx+1}")
                passenger = safe_str_strict(flt.get('passenger', ''), f"نام مسافر پرواز ردیف {idx+1}")
                route = safe_str_strict(flt.get('route', ''), f"مسیر پرواز ردیف {idx+1}")
                date_val = flt.get('date', '')
                if pd.isna(date_val):
                    date_val = ''
                notes = flt.get('notes', '')
                if pd.isna(notes):
                    notes = ''
                flight_rows.append({
                    'contract': contract,
                    'passenger': passenger,
                    'description': route,
                    'date': str(date_val),
                    'notes': str(notes),
                    'debt': debt,
                    'credit': credit,
                    'balance': balance
                })
                flight_total += balance
            except Exception as e:
                conn.close()
                raise ValueError(f"خطا در رکورد پرواز شماره {idx+1}: {str(e)}")
        if flight_rows:
            service_packages.append({'type': 'پرواز', 'total_balance': flight_total, 'rows': flight_rows})

    # هتل‌ها
    if hotel_data:
        hotel_rows = []
        hotel_total = 0
        for idx, htl in enumerate(hotel_data):
            try:
                debt = persian_to_english_number_strict(htl.get('debt'), f"بدهکار هتل ردیف {idx+1}")
                credit = persian_to_english_number_strict(htl.get('credit'), f"بستانکار هتل ردیف {idx+1}")
                balance = debt - credit
                contract = safe_str_strict(htl.get('contract', ''), f"شماره قرارداد هتل ردیف {idx+1}")
                passenger = safe_str_strict(htl.get('passenger', ''), f"نام مسافر هتل ردیف {idx+1}")
                hotel_name = safe_str_strict(htl.get('hotel', ''), f"نام هتل ردیف {idx+1}")
                room_type = safe_str_strict(htl.get('room', ''), f"نوع اتاق هتل ردیف {idx+1}")
                description = f"{hotel_name} - {room_type}"
                pax_val = htl.get('pax', 1)
                try:
                    pax = int(persian_to_english_number_strict(pax_val, f"تعداد نفرات هتل ردیف {idx+1}"))
                    if pax < 1:
                        pax = 1
                except:
                    pax = 1
                date_raw = htl.get('date', '')
                if pd.isna(date_raw):
                    date_raw = ''
                notes = htl.get('notes', '')
                if pd.isna(notes):
                    notes = ''
                if not notes and pax > 1:
                    notes = f"تعداد نفرات: {pax}"
                hotel_rows.append({
                    'contract': contract,
                    'passenger': passenger,
                    'description': description,
                    'date': str(date_raw),
                    'notes': str(notes),
                    'debt': debt,
                    'credit': credit,
                    'balance': balance
                })
                hotel_total += balance
            except Exception as e:
                conn.close()
                raise ValueError(f"خطا در رکورد هتل شماره {idx+1}: {str(e)}")
        if hotel_rows:
            service_packages.append({'type': 'هتل', 'total_balance': hotel_total, 'rows': hotel_rows})

    # خدمات
    if service_data:
        service_rows = []
        service_total = 0
        for idx, svc in enumerate(service_data):
            try:
                debt = persian_to_english_number_strict(svc.get('debt'), f"بدهکار خدمات ردیف {idx+1}")
                credit = persian_to_english_number_strict(svc.get('credit'), f"بستانکار خدمات ردیف {idx+1}")
                balance = debt - credit
                contract = safe_str_strict(svc.get('contract', ''), f"شماره قرارداد خدمات ردیف {idx+1}")
                passenger = safe_str_strict(svc.get('passenger', ''), f"نام مسافر خدمات ردیف {idx+1}")
                svc_type = safe_str_strict(svc.get('type', 'سایر'), f"نوع خدمات ردیف {idx+1}")
                date_val = svc.get('date', '')
                if pd.isna(date_val):
                    date_val = ''
                notes = svc.get('notes', '')
                if pd.isna(notes):
                    notes = ''
                service_rows.append({
                    'contract': contract,
                    'passenger': passenger,
                    'description': svc_type,
                    'date': str(date_val),
                    'notes': str(notes),
                    'debt': debt,
                    'credit': credit,
                    'balance': balance
                })
                service_total += balance
            except Exception as e:
                conn.close()
                raise ValueError(f"خطا در رکورد خدمات شماره {idx+1}: {str(e)}")
        if service_rows:
            service_packages.append({'type': 'سایر خدمات', 'total_balance': service_total, 'rows': service_rows})

    if not service_packages:
        conn.close()
        raise ValueError("هیچ داده معتبری (پرواز، هتل، خدمات) برای تولید وجود ندارد")

    result_files = []
    try:
        for idx, package in enumerate(service_packages):
            current_letter = final_letter_base if idx == 0 else f"{final_letter_base}-{idx+1}"
            safe_letter = safe_filename(current_letter)
            total_balance = package['total_balance']
            is_creditor = total_balance < 0
            template_path = resource_path("template_creditor.docx" if is_creditor else "template.docx")

            # اکسل
            df_excel = pd.DataFrame(package['rows'])
            column_mapping = {
                'contract': 'شماره قرارداد',
                'passenger': 'نام مسافر',
                'description': 'شرح خدمات',
                'date': 'تاریخ',
                'notes': 'توضیحات',
                'debt': 'بدهکار',
                'credit': 'بستانکار',
                'balance': 'مانده'
            }
            df_excel = df_excel.rename(columns=column_mapping)
            df_excel.insert(0, 'ردیف', range(1, len(df_excel)+1))
            excel_path = os.path.join(final_dir, f"صورت حساب {package['type']} {safe_payer} {safe_letter}.xlsx")
            df_excel.to_excel(excel_path, index=False)

            # Word
            if os.path.exists(template_path):
                word_output = os.path.join(final_dir, f"اعلامیه {package['type']} {safe_payer} {safe_letter}.docx")
                doc = DocxTemplate(template_path)
                context_word = {
                    'letter_date': '/'.join(reversed(letter_date.split('/'))) if '/' in letter_date else letter_date,
                    'letter_number': '/'.join(reversed(current_letter.split('/'))) if '/' in current_letter else current_letter,
                    'payer': buyer_info['name'],
                    'subsidiary': payer_name,
                    'service_type': package['type'],
                    'period_range': period_range,
                    'total_debt': f"{abs(total_balance):,}" if not is_creditor else "",
                    'total_credit': f"{abs(total_balance):,}" if is_creditor else "",
                    'bank_name1': company_settings.get('bank_name1', ''),
                    'account_number1': company_settings.get('account_number1', ''),
                    'shaba_number1': company_settings.get('shaba_number1', ''),
                    'bank_name2': company_settings.get('bank_name2', ''),
                    'account_number2': company_settings.get('account_number2', ''),
                    'shaba_number2': company_settings.get('shaba_number2', ''),
                    'company_name': company_settings.get('company_name', ''),
                    'manager_name': company_settings.get('manager_name', ''),
                    'manager_position': company_settings.get('manager_position', '')
                }
                doc.render(context_word)
                doc.save(word_output)

            # PDF
            pdf_context = {
                'seller': {
                    'company_name': company_settings.get('company_name', ''),
                    'national_id': company_settings.get('national_id', ''),
                    'economic_code': company_settings.get('economic_code', ''),
                    'address': company_settings.get('address', ''),
                    'postal_code': company_settings.get('postal_code', ''),
                    'phone': company_settings.get('phone', ''),
                    'manager_name': company_settings.get('manager_name', ''),
                    'manager_position': company_settings.get('manager_position', ''),
                    'signature_path': company_settings.get('signature_path', ''),
                    'stamp_path': company_settings.get('stamp_path', '')
                },
                'buyer': buyer_info,
                'invoice_number': current_letter,
                'invoice_date': letter_date,
                'service_type': package['type'],
                'rows': package['rows'],
                'total_debt': abs(total_balance) if not is_creditor else 0,
                'total_credit': abs(total_balance) if is_creditor else 0,
                'total_balance': total_balance
            }
            pdf_bytes_io = generate_invoice_pdf_in_memory(pdf_context)
            pdf_bytes = pdf_bytes_io.getvalue()
            pdf_path = os.path.join(final_dir, f"فاکتور {package['type']} {safe_payer} {safe_letter}.pdf")
            with open(pdf_path, 'wb') as f:
                f.write(pdf_bytes)

            result_files.append({
                'letter_number': current_letter,
                'service_type': package['type'],
                'file_name': f"فاکتور {package['type']} {safe_payer} {safe_letter}.pdf",
                'file_data': pdf_bytes,
                'disk_path': pdf_path
            })
        conn.close()
        return result_files
    except Exception as e:
        conn.close()
        raise e

# ==================== ENDPOINT های Flask ====================
@app.route('/generate', methods=['POST'])
def generate_final():
    data = request.get_json()
    req_id = data.get('request_id')
    if not req_id:
        return jsonify({'error': 'request_id required'}), 400
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM approval_requests WHERE id = ? AND status = 'pending'", (req_id,))
    row = cursor.fetchone()
    if not row:
        return jsonify({'error': 'Request not found or not pending'}), 404
    columns = [desc[0] for desc in cursor.description]
    req_dict = dict(zip(columns, row))
    try:
        files_info = generate_files_for_request(req_dict, FINAL_STORAGE)
        for f in files_info:
            cursor.execute("""
                INSERT INTO final_files (request_id, letter_number, service_type, file_name, file_path, file_data)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (req_id, f['letter_number'], f['service_type'], f['file_name'], f['disk_path'], pyodbc.Binary(f['file_data'])))
        cursor.execute("UPDATE approval_requests SET status = 'approved', approved_at = GETDATE() WHERE id = ?", (req_id,))
        conn.commit()
        return jsonify({'status': 'success', 'files_count': len(files_info)})
    except Exception as e:
        conn.rollback()
        error_msg = str(e)
        logging.error(f"خطا در generate برای request_id {req_id}: {traceback.format_exc()}")
        return jsonify({'error': error_msg}), 500
    finally:
        conn.close()

@app.route('/download/<int:file_id>')
def download_file(file_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT file_path, file_data, file_name FROM final_files WHERE id = ?", (file_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return jsonify({'error': 'File not found'}), 404
    disk_path, blob_data, file_name = row
    if disk_path and os.path.exists(disk_path):
        try:
            return send_file(disk_path, as_attachment=True, download_name=file_name)
        except:
            pass
    if blob_data:
        return send_file(io.BytesIO(blob_data), as_attachment=True, download_name=file_name, mimetype='application/pdf')
    return jsonify({'error': 'No data available'}), 404

@app.route('/health')
def health():
    return jsonify({'status': 'alive'})

@app.route('/stats')
def stats():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COUNT(*) FROM final_files")
        total_files = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM approval_requests WHERE status='approved'")
        total_approved = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM final_files WHERE last_accessed IS NOT NULL")
        total_downloads = cursor.fetchone()[0]
        return jsonify({'status':'running','total_files':total_files,'total_approved_requests':total_approved,'total_downloads':total_downloads})
    except Exception as e:
        return jsonify({'error':str(e)}),500
    finally:
        conn.close()

@app.route('/regenerate', methods=['POST'])
def regenerate():
    data = request.get_json()
    letter_number = data.get('letter_number')
    if not letter_number:
        return jsonify({'error': 'letter_number required'}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, payer_code, payer_name, service_data_json, flight_data_json, hotel_data_json, letter_date, period_range FROM approval_requests WHERE final_letter_number = ?", (letter_number,))
    row = cursor.fetchone()
    if not row:
        return jsonify({'error': 'درخواست با این شماره نامه یافت نشد'}), 404

    req_id, payer_code, payer_name, service_json, flight_json, hotel_json, letter_date, period_range = row
    req_dict = {
        'id': req_id,
        'payer_code': payer_code,
        'payer_name': payer_name,
        'letter_number': letter_number,
        'letter_date': letter_date,
        'period_range': period_range,
        'service_data_json': service_json,
        'flight_data_json': flight_json,
        'hotel_data_json': hotel_json
    }
    try:
        files_info = generate_files_for_request(req_dict, FINAL_STORAGE)
        cursor.execute("DELETE FROM final_files WHERE letter_number = ?", (letter_number,))
        for f in files_info:
            cursor.execute("""
                INSERT INTO final_files (request_id, letter_number, service_type, file_name, file_path, file_data)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (req_id, f['letter_number'], f['service_type'], f['file_name'], f['disk_path'], pyodbc.Binary(f['file_data'])))
        conn.commit()
        cursor.execute("SELECT TOP 1 id FROM final_files WHERE letter_number = ?", (letter_number,))
        file_id = cursor.fetchone()[0]
        return jsonify({'status': 'success', 'file_id': file_id})
    except Exception as e:
        conn.rollback()
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

@app.route('/reserve-letter', methods=['POST'])
def reserve_letter():
    data = request.get_json()
    payer_code = data.get('payer_code')
    if not payer_code:
        return jsonify({'error': 'payer_code required'}), 400

    conn = get_db_connection()
    try:
        year_two_digit = str(jdatetime.date.today().year)[-2:]
        next_num = get_next_letter_number(conn, payer_code, year_two_digit)
        update_letter_sequence(conn, payer_code, year_two_digit, next_num)
        format_str = "YYYY/CODE/NNNNN"
        today = jdatetime.date.today()
        letter_number = format_str.replace("YYYY", str(today.year)).replace("YY", year_two_digit).replace("CODE", payer_code).replace("CCC", payer_code.zfill(3)).replace("NUM", str(next_num)).replace("NNNNN", str(next_num).zfill(5))
        return jsonify({'status': 'success', 'letter_number': letter_number})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

@app.route('/release-letter', methods=['POST'])
def release_letter():
    data = request.get_json()
    letter_number = data.get('letter_number')
    if not letter_number:
        return jsonify({'error': 'letter_number required'}), 400

    match = re.search(r'(\d+)$', letter_number)
    if not match:
        return jsonify({'error': 'Invalid letter number format'}), 400
    num = int(match.group(1))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT payer_code FROM approval_requests WHERE letter_number = ? AND status IN ('pending', 'rejected')", (letter_number,))
    row = cursor.fetchone()
    if not row:
        return jsonify({'error': 'No pending request found for this letter number'}), 404
    payer_code = row[0]
    year_two_digit = str(jdatetime.date.today().year)[-2:]

    cursor.execute("""
        IF NOT EXISTS (SELECT 1 FROM available_letter_numbers WHERE payer_code = ? AND year = ? AND number = ?)
        INSERT INTO available_letter_numbers (payer_code, year, number) VALUES (?, ?, ?)
    """, (payer_code, year_two_digit, num, payer_code, year_two_digit, num))
    conn.commit()
    conn.close()
    return jsonify({'status': 'released', 'letter_number': letter_number})

@app.route('/get-active-template', methods=['GET'])
def get_active_template_endpoint():
    return jsonify(get_active_template())

@app.route('/save-template', methods=['POST'])
def save_template_endpoint():
    data = request.json
    html = data.get('html', '')
    css = data.get('css', '')
    settings = data.get('settings', {})
    save_active_template(html, css, settings, created_by=1)
    return jsonify({'status': 'ok'})

@app.route('/template-editor')
def template_editor():
    editor_html = """
    <!DOCTYPE html>
    <html dir="rtl" lang="fa">
    <head>
        <meta charset="UTF-8">
        <title>ویرایشگر حرفه‌ای قالب فاکتور (آفلاین)</title>
        <style>
            * {
                box-sizing: border-box;
            }
            body {
                background: #2b2b2b;
                font-family: 'B Nazanin', Tahoma, 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                direction: rtl;
            }
            .editor-container {
                max-width: 1400px;
                margin: 0 auto;
                background: #fff;
                border-radius: 12px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.2);
                overflow: hidden;
            }
            .toolbar {
                background: #3c3c3c;
                padding: 8px 12px;
                border-bottom: 1px solid #555;
                display: flex;
                flex-wrap: wrap;
                gap: 5px;
                align-items: center;
                justify-content: flex-start;
                direction: ltr;
            }
            .toolbar-group {
                display: flex;
                gap: 3px;
                background: #4a4a4a;
                padding: 4px 6px;
                border-radius: 6px;
                margin-left: 10px;
            }
            .toolbar button {
                background: #5a5a5a;
                border: none;
                color: white;
                font-size: 14px;
                font-weight: bold;
                width: 32px;
                height: 32px;
                border-radius: 4px;
                cursor: pointer;
                transition: 0.2s;
                font-family: monospace;
            }
            .toolbar button:hover {
                background: #6e6e6e;
                transform: scale(1.02);
            }
            .toolbar button.active {
                background: #4CAF50;
                color: white;
            }
            .toolbar select, .toolbar input {
                background: #5a5a5a;
                border: none;
                color: white;
                padding: 4px 8px;
                border-radius: 4px;
                font-size: 13px;
                cursor: pointer;
                font-family: 'B Nazanin', Tahoma;
            }
            .color-preview {
                width: 20px;
                height: 20px;
                display: inline-block;
                border-radius: 4px;
                background: black;
                vertical-align: middle;
                margin-right: 4px;
            }
            .code-area {
                display: none;
                padding: 15px;
                background: #1e1e1e;
                color: #d4d4d4;
                font-family: monospace;
                font-size: 13px;
                direction: ltr;
                text-align: left;
            }
            .code-area textarea {
                width: 100%;
                height: 400px;
                background: #1e1e1e;
                color: #d4d4d4;
                border: none;
                font-family: monospace;
                font-size: 13px;
                direction: ltr;
                resize: vertical;
            }
            .editor {
                padding: 15px;
                min-height: 500px;
                background: white;
                outline: none;
                font-family: 'B Nazanin', Tahoma, Arial, sans-serif;
                font-size: 14px;
                line-height: 1.5;
                direction: rtl;
                text-align: right;
            }
            .editor:focus {
                background: #fefefe;
            }
            .status-bar {
                background: #e9e9e9;
                padding: 6px 12px;
                font-size: 12px;
                display: flex;
                justify-content: space-between;
                border-top: 1px solid #ccc;
            }
            button.primary {
                background: #4CAF50;
                padding: 8px 20px;
                width: auto;
                font-weight: bold;
            }
            button.danger {
                background: #f44336;
            }
            .message {
                position: fixed;
                bottom: 20px;
                left: 20px;
                background: #333;
                color: white;
                padding: 8px 16px;
                border-radius: 8px;
                font-size: 12px;
                z-index: 1000;
                opacity: 0;
                transition: 0.3s;
            }
            @media (max-width: 700px) {
                .toolbar { flex-direction: column; align-items: stretch; }
                .toolbar-group { justify-content: center; }
            }
        </style>
    </head>
    <body>
        <div class="editor-container">
            <div class="toolbar">
                <div class="toolbar-group">
                    <button title="پررنگ (Ctrl+B)" onclick="execCmd('bold')"><b>B</b></button>
                    <button title="کج (Ctrl+I)" onclick="execCmd('italic')"><i>I</i></button>
                    <button title="زیرخط (Ctrl+U)" onclick="execCmd('underline')"><u>U</u></button>
                    <button title="خط‌خورده" onclick="execCmd('strikeThrough')"><s>S</s></button>
                </div>
                <div class="toolbar-group">
                    <button title="چینش راست" onclick="execCmd('justifyRight')">➡️</button>
                    <button title="چینش چپ" onclick="execCmd('justifyLeft')">⬅️</button>
                    <button title="وسط‌چین" onclick="execCmd('justifyCenter')">🔲</button>
                    <button title="هم‌تراز" onclick="execCmd('justifyFull')">📏</button>
                </div>
                <div class="toolbar-group">
                    <button title="لیست شماره‌دار" onclick="execCmd('insertOrderedList')">1️⃣</button>
                    <button title="لیست بولت دار" onclick="execCmd('insertUnorderedList')">•</button>
                    <button title="کاهش تورفتگی" onclick="execCmd('outdent')">↩️</button>
                    <button title="افزایش تورفتگی" onclick="execCmd('indent')">↪️</button>
                </div>
                <div class="toolbar-group">
                    <select id="fontSelect" onchange="execCmd('fontName', this.value)">
                        <option value="B Nazanin">B Nazanin</option>
                        <option value="Tahoma">Tahoma</option>
                        <option value="Arial">Arial</option>
                        <option value="Times New Roman">Times New Roman</option>
                    </select>
                    <select id="fontSize" onchange="execCmd('fontSize', this.value)">
                        <option value="1">بسیار ریز</option>
                        <option value="2">ریز</option>
                        <option value="3">معمولی</option>
                        <option value="4">متوسط</option>
                        <option value="5">بزرگ</option>
                        <option value="6">بسیار بزرگ</option>
                    </select>
                </div>
                <div class="toolbar-group">
                    <input type="color" id="textColor" value="#000000" onchange="execCmd('foreColor', this.value)">
                    <label style="color:white;">متن</label>
                    <input type="color" id="bgColor" value="#ffffff" onchange="execCmd('hiliteColor', this.value)">
                    <label style="color:white;">پس‌زمینه</label>
                </div>
                <div class="toolbar-group">
                    <button title="ایجاد لینک" onclick="createLink()">🔗</button>
                    <button title="حذف لینک" onclick="execCmd('unlink')">🔗❌</button>
                    <button title="پاک کردن فرمت" onclick="execCmd('removeFormat')">🧹</button>
                </div>
                <div class="toolbar-group">
                    <button id="sourceBtn" title="نمایش/ویرایش کد HTML" onclick="toggleSource()">&lt;/&gt;</button>
                    <button class="primary" onclick="saveTemplate()">💾 ذخیره قالب</button>
                    <button onclick="loadTemplate()">📂 بارگذاری</button>
                    <button onclick="previewTemplate()">👁 پیش‌نمایش</button>
                </div>
            </div>
            <div id="editor" class="editor" contenteditable="true">
                <!-- محتوای پیش‌فرض (بارگذاری می‌شود) -->
            </div>
            <div id="codeArea" class="code-area">
                <textarea id="htmlSource" placeholder="کد HTML قالب را اینجا ویرایش کنید..."></textarea>
            </div>
            <div class="status-bar">
                <span>✅ ویرایشگر آفلاین حرفه‌ای</span>
                <span id="statusMsg">قالب بارگذاری نشده</span>
            </div>
        </div>
        <div id="message" class="message"></div>

        <script>
            let editorDiv = document.getElementById('editor');
            let codeAreaDiv = document.getElementById('codeArea');
            let htmlSource = document.getElementById('htmlSource');
            let sourceMode = false;
            let currentTemplate = '';

            // اجرای دستورات document.execCommand
            function execCmd(command, value = null) {
                if (sourceMode) return;
                document.execCommand(command, false, value);
                editorDiv.focus();
                updateStatus();
            }

            function createLink() {
                let url = prompt('آدرس لینک را وارد کنید:', 'https://');
                if (url) {
                    execCmd('createLink', url);
                }
            }

            function updateStatus() {
                document.getElementById('statusMsg').innerText = 'در حال ویرایش...';
            }

            function toggleSource() {
                if (!sourceMode) {
                    // رفتن به حالت کد
                    let html = editorDiv.innerHTML;
                    htmlSource.value = html;
                    editorDiv.style.display = 'none';
                    codeAreaDiv.style.display = 'block';
                    sourceMode = true;
                    document.getElementById('sourceBtn').style.background = '#f44336';
                } else {
                    // برگشت به حالت بصری
                    let newHtml = htmlSource.value;
                    editorDiv.innerHTML = newHtml;
                    codeAreaDiv.style.display = 'none';
                    editorDiv.style.display = 'block';
                    sourceMode = false;
                    document.getElementById('sourceBtn').style.background = '';
                }
            }

            async function loadTemplate() {
                try {
                    let response = await fetch('/get-active-template');
                    let data = await response.json();
                    let html = data.html || '<div><h3>قالب خالی</h3><p>می‌توانید اینجا را ویرایش کنید...</p></div>';
                    currentTemplate = html;
                    if (sourceMode) {
                        htmlSource.value = html;
                    } else {
                        editorDiv.innerHTML = html;
                    }
                    showMessage('قالب با موفقیت بارگذاری شد', 'green');
                    document.getElementById('statusMsg').innerText = 'قالب بارگذاری شد';
                } catch(e) {
                    showMessage('خطا در بارگذاری قالب', 'red');
                }
            }

            async function saveTemplate() {
                let html = sourceMode ? htmlSource.value : editorDiv.innerHTML;
                try {
                    let response = await fetch('/save-template', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ html: html, css: '', settings: {} })
                    });
                    let data = await response.json();
                    if (data.status === 'ok') {
                        showMessage('✅ قالب ذخیره شد', 'green');
                        document.getElementById('statusMsg').innerText = 'قالب ذخیره شد';
                    } else {
                        showMessage('❌ خطا در ذخیره', 'red');
                    }
                } catch(e) {
                    showMessage('خطا در ارتباط با سرور', 'red');
                }
            }

            function previewTemplate() {
                let html = sourceMode ? htmlSource.value : editorDiv.innerHTML;
                let preview = window.open();
                preview.document.write('<html dir="rtl"><head><style>body{font-family:"B Nazanin"; padding:20px;}</style></head><body>' + html + '</body></html>');
                preview.document.close();
            }

            function showMessage(msg, color) {
                let msgDiv = document.getElementById('message');
                msgDiv.innerText = msg;
                msgDiv.style.backgroundColor = color === 'green' ? '#4CAF50' : '#f44336';
                msgDiv.style.opacity = '1';
                setTimeout(() => { msgDiv.style.opacity = '0'; }, 2000);
            }

            // بارگذاری خودکار در شروع
            window.addEventListener('load', () => {
                loadTemplate();
                // تنظیم پیش‌فرض فونت و سایز
                document.getElementById('fontSelect').value = 'B Nazanin';
                document.getElementById('fontSize').value = '3';
                // فعال کردن دکمه‌های میانبر صفحه کلید
                editorDiv.addEventListener('keydown', (e) => {
                    if (e.ctrlKey && e.key === 'b') { e.preventDefault(); execCmd('bold'); }
                    if (e.ctrlKey && e.key === 'i') { e.preventDefault(); execCmd('italic'); }
                    if (e.ctrlKey && e.key === 'u') { e.preventDefault(); execCmd('underline'); }
                });
            });
        </script>
    </body>
    </html>
    """
    return render_template_string(editor_html)

# ==================== راه‌اندازی سرویس ====================
if __name__ == '__main__':
    try:
        ensure_template_table()
    except Exception as e:
        logging.error(f"خطا در ایجاد جدول قالب: {e}")
    app.run(host='0.0.0.0', port=5000, debug=False)