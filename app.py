import streamlit as st
import pdfplumber
import base64
from pyzbar.pyzbar import decode
from PIL import Image
import pandas as pd
import io
from datetime import datetime
import re

# ----------------- إعدادات النظام الأساسية -----------------
TARGET_COMPANY_NAME = "شركة بناء بيتكو للمقاولات شركة شخص واحد"
TARGET_TAX_NUMBER = "300472267500003"

USERS = {
    "admin": "12345",
    "manager": "بتكو2026",
    "employee": "0000"
}

# ----------------- دليل الحسابات المعتمد للشركة -----------------
# تم ربط الكلمات المفتاحية بأسماء الحسابات الرئيسية لتبويب الفواتير آلياً
EXPENSE_CATEGORIES = {
    "تكاليف المشاريع - مواد بناء 🧱": ["اسمنت", "أسمنت", "حديد", "خرسانة", "بلك", "طابوق", "رمل", "بحص", "دهانات", "سباكة", "كهرباء", "أخشاب", "عوازل", "جبس", "مواسير", "كيابل", "مسمار", "أدوات بناء"],
    "تكاليف المشاريع - أجور ومقاولي باطن 👷": ["مصنعية", "أجور", "مقاول", "باطن", "عامل", "عمالة", "تركيب", "تنفيذ", "حفريات", "تسوية"],
    "تكاليف المشاريع - إيجار معدات 🚜": ["إيجار", "ايجار", "معدة", "معدات", "كرين", "رافعة", "شيول", "حفار", "بوبكات", "رصاصة", "سقالات", "مضخة"],
    "مصروفات التشغيل - محروقات وصيانة ⛽": ["بنزين", "ديزل", "وقود", "زيت", "غيار", "صيانة", "كفرات", "إطارات", "بطارية", "قطع غيار", "مغسلة"],
    "مصروفات إدارية وعمومية 🗂️": ["قرطاسية", "مكتبية", "طباعة", "تصوير", "أحبار", "انترنت", "اتصالات", "مياه", "كهرباء", "ضيافة", "بوفيه", "نظافة", "رسوم", "اشتراك", "بريد", "تسويق", "إعلان"],
    "أصول ثابتة (مشتريات رأسمالية) 🏢": ["لابتوب", "كمبيوتر", "مكتب", "أثاث", "سيارة", "شاحنة", "ماكينة", "جهاز", "مكيف"]
}
# ---------------------------------------------------------

def decode_zatca_qr(base64_string):
    try:
        decoded_bytes = base64.b64decode(base64_string)
        tlv_data = {}
        i = 0
        while i < len(decoded_bytes):
            tag = decoded_bytes[i]
            length = decoded_bytes[i+1]
            value = decoded_bytes[i+2 : i+2+length].decode('utf-8')
            tlv_data[tag] = value
            i += 2 + length
        return tlv_data
    except Exception as e:
        return None

def auto_categorize_invoice(invoice_text):
    for category, keywords in EXPENSE_CATEGORIES.items():
        for keyword in keywords:
            if keyword in invoice_text:
                return category
    return "حساب معلق (غير مصنف) ❓"

def generate_excel_report(report_data_list):
    df = pd.DataFrame(report_data_list)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير الفواتير المجمع')
    return output.getvalue()

st.set_page_config(page_title="نظام إدارة ومراجعة الفواتير", page_icon="🧾", layout="wide")

# --- نظام تسجيل الدخول ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.title("🔒 تسجيل الدخول للنظام")
    st.write("نظام الفحص والتجميع والتبويب الآلي لشركة بناء بيتكو للمقاولات.")
    
    with st.form("login_form"):
        username = st.text_input("👤 اسم المستخدم")
        password = st.text_input("🔑 كلمة المرور", type="password")
        submit_button = st.form_submit_button("تسجيل الدخول")
        
        if submit_button:
            if username in USERS and USERS[username] == password:
                st.session_state['logged_in'] = True
                st.session_state['username'] = username
                st.success("✅ تم تسجيل الدخول بنجاح! جاري التوجيه...")
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")
    st.stop()

# =====================================================================
# --- واجهة التطبيق الرئيسية ---
# =====================================================================

st.sidebar.write(f"مرحباً بك: **{st.session_state['username']}** 👋")
if st.sidebar.button("🚪 تسجيل الخروج"):
    st.session_state['logged_in'] = False
    st.rerun()

st.title("نظام المراجعة والتبويب المحاسبي للفواتير 🧾📊")
st.subheader(f"الشركة: {TARGET_COMPANY_NAME}")
st.info("💡 قم برفع الفواتير. سيتم فحصها وربطها تلقائياً بدليل الحسابات الخاص بالشركة.")

uploaded_files = st.file_uploader("قم برفع ملفات الفواتير بصيغة PDF هنا", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.write(f"### 🔄 جاري معالجة وتبويب {len(uploaded_files)} ملف/ملفات...")
    
    all_reports_data = []
    progress_bar = st.progress(0)
    grand_total_amount = 0.0 
    
    for idx, uploaded_file in enumerate(uploaded_files):
        text = ""
        qr_data_extracted = None
        extracted_amount = 0.0
        
        file_report = {
            "اسم الملف": uploaded_file.name,
            "تاريخ المعالجة": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "المُراجع": st.session_state['username'],
            "نوع الفاتورة": "-",
            "الحساب المحاسبي": "جاري التبويب...",
            "اسم المورد": "-",
            "الرقم الضريبي للمورد": "-",
            "الإجمالي": 0.0,
            "حالة المطابقة لشركتنا": "-",
            "النتيجة النهائية": "-"
        }

        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    extracted_text = page.extract_text()
                    if extracted_text:
                        text += extracted_text + "\n"
                    
                    try:
                        pil_image = page.to_image(resolution=300).original
                        decoded_objects = decode(pil_image)
                        for obj in decoded_objects:
                            qr_base64 = obj.data.decode('utf-8')
                            zatca_data = decode_zatca_qr(qr_base64)
                            if zatca_data:
                                qr_data_extracted = zatca_data
                    except:
                        pass

            accounting_category = auto_categorize_invoice(text)
            file_report["الحساب المحاسبي"] = accounting_category

            is_our_company = TARGET_COMPANY_NAME in text or TARGET_TAX_NUMBER in text
            file_report["حالة المطابقة لشركتنا"] = "متطابق ✅" if is_our_company else "غير متطابق ❌"

            if qr_data_extracted:
                file_report["نوع الفاتورة"] = "ضريبية إلكترونية 🧾"
                file_report["اسم المورد"] = qr_data_extracted.get(1, "غير متوفر")
                file_report["الرقم الضريبي للمورد"] = qr_data_extracted.get(2, "غير متوفر")
                
                amount_str = qr_data_extracted.get(4, "0")
                try:
                    extracted_amount = float(amount_str)
                except:
                    extracted_amount = 0.0
                
                if qr_data_extracted.get(2, "") == TARGET_TAX_NUMBER:
                    file_report["النتيجة النهائية"] = "مقبولة (تشفير صحيح) ✅"
                else:
                    file_report["النتيجة النهائية"] = "مرفوضة ضريبياً ⚠️"

            else:
                file_report["نوع الفاتورة"] = "ورقية / بدون تشفير 📄"
                file_report["النتيجة النهائية"] = "مضافة للسجل 📁"
                
                pattern = r'(?:الإجمالي|المجموع|Total|Amount|الصافي)[^\d]*([\d.,]+)'
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    try:
                        clean_num = matches[-1].replace(',', '')
                        extracted_amount = float(clean_num)
                    except:
                        extracted_amount = 0.0

            file_report["الإجمالي"] = extracted_amount
            grand_total_amount += extracted_amount
            
        except Exception as e:
            file_report["النتيجة النهائية"] = f"خطأ في القراءة: {e}"
            file_report["الحساب المحاسبي"] = "خطأ في القراءة ❌"
        
        all_reports_data.append(file_report)
        progress_bar.progress((idx + 1) / len(uploaded_files))
        
        with st.expander(f"{file_report['نوع الفاتورة']} | {uploaded_file.name} | {extracted_amount:,.2f} ريال | {accounting_category}"):
            st.write(f"- **توجيه الحساب:** {accounting_category}")
            st.write(f"- **المطابقة للشركة:** {file_report['حالة المطابقة لشركتنا']}")
            st.write(f"- **النتيجة:** {file_report['النتيجة النهائية']}")

    st.success("🎉 تم الانتهاء من المعالجة والتوجيه المحاسبي!")
    
    st.metric(label="💰 الإجمالي العام لجميع الفواتير المرفوعة", value=f"{grand_total_amount:,.2f} ريال سعودي")
    
    excel_file = generate_excel_report(all_reports_data)
    st.divider()
    st.markdown("### 📊 تحميل التقرير النهائي المجمع")
    st.download_button(
        label="📥 تحميل التقرير الشامل (Excel)",
        data=excel_file,
        file_name=f"تقرير_الفواتير_المبوب_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
