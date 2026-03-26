import streamlit as st
import pdfplumber
import base64
from pyzbar.pyzbar import decode
from PIL import Image
import pandas as pd
import io
from datetime import datetime

# ----------------- إعدادات النظام الأساسية -----------------
TARGET_COMPANY_NAME = "شركة بناء بيتكو للمقاولات شركة شخص واحد"
TARGET_TAX_NUMBER = "300472267500003"

# ----------------- بيانات المستخدمين (للتسجيل) -----------------
# يمكنك تعديل أسماء المستخدمين وكلمات المرور من هنا
USERS = {
    "admin": "12345",
    "manager": "بتكو2026",
    "employee": "0000"
}
# ---------------------------------------------------------

def decode_zatca_qr(base64_string):
    """دالة لفك تشفير بيانات QR Code الخاصة بهيئة الزكاة"""
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

def generate_excel_report(report_data_list):
    """دالة لتحويل قائمة البيانات إلى ملف إكسل مجمع"""
    df = pd.DataFrame(report_data_list)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير المراجعة المجمع')
    return output.getvalue()

st.set_page_config(page_title="مراجعة الفواتير المجمعة", page_icon="🧾", layout="wide")

# --- نظام تسجيل الدخول (Login System) ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.title("🔒 تسجيل الدخول للنظام")
    st.write("يرجى إدخال بيانات الاعتماد للوصول إلى نظام مراجعة الفواتير الخاص بشركة بناء بيتكو للمقاولات.")
    
    with st.form("login_form"):
        username = st.text_input("👤 اسم المستخدم")
        password = st.text_input("🔑 كلمة المرور", type="password")
        submit_button = st.form_submit_button("تسجيل الدخول")
        
        if submit_button:
            if username in USERS and USERS[username] == password:
                st.session_state['logged_in'] = True
                st.session_state['username'] = username
                st.success("✅ تم تسجيل الدخول بنجاح! جاري التوجيه...")
                st.rerun()  # إعادة تحميل الصفحة لعرض التطبيق
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")
    
    # إيقاف تنفيذ باقي الكود إذا لم يتم تسجيل الدخول
    st.stop()

# =====================================================================
# --- واجهة التطبيق الرئيسية (تظهر فقط بعد تسجيل الدخول) ---
# =====================================================================

# زر تسجيل الخروج في القائمة الجانبية
st.sidebar.write(f"مرحباً بك: **{st.session_state['username']}** 👋")
if st.sidebar.button("🚪 تسجيل الخروج"):
    st.session_state['logged_in'] = False
    st.rerun()

st.title("نظام المراجعة الشامل للفواتير الضريبية (ZATCA) 🧾✅")
st.subheader(f"الشركة: {TARGET_COMPANY_NAME}")
st.info("💡 يمكنك رفع عدة فواتير دفعة واحدة. سيقوم النظام بفحصها جميعاً وإصدار تقرير إكسل مجمع.")

uploaded_files = st.file_uploader("قم برفع ملفات الفواتير بصيغة PDF هنا", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.write(f"### 🔄 جاري معالجة {len(uploaded_files)} ملف/ملفات...")
    
    all_reports_data = []
    progress_bar = st.progress(0)
    
    for idx, uploaded_file in enumerate(uploaded_files):
        text = ""
        qr_data_extracted = None
        all_passed = True
        
        file_report = {
            "اسم الملف": uploaded_file.name,
            "تاريخ الفحص": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "المُراجع (المستخدم)": st.session_state['username'], # تسجيل اسم الموظف الذي قام بالفحص
            "حالة اسم الشركة": "غير متطابق ❌",
            "حالة الرقم الضريبي": "غير متطابق ❌",
            "اسم المورد (من الـ QR)": "غير متوفر",
            "الرقم الضريبي (من الـ QR)": "غير متوفر",
            "الإجمالي (من الـ QR)": "غير متوفر",
            "التاريخ (من الـ QR)": "غير متوفر",
            "مطابقة التشفير للشركة": "غير متطابق ❌",
            "النتيجة النهائية": "مرفوضة ⚠️"
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

            if TARGET_COMPANY_NAME in text:
                file_report["حالة اسم الشركة"] = "متطابق ✅"
            else:
                all_passed = False
                
            if TARGET_TAX_NUMBER in text:
                file_report["حالة الرقم الضريبي"] = "متطابق ✅"
            else:
                all_passed = False

            if qr_data_extracted:
                file_report["اسم المورد (من الـ QR)"] = qr_data_extracted.get(1, "غير متوفر")
                qr_tax_num = qr_data_extracted.get(2, "غير متوفر")
                file_report["الرقم الضريبي (من الـ QR)"] = qr_tax_num
                file_report["التاريخ (من الـ QR)"] = qr_data_extracted.get(3, "غير متوفر")
                file_report["الإجمالي (من الـ QR)"] = qr_data_extracted.get(4, "غير متوفر")
                
                if qr_tax_num == TARGET_TAX_NUMBER:
                    file_report["مطابقة التشفير للشركة"] = "متطابق ✅"
                else:
                    all_passed = False
            else:
                all_passed = False

            if all_passed and qr_data_extracted:
                file_report["النتيجة النهائية"] = "مقبولة وصحيحة ✅"
            
        except Exception as e:
            file_report["النتيجة النهائية"] = f"خطأ في القراءة: {e}"
        
        all_reports_data.append(file_report)
        progress_bar.progress((idx + 1) / len(uploaded_files))
        
        with st.expander(f"📄 نتيجة فحص: {uploaded_file.name} - {file_report['النتيجة النهائية']}"):
            st.write(f"- **مطابقة التشفير (QR):** {file_report['مطابقة التشفير للشركة']}")

    st.success("🎉 تم الانتهاء من فحص جميع الملفات بنجاح!")
    
    excel_file = generate_excel_report(all_reports_data)
    st.divider()
    st.markdown("### 📊 تحميل التقرير النهائي")
    st.download_button(
        label="📥 تحميل تقرير المراجعة المجمع (Excel)",
        data=excel_file,
        file_name=f"تقرير_مراجعة_مجمع_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )