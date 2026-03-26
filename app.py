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

USERS = {
    "admin": "12345",
    "manager": "بتكو2026",
    "employee": "0000"
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
    st.write("نظام الفحص والتجميع الآلي لشركة بناء بيتكو للمقاولات.")
    
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

st.title("نظام المراجعة والتجميع الآلي للفواتير 🧾✅")
st.subheader(f"الشركة: {TARGET_COMPANY_NAME}")
st.info("💡 قم برفع الفواتير. سيتم فحصها، وتجميع الفواتير الورقية، وحساب الإجمالي للمبالغ الضريبية تلقائياً.")

uploaded_files = st.file_uploader("قم برفع ملفات الفواتير بصيغة PDF هنا", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.write(f"### 🔄 جاري معالجة {len(uploaded_files)} ملف/ملفات...")
    
    all_reports_data = []
    progress_bar = st.progress(0)
    
    # متغير لحساب إجمالي المبالغ من الفواتير الضريبية الصحيحة
    total_tax_amount = 0.0 
    
    for idx, uploaded_file in enumerate(uploaded_files):
        text = ""
        qr_data_extracted = None
        
        file_report = {
            "اسم الملف": uploaded_file.name,
            "تاريخ المعالجة": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "المُراجع": st.session_state['username'],
            "نوع الفاتورة": "جاري التحديد...",
            "حالة اسم الشركة": "غير متطابق ❌",
            "حالة الرقم الضريبي": "غير متطابق ❌",
            "اسم المورد (من الـ QR)": "-",
            "الرقم الضريبي (من الـ QR)": "-",
            "التاريخ (من الـ QR)": "-",
            "الإجمالي (من الـ QR)": 0.0, # تم تغييره لرقم بدلاً من نص لكي يجمع في الإكسل
            "مطابقة التشفير للشركة": "-",
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

            if TARGET_COMPANY_NAME in text:
                file_report["حالة اسم الشركة"] = "متطابق ✅"
            if TARGET_TAX_NUMBER in text:
                file_report["حالة الرقم الضريبي"] = "متطابق ✅"

            if qr_data_extracted:
                file_report["نوع الفاتورة"] = "ضريبية إلكترونية 🧾"
                file_report["اسم المورد (من الـ QR)"] = qr_data_extracted.get(1, "غير متوفر")
                qr_tax_num = qr_data_extracted.get(2, "غير متوفر")
                file_report["الرقم الضريبي (من الـ QR)"] = qr_tax_num
                file_report["التاريخ (من الـ QR)"] = qr_data_extracted.get(3, "غير متوفر")
                
                # استخراج المبلغ وتحويله لرقم عشري لجمعه
                amount_str = qr_data_extracted.get(4, "0")
                try:
                    amount_float = float(amount_str)
                except:
                    amount_float = 0.0
                
                file_report["الإجمالي (من الـ QR)"] = amount_float
                
                if qr_tax_num == TARGET_TAX_NUMBER:
                    file_report["مطابقة التشفير للشركة"] = "متطابق ✅"
                    file_report["النتيجة النهائية"] = "مقبولة (ضريبية صحيحة) ✅"
                    # إضافة المبلغ للإجمالي العام فقط إذا كانت الفاتورة صحيحة ومقبولة
                    total_tax_amount += amount_float
                else:
                    file_report["مطابقة التشفير للشركة"] = "غير متطابق ❌"
                    file_report["النتيجة النهائية"] = "مرفوضة (تشفير خاطئ) ⚠️"
            else:
                file_report["نوع الفاتورة"] = "ورقية / بدون باركود 📄"
                file_report["النتيجة النهائية"] = "تم التجميع (بدون مراجعة ضريبية) 📁"
                file_report["الإجمالي (من الـ QR)"] = 0.0
            
        except Exception as e:
            file_report["النتيجة النهائية"] = f"خطأ في القراءة: {e}"
        
        all_reports_data.append(file_report)
        progress_bar.progress((idx + 1) / len(uploaded_files))
        
        # عرض مختصر للنتيجة في الواجهة
        if "ضريبية" in file_report["نوع الفاتورة"]:
            icon = "✅" if "مقبولة" in file_report["النتيجة النهائية"] else "⚠️"
            with st.expander(f"🧾 {uploaded_file.name} | الإجمالي: {file_report['الإجمالي (من الـ QR)']} ريال - ({icon})"):
                st.write(f"- **مطابقة التشفير:** {file_report['مطابقة التشفير للشركة']}")
                st.write(f"- **النتيجة:** {file_report['النتيجة النهائية']}")
        else:
            with st.expander(f"📄 {uploaded_file.name} - ورقية (تم التجميع 📁)"):
                st.write("- هذه الفاتورة لا تحتوي على باركود ZATCA وتم إضافتها للتقرير كسجل ورقي.")

    st.success("🎉 تم الانتهاء من المعالجة والتجميع بنجاح!")
    
    # عرض الإجمالي بشكل بارز في التطبيق
    st.metric(label="💰 إجمالي الفواتير الضريبية المقبولة (المطابقة للشركة)", value=f"{total_tax_amount:,.2f} ريال سعودي")
    
    excel_file = generate_excel_report(all_reports_data)
    st.divider()
    st.markdown("### 📊 تحميل التقرير النهائي المجمع")
    st.download_button(
        label="📥 تحميل التقرير الشامل (Excel)",
        data=excel_file,
        file_name=f"تقرير_الفواتير_الشامل_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
