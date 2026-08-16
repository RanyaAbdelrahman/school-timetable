import os
import re
import tempfile
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path

import streamlit as st
from supabase import create_client

st.set_page_config(page_title="⭐ نظام الإدارة الذكية للجداول المدرسية ⭐", page_icon="🏫", layout="centered")
APP_VERSION = "LICENSED-GENERATOR-2026-08-16"


def secret_value(name, default=None):
    value = st.secrets.get(name, None)
    if value not in (None, ""):
        return value
    try:
        section = st.secrets.get("supabase", {})
        if isinstance(section, dict):
            value = section.get(name, None)
            if value not in (None, ""):
                return value
    except Exception:
        pass
    return default


try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("❌ لم يتم العثور على [supabase] url/key في Streamlit Secrets.")
    st.stop()

ADMIN_PASSWORD = secret_value("ADMIN_PASSWORD")
SERVICE_ROLE_KEY = secret_value("SUPABASE_SERVICE_ROLE_KEY")
if not ADMIN_PASSWORD:
    st.error("❌ ADMIN_PASSWORD غير موجود في Streamlit Secrets.")
    st.stop()

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
admin_supabase = create_client(SUPABASE_URL, SERVICE_ROLE_KEY) if SERVICE_ROLE_KEY else None


def normalize_email(email):
    return (email or "").strip().lower()


def valid_email(email):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", normalize_email(email)))


def rpc_data(result):
    data = getattr(result, "data", None)
    if isinstance(data, list):
        return data[0] if data else None
    return data


def get_school(email):
    result = supabase.rpc("check_school_license", {"p_email": normalize_email(email)}).execute()
    data = rpc_data(result)
    return data if data and data.get("found") else None


def create_school_request(school_name, email, phone):
    result = supabase.rpc("register_school", {
        "p_school_name": school_name.strip(),
        "p_email": normalize_email(email),
        "p_phone": (phone or "").strip(),
    }).execute()
    data = rpc_data(result)
    if not data:
        raise RuntimeError("لم تُرجع قاعدة البيانات بيانات المدرسة.")
    return data, not bool(data.get("exists", False))


def generate_license_key():
    import secrets
    return "CW-" + secrets.token_hex(8).upper()


def license_state(school):
    if not school:
        return "not_found", "المدرسة غير مسجلة."
    status = str(school.get("status") or "pending").lower()
    if status == "pending":
        return "pending", "طلب التفعيل ما زال قيد المراجعة."
    if status == "rejected":
        return "rejected", "تم رفض طلب تفعيل هذه المدرسة."
    if status == "blocked":
        return "blocked", "تم إيقاف ترخيص هذه المدرسة."
    if status != "approved":
        return "blocked", "حالة الترخيص غير صالحة."
    start_raw, expiry_raw = school.get("start_date"), school.get("expiry_date")
    if not start_raw or not expiry_raw:
        return "blocked", "تمت الموافقة ولكن لم يتم تحديد مدة الترخيص بعد."
    try:
        start = date.fromisoformat(str(start_raw)[:10])
        expiry = date.fromisoformat(str(expiry_raw)[:10])
    except Exception:
        return "blocked", "تواريخ الترخيص في قاعدة البيانات غير صحيحة."
    today = date.today()
    if today < start:
        return "not_started", f"الترخيص يبدأ في {start}."
    if today > expiry:
        return "expired", f"انتهت صلاحية الترخيص في {expiry}."
    return "approved", f"الترخيص صالح حتى {expiry}."


def require_admin_client():
    if not admin_supabase:
        raise RuntimeError("SUPABASE_SERVICE_ROLE_KEY غير موجود في Streamlit Secrets.")
    return admin_supabase


def admin_schools():
    return (require_admin_client().table("Schools").select("*").order("created_at", desc=True).execute()).data or []


def admin_update_school(school_id, values):
    return require_admin_client().table("Schools").update(values).eq("id", school_id).execute()


def save_license(school_id, start_date, expiry_date):
    client = require_admin_client()
    current = client.table("Schools").select("license_key").eq("id", school_id).limit(1).execute()
    current_data = current.data[0] if current.data else {}
    license_key = current_data.get("license_key") or generate_license_key()
    return client.table("Schools").update({
        "status": "approved",
        "start_date": str(start_date),
        "expiry_date": str(expiry_date),
        "license_key": license_key,
    }).eq("id", school_id).execute()


st.markdown("""
<style>
.stApp { background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%); font-family: 'Cairo','Segoe UI',Tahoma,sans-serif; }
.main-header { background: linear-gradient(135deg,#6366f1 0%,#a855f7 100%); padding:30px; border-radius:20px; color:white; text-align:center; margin-bottom:25px; }
.main-header h1 { font-size:30px; font-weight:800; margin-bottom:10px; }
.main-header p { font-size:16px; margin:0; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'><h1>⭐ نظام الإدارة الذكية للجداول المدرسية ⭐</h1><p>Code Wonders Academy</p></div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 🔐 الإدارة")
    st.caption(APP_VERSION)
    admin_mode = st.checkbox("فتح لوحة الإدارة")

if admin_mode:
    st.title("🔐 لوحة إدارة التراخيص")
    entered_password = st.text_input("كلمة مرور الإدارة", type="password")
    if entered_password != ADMIN_PASSWORD:
        if entered_password:
            st.error("❌ كلمة المرور غير صحيحة.")
        st.info("أدخل كلمة مرور الإدارة لعرض الطلبات.")
        st.stop()
    st.success("✅ تم تسجيل دخول الإدارة.")
    if not SERVICE_ROLE_KEY:
        st.error("❌ لوحة الإدارة تحتاج SUPABASE_SERVICE_ROLE_KEY في Streamlit Secrets.")
        st.stop()
    try:
        schools = admin_schools()
    except Exception as e:
        st.error(f"❌ تعذر قراءة جدول Schools: {e}")
        st.stop()
    if not schools:
        st.info("لا توجد مدارس مسجلة حتى الآن.")
    else:
        st.write(f"### 📋 عدد المدارس: {len(schools)}")
        for school in schools:
            school_id = school.get("id")
            school_name = school.get("school_name") or "بدون اسم"
            email = school.get("email") or ""
            phone = school.get("phone") or ""
            status = school.get("status") or "pending"
            start_raw = school.get("start_date")
            expiry_raw = school.get("expiry_date")
            license_key = school.get("license_key") or "-"
            with st.expander(f"🏫 {school_name} — {status.upper()}"):
                st.write(f"**Email:** {email}")
                st.write(f"**Phone:** {phone}")
                st.write(f"**Status:** {status}")
                st.write(f"**License:** {license_key}")
                try: start_default = date.fromisoformat(str(start_raw)[:10]) if start_raw else date.today()
                except Exception: start_default = date.today()
                try: expiry_default = date.fromisoformat(str(expiry_raw)[:10]) if expiry_raw else date.today()
                except Exception: expiry_default = date.today()
                c1, c2 = st.columns(2)
                with c1: new_start = st.date_input("بداية الترخيص", value=start_default, key=f"start_{school_id}")
                with c2: new_expiry = st.date_input("نهاية الترخيص", value=expiry_default, key=f"expiry_{school_id}")
                a,b,c = st.columns(3)
                with a:
                    if st.button("✅ اعتماد", key=f"approve_{school_id}", use_container_width=True):
                        if new_expiry < new_start:
                            st.error("تاريخ الانتهاء يجب أن يكون بعد البداية.")
                        else:
                            try:
                                save_license(school_id, new_start, new_expiry)
                                st.success("تم اعتماد المدرسة.")
                                st.rerun()
                            except Exception as e: st.error(f"خطأ أثناء الاعتماد: {e}")
                with b:
                    if st.button("❌ رفض", key=f"reject_{school_id}", use_container_width=True):
                        try:
                            admin_update_school(school_id, {"status":"rejected"})
                            st.success("تم رفض الطلب."); st.rerun()
                        except Exception as e: st.error(f"خطأ: {e}")
                with c:
                    if st.button("⛔ إيقاف", key=f"block_{school_id}", use_container_width=True):
                        try:
                            admin_update_school(school_id, {"status":"blocked"})
                            st.success("تم إيقاف الترخيص."); st.rerun()
                        except Exception as e: st.error(f"خطأ: {e}")
    st.stop()

if "school_verified" not in st.session_state: st.session_state.school_verified = False
if "school_record" not in st.session_state: st.session_state.school_record = None

if not st.session_state.school_verified:
    st.markdown("## 🔐 تفعيل البرنامج")
    school_name = st.text_input("🏫 اسم المدرسة", placeholder="اكتب اسم المدرسة")
    email = st.text_input("📧 البريد الإلكتروني", placeholder="school@example.com")
    phone = st.text_input("📱 رقم الهاتف", placeholder="01xxxxxxxxx")
    if st.button("📨 التحقق / طلب تفعيل", use_container_width=True):
        if not school_name.strip(): st.warning("⚠️ اكتب اسم المدرسة."); st.stop()
        if not valid_email(email): st.warning("⚠️ اكتب بريدًا إلكترونيًا صحيحًا."); st.stop()
        try:
            school, created = create_school_request(school_name, email, phone)
            if created:
                st.success("✅ تم إرسال طلب التفعيل بنجاح. انتظر موافقة الإدارة.")
                st.info("يمكنك العودة لاحقًا واستخدام نفس البريد الإلكتروني للتحقق من حالة الطلب.")
            else:
                state, message = license_state(school)
                if state == "approved":
                    st.session_state.school_verified = True
                    st.session_state.school_record = school
                    st.rerun()
                elif state == "pending": st.warning("⏳ " + message)
                elif state in ("expired","rejected","blocked"): st.error("❌ " + message)
                elif state == "not_started": st.warning("⏳ " + message)
                else: st.info("ℹ️ " + message)
        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء الاتصال بقاعدة البيانات: {e}")
    st.info("إذا كانت المدرسة مسجلة بالفعل، استخدم نفس البريد الإلكتروني المسجل لدى الإدارة.")
    st.stop()

school_record = st.session_state.school_record or {}
try:
    current_school = get_school(school_record.get("email", ""))
except Exception as e:
    st.session_state.school_verified = False
    st.session_state.school_record = None
    st.error(f"❌ تعذر التحقق من الترخيص: {e}")
    st.stop()

state, message = license_state(current_school)
if state != "approved":
    st.session_state.school_verified = False
    st.session_state.school_record = None
    st.error("❌ " + message if state in ("expired","rejected","blocked") else "⏳ " + message)
    st.stop()

st.success(f"✅ الترخيص مفعل للمدرسة: **{current_school.get('school_name','')}** | ساري حتى: **{str(current_school.get('expiry_date',''))[:10]}**")
if st.button("🚪 تسجيل الخروج"):
    st.session_state.school_verified = False
    st.session_state.school_record = None
    st.rerun()

try:
    from generator_final import generate_timetable
except Exception as e:
    st.error(f"❌ تعذر تحميل generator_final.py: {e}")
    st.stop()

st.markdown("### 📂 اختر ملف البيانات بصيغة Excel (inputs.xlsx)")
uploaded_file = st.file_uploader("ارفع ملف Excel", type=["xlsx"], label_visibility="collapsed")
if uploaded_file is not None:
    st.success(f"📄 الملف: {uploaded_file.name}")
    if st.button("🚀 إنشاء الجدول المدرسي", use_container_width=True):
        with st.spinner("✨ جاري معالجة البيانات وبناء الجداول بدقة، يرجى الانتظار..."):
            workdir = tempfile.mkdtemp(prefix="school_timetable_")
            Path(workdir, "inputs.xlsx").write_bytes(uploaded_file.getvalue())
            os.environ["TIMETABLE_WORKDIR"] = workdir
            os.environ["SCHOOL_NAME"] = str(current_school.get("school_name") or "")
            output_path = Path(workdir, "final_timetable.xlsx")
            log_buffer = StringIO()
            try:
                with redirect_stdout(log_buffer):
                    generate_timetable()
                log_text = log_buffer.getvalue()
                if not output_path.exists():
                    st.error("❌ لم يتم إنشاء ملف final_timetable.xlsx.")
                    if log_text.strip():
                        with st.expander("📋 تفاصيل عملية التوليد"): st.text(log_text[-8000:])
                else:
                    st.success("🎉 تم إنشاء الجدول بنجاح!")
                    st.download_button("📥 تحميل الجدول النهائي", data=output_path.read_bytes(), file_name="final_timetable.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    if log_text.strip():
                        with st.expander("📋 تفاصيل عملية التوليد"): st.text(log_text[-8000:])
            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
                log_text = log_buffer.getvalue()
                if log_text.strip():
                    with st.expander("📋 تفاصيل المولد"): st.text(log_text[-8000:])

st.markdown("<div style='text-align:center;padding:15px;colo
