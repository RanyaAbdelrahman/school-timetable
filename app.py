import os
import hashlib
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


try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
    ADMIN_PASSWORD = st.secrets["supabase"]["ADMIN_PASSWORD"]
    SERVICE_ROLE_KEY = st.secrets["supabase"]["SUPABASE_SERVICE_ROLE_KEY"]
except Exception as e:
    st.error(
        "❌ تعذر قراءة Streamlit Secrets.\n\n"
        "يجب أن تكون القيم داخل [supabase] بهذا الشكل:\n\n"
        "[supabase]\n"
        'url = "https://....supabase.co"\n'
        'key = "sb_publishable_..."\n'
        'ADMIN_PASSWORD = "..."\n'
        'SUPABASE_SERVICE_ROLE_KEY = "sb_secret_..."\n\n'
        f"التفاصيل: {e}"
    )
    st.stop()

if not str(ADMIN_PASSWORD).strip():
    st.error("❌ ADMIN_PASSWORD فارغة داخل Streamlit Secrets.")
    st.stop()

if not str(SERVICE_ROLE_KEY).strip():
    st.error("❌ SUPABASE_SERVICE_ROLE_KEY فارغة داخل Streamlit Secrets.")
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
    from generator import generate_timetable
except Exception as e:
    st.error(f"❌ تعذر تحميل generator.py: {e}")
    st.stop()

st.markdown("### 📂 اختر ملف البيانات بصيغة Excel (inputs.xlsx)")

uploaded_file = st.file_uploader(
    "ارفع ملف Excel",
    type=["xlsx"],
    label_visibility="collapsed",
)

# ============================================================
# حالة التوليد
# ============================================================

if "uploaded_file_hash" not in st.session_state:
    st.session_state.uploaded_file_hash = None

if "generated_file_hash" not in st.session_state:
    st.session_state.generated_file_hash = None

if "generated_files" not in st.session_state:
    st.session_state.generated_files = {}

if "generation_log" not in st.session_state:
    st.session_state.generation_log = ""


if uploaded_file is not None:

    uploaded_bytes = uploaded_file.getvalue()

    # بصمة الملف تمنع إعادة التوليد لنفس Excel.
    file_hash = hashlib.sha256(uploaded_bytes).hexdigest()

    # رفع ملف جديد = تصفير حالة التوليد السابقة.
    if st.session_state.uploaded_file_hash != file_hash:

        st.session_state.uploaded_file_hash = file_hash
        st.session_state.generated_file_hash = None
        st.session_state.generated_files = {}
        st.session_state.generation_log = ""

    st.success(f"📄 الملف: {uploaded_file.name}")

    generated_data = st.session_state.generated_files.get(file_hash)

    # ========================================================
    # إنشاء الجدول مرة واحدة فقط
    # ========================================================

    if (
        st.session_state.generated_file_hash != file_hash
        or not generated_data
    ):

        if st.button(
            "🚀 إنشاء الجدول المدرسي",
            use_container_width=True,
        ):

            with st.spinner(
                "✨ جاري معالجة البيانات وبناء الجداول بدقة، يرجى الانتظار..."
            ):

                workdir = tempfile.mkdtemp(
                    prefix="school_timetable_"
                )

                Path(
                    workdir,
                    "inputs.xlsx"
                ).write_bytes(uploaded_bytes)

                os.environ["TIMETABLE_WORKDIR"] = workdir

                school_name_value = str(
                    current_school.get("school_name") or ""
                ).strip()

                os.environ["SCHOOL_NAME"] = school_name_value

                log_buffer = StringIO()

                try:

                    with redirect_stdout(log_buffer):
                        generate_timetable()

                    log_text = log_buffer.getvalue()

                    st.session_state.generation_log = log_text

                    # نفس طريقة تسمية الملفات الموجودة في generator.py
                    safe_school_name = re.sub(
                        r'[\\/:*?"<>|]+',
                        "_",
                        school_name_value,
                    )

                    safe_school_name = re.sub(
                        r"\s+",
                        " ",
                        safe_school_name,
                    ).strip(" .")

                    safe_school_name = (
                        safe_school_name
                        or "مدرسة"
                    )

                    output_path = Path(
                        workdir,
                        f"{safe_school_name}_final_timetable.xlsx",
                    )

                    master_output_path = Path(
                        workdir,
                        f"{safe_school_name}_all_classes.xlsx",
                    )

                    if not output_path.exists():

                        st.error(
                            "❌ لم يتم إنشاء ملف الجدول النهائي."
                        )

                        if log_text.strip():

                            with st.expander(
                                "📋 تفاصيل عملية التوليد"
                            ):
                                st.text(
                                    log_text[-8000:]
                                )

                    elif not master_output_path.exists():

                        st.error(
                            "❌ تم إنشاء الجدول النهائي، "
                            "لكن ملف All Classes لم يتم إنشاؤه."
                        )

                        if log_text.strip():

                            with st.expander(
                                "📋 تفاصيل عملية التوليد"
                            ):
                                st.text(
                                    log_text[-8000:]
                                )

                    else:

                        # ====================================================
                        # حفظ الملفات في session_state.
                        # بعد ذلك Download لا يعيد قراءة أو إنشاء Excel.
                        # ====================================================

                        st.session_state.generated_files[file_hash] = {

                            "final_name":
                                f"{safe_school_name}_final_timetable.xlsx",

                            "final_data":
                                output_path.read_bytes(),

                            "all_name":
                                f"{safe_school_name}_all_classes.xlsx",

                            "all_data":
                                master_output_path.read_bytes(),
                        }

                        st.session_state.generated_file_hash = file_hash

                        st.success(
                            "🎉 تم إنشاء الجدول بنجاح!"
                        )

                        # إعادة تشغيل الواجهة لإظهار الزرين
                        # خارج زر التوليد.
                        st.rerun()

                except Exception as e:

                    st.error(
                        f"❌ حدث خطأ أثناء المعالجة: {e}"
                    )

                    log_text = log_buffer.getvalue()

                    st.session_state.generation_log = log_text

                    if log_text.strip():

                        with st.expander(
                            "📋 تفاصيل المولد"
                        ):
                            st.text(
                                log_text[-8000:]
                            )

    # ========================================================
    # أزرار التحميل
    #
    # هذه المنطقة خارج زر إنشاء الجدول.
    # الضغط على Download لا يشغل generate_timetable().
    # ========================================================

    generated_data = st.session_state.generated_files.get(
        file_hash
    )

    if (
        st.session_state.generated_file_hash == file_hash
        and generated_data
    ):

        st.success(
            "✅ الجدول جاهز للتحميل."
        )

        st.download_button(
            "📥 تحميل الجدول النهائي",
            data=generated_data["final_data"],
            file_name=generated_data["final_name"],
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
            key=f"download_final_{file_hash}",
        )

        st.download_button(
            "📘 تحميل ملف All Classes الشامل",
            data=generated_data["all_data"],
            file_name=generated_data["all_name"],
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
            key=f"download_all_{file_hash}",
        )

        if st.session_state.generation_log.strip():

            with st.expander(
                "📋 تفاصيل عملية التوليد"
            ):
                st.text(
                    st.session_state.generation_log[-8000:]
                )


st.markdown(
    """
    <div style="
        text-align:center;
        padding:15px;
        color:#4f46e5;
        font-weight:bold;
    ">
        Code Wonders Academy — 01060572506
    </div>
    """,
    unsafe_allow_html=True,
)
