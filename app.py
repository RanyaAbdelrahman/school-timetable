import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from supabase import create_client, Client
from datetime import date
import secrets
import re

# ============================================================
# إعدادات Supabase و Secrets
# ============================================================
st.set_page_config(
    page_title="⭐ نظام الإدارة الذكية للجداول المدرسية ⭐",
    page_icon="🏫",
    layout="centered"
)

try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("❌ لم يتم العثور على [supabase] url/key في Streamlit Secrets.")
    st.stop()

# ندعم وضع Secrets داخل [supabase] أو في المستوى الأعلى.
try:
    ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD")
except Exception:
    ADMIN_PASSWORD = None
if not ADMIN_PASSWORD:
    try:
        ADMIN_PASSWORD = st.secrets["supabase"].get("ADMIN_PASSWORD")
    except Exception:
        ADMIN_PASSWORD = None

try:
    SERVICE_ROLE_KEY = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
except Exception:
    SERVICE_ROLE_KEY = None
if not SERVICE_ROLE_KEY:
    try:
        SERVICE_ROLE_KEY = st.secrets["supabase"].get("SUPABASE_SERVICE_ROLE_KEY")
    except Exception:
        SERVICE_ROLE_KEY = None

if not ADMIN_PASSWORD:
    st.error("❌ ADMIN_PASSWORD غير موجود في Streamlit Secrets.")
    st.stop()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
admin_supabase = create_client(SUPABASE_URL, SERVICE_ROLE_KEY) if SERVICE_ROLE_KEY else None

# ============================================================
# أدوات الترخيص
# ============================================================
def normalize_email(email):
    return str(email or "").strip().lower()

def valid_email(email):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", normalize_email(email)))

def generate_license_key():
    return "CW-" + secrets.token_hex(8).upper()

def rpc_data(result):
    data = result.data
    if isinstance(data, list):
        return data[0] if data else None
    return data

def create_school_request(school_name, email, phone):
    result = supabase.rpc(
        "register_school",
        {
            "p_school_name": school_name.strip(),
            "p_email": normalize_email(email),
            "p_phone": str(phone or "").strip()
        }
    ).execute()
    data = rpc_data(result)
    if not data:
        return None, False
    return data, not bool(data.get("exists", False))

def get_school(email):
    result = supabase.rpc(
        "check_school_license",
        {"p_email": normalize_email(email)}
    ).execute()
    data = rpc_data(result)
    if not data or not data.get("found", False):
        return None
    return data

def license_state(school):
    if not school:
        return "not_found", "المدرسة غير مسجلة."

    status = str(school.get("status") or "pending").lower().strip()

    if status == "pending":
        return "pending", "طلب التفعيل ما زال قيد المراجعة."
    if status == "rejected":
        return "rejected", "تم رفض طلب تفعيل هذه المدرسة."
    if status in ("blocked", "disabled"):
        return "blocked", "تم إيقاف ترخيص هذه المدرسة."
    if status != "approved":
        return "blocked", "حالة الترخيص غير صالحة."

    start_raw = school.get("start_date")
    expiry_raw = school.get("expiry_date")

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
    if admin_supabase is None:
        st.error(
            "❌ لوحة الإدارة تحتاج SUPABASE_SERVICE_ROLE_KEY في Streamlit Secrets. "
            "أضفه دون وضعه في GitHub أو داخل app.py."
        )
        st.stop()
    return admin_supabase

def save_license(school_id, start_date, expiry_date, status="approved"):
    client = require_admin_client()
    current = (
        client.table("Schools")
        .select("license_key")
        .eq("id", school_id)
        .limit(1)
        .execute()
    )
    current_data = current.data[0] if current.data else {}
    license_key = current_data.get("license_key") or generate_license_key()

    result = (
        client.table("Schools")
        .update({
            "status": status,
            "start_date": str(start_date),
            "expiry_date": str(expiry_date),
            "license_key": license_key
        })
        .eq("id", school_id)
        .execute()
    )
    return result.data

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%);
    font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif;
}
.main-header {
    background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
    padding: 35px;
    border-radius: 20px;
    color: white;
    text-align: center;
    box-shadow: 0 10px 25px rgba(99, 102, 241, 0.3);
    margin-bottom: 25px;
}
.main-header h1 {
    font-size: 30px;
    font-weight: 800;
    margin-bottom: 10px;
    color: #ffffff;
}
.main-header p {
    font-size: 16px;
    color: #f3e8ff;
    margin: 0;
}
.stTextInput > div > div > input {
    border-radius: 12px;
    border: 2px solid #cbd5e1;
    padding: 12px;
    font-size: 16px;
    background-color: #ffffff;
    transition: all 0.3s ease;
}
.stTextInput > div > div > input:focus {
    border-color: #6366f1;
    box-shadow: 0 0 10px rgba(99, 102, 241, 0.25);
}
.stButton > button {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
    color: white;
    font-weight: 700;
    font-size: 18px;
    padding: 14px 20px;
    border-radius: 14px;
    border: none;
    box-shadow: 0 6px 15px rgba(16, 185, 129, 0.35);
    transition: all 0.3s ease;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #059669 0%, #047857 100%);
    box-shadow: 0 8px 20px rgba(16, 185, 129, 0.5);
    transform: translateY(-2px);
}
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #ffffff;
    color: #4f46e5;
    text-align: center;
    padding: 12px;
    font-weight: bold;
    border-top: 2px solid #e2e8f0;
    font-size: 14px;
    box-shadow: 0 -4px 10px rgba(0,0,0,0.05);
    z-index: 100;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# لوحة الإدارة
# ============================================================
with st.sidebar:
    st.markdown("## 🔐 الإدارة")
    admin_mode = st.checkbox("فتح لوحة الإدارة")

if admin_mode:
    st.markdown("# 🔐 لوحة إدارة التراخيص")
    entered_password = st.text_input("كلمة مرور الإدارة", type="password")

    if entered_password != ADMIN_PASSWORD:
        if entered_password:
            st.error("❌ كلمة المرور غير صحيحة.")
        st.info("أدخل كلمة مرور الإدارة لعرض المدارس.")
        st.stop()

    client = require_admin_client()
    st.success("✅ تم تسجيل دخول الإدارة.")

    try:
        schools_result = (
            client.table("Schools")
            .select("*")
            .order("created_at", desc=True)
            .execute()
        )
        schools = schools_result.data or []
    except Exception as e:
        st.error(f"❌ تعذر قراءة جدول Schools: {e}")
        st.stop()

    st.write(f"### 📋 عدد المدارس: {len(schools)}")

    if not schools:
        st.info("لا توجد مدارس مسجلة حتى الآن.")
    else:
        for school in schools:
            school_id = school.get("id")
            school_name = school.get("school_name") or "بدون اسم"
            email = school.get("email") or ""
            phone = school.get("phone") or ""
            status = school.get("status") or "pending"
            start_date = school.get("start_date")
            expiry_date = school.get("expiry_date")
            license_key = school.get("license_key")

            with st.expander(f"🏫 {school_name} — {status.upper()}"):
                st.write(f"**Email:** {email}")
                st.write(f"**Phone:** {phone}")
                st.write(f"**Status:** {status}")
                st.write(f"**License:** {license_key or '-'}")

                c1, c2 = st.columns(2)
                with c1:
                    try:
                        start_default = date.fromisoformat(str(start_date)[:10]) if start_date else date.today()
                    except Exception:
                        start_default = date.today()
                    new_start = st.date_input("بداية الترخيص", value=start_default, key=f"start_{school_id}")

                with c2:
                    try:
                        expiry_default = date.fromisoformat(str(expiry_date)[:10]) if expiry_date else date.today()
                    except Exception:
                        expiry_default = date.today()
                    new_expiry = st.date_input("نهاية الترخيص", value=expiry_default, key=f"expiry_{school_id}")

                a, b, c = st.columns(3)

                with a:
                    if st.button("✅ اعتماد / تمديد", key=f"approve_{school_id}", use_container_width=True):
                        if new_expiry < new_start:
                            st.error("تاريخ الانتهاء يجب أن يكون بعد تاريخ البداية.")
                        else:
                            try:
                                save_license(school_id, new_start, new_expiry, "approved")
                                st.success("تم اعتماد/تمديد الترخيص.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ خطأ أثناء الحفظ: {e}")

                with b:
                    if st.button("❌ رفض", key=f"reject_{school_id}", use_container_width=True):
                        try:
                            client.table("Schools").update({"status": "rejected"}).eq("id", school_id).execute()
                            st.success("تم رفض الطلب.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ خطأ: {e}")

                with c:
                    if st.button("⛔ إيقاف", key=f"block_{school_id}", use_container_width=True):
                        try:
                            client.table("Schools").update({"status": "blocked"}).eq("id", school_id).execute()
                            st.success("تم إيقاف الترخيص.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ خطأ: {e}")

    st.divider()
    st.caption("Admin يستخدم Service Role من Streamlit Secrets فقط. لا تضع المفتاح داخل GitHub.")
    st.stop()

# ============================================================
# بوابة المدرسة والترخيص
# ============================================================
st.markdown("""
<div class="main-header">
    <h1> ⭐ نظام الإدارة الذكية للجداول المدرسية ⭐ </h1>
    <p> Code Wonders Academy </p>
</div>
""", unsafe_allow_html=True)

if "school_verified" not in st.session_state:
    st.session_state.school_verified = False
if "school_record" not in st.session_state:
    st.session_state.school_record = None

if not st.session_state.school_verified:
    st.markdown("## 🔐 تفعيل البرنامج")

    school_input_name = st.text_input("🏫 اسم المدرسة", value="", key="school_name_input")
    email_input = st.text_input("📧 البريد الإلكتروني", value="", key="school_email_input")
    phone_input = st.text_input("📱 رقم الهاتف", value="", key="school_phone_input")

    if st.button("📨 التحقق / طلب تفعيل", use_container_width=True):
        if not school_input_name.strip():
            st.warning("⚠️ اكتب اسم المدرسة.")
            st.stop()
        if not valid_email(email_input):
            st.warning("⚠️ اكتب بريدًا إلكترونيًا صحيحًا.")
            st.stop()

        try:
            school, created = create_school_request(school_input_name, email_input, phone_input)

            if created:
                st.success("✅ تم إرسال طلب التفعيل بنجاح. انتظر موافقة الإدارة.")
                st.info("يمكنك العودة لاحقًا واستخدام نفس البريد الإلكتروني للتحقق من حالة الطلب.")
                st.stop()

            state, message = license_state(school)
            if state == "approved":
                st.session_state.school_verified = True
                st.session_state.school_record = school
                st.rerun()
            elif state == "pending":
                st.warning("⏳ " + message)
            elif state == "expired":
                st.error("❌ " + message)
            elif state == "rejected":
                st.error("❌ " + message)
            elif state == "not_started":
                st.warning("⏳ " + message)
            else:
                st.error("❌ " + message)
        except Exception as e:
            st.error(f"❌ حدث خطأ أثناء الاتصال بقاعدة البيانات: {e}")

    st.divider()
    st.info("إذا كانت المدرسة مسجلة بالفعل، استخدم نفس البريد الإلكتروني المسجل لدى الإدارة.")
    st.stop()

# إعادة فحص الترخيص في كل تشغيل للصفحة
try:
    current_school = get_school(st.session_state.school_record.get("email", ""))
except Exception as e:
    st.session_state.school_verified = False
    st.session_state.school_record = None
    st.error(f"❌ تعذر التحقق من الترخيص: {e}")
    st.stop()

state, message = license_state(current_school)
if state != "approved":
    st.session_state.school_verified = False
    st.session_state.school_record = None
    if state in ("expired", "rejected", "blocked"):
        st.error("❌ " + message)
    elif state == "not_started":
        st.warning("⏳ " + message)
    else:
        st.info("ℹ️ " + message)
    st.stop()

# ============================================================
# الترخيص ساري: يبدأ مولد الجداول الأصلي هنا
# ============================================================
school_input_name = current_school.get("school_name") or ""
st.markdown(
    f"<div style='background:#ecfdf5;padding:12px;border-radius:12px;margin-bottom:15px;'>"
    f"✅ الترخيص مفعل للمدرسة: <b>{school_input_name}</b> | "
    f"ساري حتى: <b>{current_school.get('expiry_date', '')}</b></div>",
    unsafe_allow_html=True
)

if st.button("🚪 تسجيل الخروج من الترخيص", key="license_logout"):
    st.session_state.school_verified = False
    st.session_state.school_record = None
    st.session_state.generated = False
    st.rerun()

# ============================================================
# مولد الجداول الأصلي - محفوظ كما هو من app.py المرفوع
# ============================================================
uploaded_file = st.file_uploader("📂 اختر ملف البيانات بصيغة Excel (inputs.xlsx)", type=["xlsx"])

if "generated" not in st.session_state:
    st.session_state.generated = False

if uploaded_file is not None:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 إنشاء الجدول المدرسي ", use_container_width=True):
        if not school_input_name.strip():
            st.warning("⚠️ تنبيه: يرجى كتابة اسم المدرسة أولاً ليظهر في ترويسة ملفات الإكسل الناتجة.")
        else:
            with st.spinner("✨ جاري معالجة البيانات وبناء الجداول بدقة، يرجى الانتظار..."):
                try:
                    df_teachers = pd.read_excel(uploaded_file, sheet_name="Teachers")
                    df_classes = pd.read_excel(uploaded_file, sheet_name="Classes")
                    df_assignments = pd.read_excel(uploaded_file, sheet_name="Assignments")
                    df_settings = pd.read_excel(uploaded_file, sheet_name="Settings")
                    df_days = pd.read_excel(uploaded_file, sheet_name="Days")

                    days = [str(d).strip() for d in df_days["DayName"].dropna().tolist()]
                    num_days = len(days)
                    num_periods = int(df_settings["PeriodsPerDay"].iloc[0])
                    classes = [str(c).strip() for c in df_classes["ClassName"].dropna().tolist()]
                    teachers = [str(t).strip() for t in df_teachers["Teacher"].dropna().tolist()]
                    periods = [f"الحصة {p + 1}" for p in range(num_periods)]

                    teacher_off_days = {}
                    teacher_unwanted_periods = {}
                    for _, row in df_teachers.iterrows():
                        if pd.isna(row["Teacher"]): continue
                        t_name = str(row["Teacher"]).strip()
                        off_days = clean_off_days(row.get("OffDays", []))
                        teacher_off_days[t_name] = [d for d in off_days if d in days]
                        
                        if "UnwantedPeriods" in df_teachers.columns:
                            teacher_unwanted_periods[t_name] = clean_unwanted_periods(row.get("UnwantedPeriods", ""))

                    clean_assignments = []
                    for idx, row in df_assignments.iterrows():
                        clean_assignments.append({
                            "idx": idx,
                            "c": str(row["ClassName"]).strip(),
                            "s": str(row["Subject"]).strip(),
                            "t": str(row["Teacher"]).strip(),
                            "r": str(row.get("PreferredRoom", "Classroom")).strip() if pd.notna(row.get("PreferredRoom")) else "Classroom",
                            "w": int(row["WeeklyLessons"])
                        })

                    rooms = list(set([item["r"] for item in clean_assignments if item["r"] and item["r"] != "Classroom"]))

                    model = cp_model.CpModel()
                    schedule = {}

                    for item in clean_assignments:
                        for d in range(num_days):
                            for p in range(num_periods):
                                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] = model.NewBoolVar(f"var_{item['idx']}_{d}_{p}")

                    for item in clean_assignments:
                        model.Add(sum(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] for d in range(num_days) for p in range(num_periods)) == item["w"])

                    for c in classes:
                        for d in range(num_days):
                            for p in range(num_periods):
                                rel_vars = [schedule[(i["idx"], i["c"], i["s"], i["t"], i["r"], d, p)] for i in clean_assignments if c in [x.strip() for x in i["c"].split(",")]]
                                if rel_vars: model.Add(sum(rel_vars) <= 1)

                    for t_name in teachers:
                        for d in range(num_days):
                            for p in range(num_periods):
                                t_vars = [schedule[(i["idx"], i["c"], i["s"], i["t"], i["r"], d, p)] for i in clean_assignments if t_name in [x.strip() for x in i["t"].split("/")]]
                                if t_vars: model.Add(sum(t_vars) <= 1)

                    for r_name in rooms:
                        for d in range(num_days):
                            for p in range(num_periods):
                                r_vars = [schedule[(i["idx"], i["c"], i["s"], i["t"], i["r"], d, p)] for i in clean_assignments if i["r"] == r_name]
                                if r_vars: model.Add(sum(r_vars) <= 1)

                    for item in clean_assignments:
                        for t_name in [x.strip() for x in item["t"].split("/") if x.strip()]:
                            for d, day_name in enumerate(days):
                                if day_name in teacher_off_days.get(t_name, []):
                                    for p in range(num_periods):
                                        model.Add(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] == 0)

                    last_period_idx = num_periods - 1
                    for item in clean_assignments:
                        if "PE" in item["s"].upper() or "ملعب" in item["r"]:
                            for d in range(num_days):
                                model.Add(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, last_period_idx)] == 0)

                    objective_terms = []

                    for key, var in schedule.items():
                        p = key[-1]
                        objective_terms.append(var * (num_periods - p) * 10)

                    for item in clean_assignments:
                        idx = item["idx"]
                        c = item["c"]
                        s = item["s"]
                        t = item["t"]
                        r = item["r"]

                        for d in range(num_days):
                            day_has_subject = model.NewBoolVar(f"day_has_{idx}_{d}")
                            day_lessons = [schedule[(idx, c, s, t, r, d, p)] for p in range(num_periods)]
                            model.Add(sum(day_lessons) >= 1).OnlyEnforceIf(day_has_subject)
                            model.Add(sum(day_lessons) == 0).OnlyEnforceIf(day_has_subject.Not())
                            objective_terms.append(day_has_subject * 50)

                    for item in clean_assignments:
                        for t_name in [x.strip() for x in item["t"].split("/") if x.strip()]:
                            unwanted_list = teacher_unwanted_periods.get(t_name, [])
                            if unwanted_list:
                                for d in range(num_days):
                                    for p in range(num_periods):
                                        if (p + 1) in unwanted_list:
                                            var = schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                                            objective_terms.append(var * -200)

                    model.Maximize(sum(objective_terms))

                    solver = cp_model.CpSolver()
                    solver.parameters.max_time_in_seconds = 30.0
                    status = solver.Solve(model)

                    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                        output_data = []
                        for item in clean_assignments:
                            for d in range(num_days):
                                for p in range(num_periods):
                                    if solver.Value(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]) == 1:
                                        output_data.append({"الفصل": item["c"], "المادة": item["s"], "المدرس": item["t"], "القاعة": item["r"], "اليوم": days[d], "الحصة": f"الحصة {p + 1}"})

                        df_result = pd.DataFrame(output_data)
                        out_file = "final_timetable.xlsx"
                        master_table_file = "all_classes_master_table.xlsx"

                        wb_master = Workbook()
                        ws_master = wb_master.active
                        ws_master.title = "الحصص_الشامل"
                        ws_master.sheet_view.rightToLeft = True
                        ws_master.views.sheetView[0].showGridLines = True

                        thin_border = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
                        header_fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
                        sub_header_fill = PatternFill(start_color="818CF8", end_color="818CF8", fill_type="solid")
                        title_fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
                        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

                        ws_master.merge_cells("A1:V1")
                        title_cell = ws_master.cell(row=1, column=1, value=f"{school_input_name} - جدول الحصص المدرسي الشامل لجميع الفصول")
                        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="4F46E5")
                        title_cell.alignment = center_align
                        title_cell.fill = title_fill
                        ws_master.row_dimensions[1].height = 40

                        ws_master.cell(row=3, column=1, value="م").font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                        ws_master.cell(row=3, column=1).alignment = center_align
                        ws_master.cell(row=3, column=1).fill = header_fill
                        ws_master.merge_cells("A3:A4")

                        ws_master.cell(row=3, column=2, value="اسم الفصل").font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                        ws_master.cell(row=3, column=2).alignment = center_align
                        ws_master.cell(row=3, column=2).fill = header_fill
                        ws_master.merge_cells("B3:B4")

                        current_col = 3
                        for day in days:
                            start_c = current_col
                            end_c = current_col + len(periods) - 1
                            ws_master.merge_cells(start_row=3, start_column=start_c, end_row=3, end_column=end_c)
                            day_cell = ws_master.cell(row=3, column=start_c, value=day)
                            day_cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                            day_cell.alignment = center_align
                            day_cell.fill = header_fill

                            for p_idx, p in enumerate(periods):
                                p_cell = ws_master.cell(row=4, column=start_c + p_idx, value=p)
                                p_cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                                p_cell.alignment = center_align
                                p_cell.fill = sub_header_fill
                            current_col += len(periods)

                        for idx, cls in enumerate(sorted(list(classes)), start=1):
                            row_num = 5 + idx - 1
                            ws_master.row_dimensions[row_num].height = 35
                            row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid") if idx % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                            ws_master.cell(row=row_num, column=1, value=idx).alignment = center_align
                            ws_master.cell(row=row_num, column=2, value=str(cls)).alignment = center_align

                            col_cursor = 3
                            for day in days:
                                for p in periods:
                                    match = df_result[(df_result["الفصل"] == cls) & (df_result["اليوم"] == day) & (df_result["الحصة"] == p)]
                                    if not match.empty:
                                        mat = match.iloc[0]["المادة"]
                                        tch = match.iloc[0].get("المدرس", "")
                                        cell_val = f"{mat}\n({tch})" if tch else mat
                                    else:
                                        cell_val = "متاحة"
                                    cell = ws_master.cell(row=row_num, column=col_cursor, value=cell_val)
                                    cell.alignment = center_align
                                    cell.border = thin_border
                                    cell.fill = row_fill
                                    cell.font = Font(name="Segoe UI", size=8)
                                    col_cursor += 1

                        for col in ws_master.columns:
                            max_len = max(len(str(line)) for cell in col for line in str(cell.value or "").split("\n"))
                            ws_master.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 15)

                        wb_master.save(master_table_file)

                        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                            df_result.to_excel(writer, sheet_name="Master_Schedule", index=False)
                            
                            summary_rows = []
                            for _, row in df_result.iterrows():
                                for teacher in [t.strip() for t in str(row["المدرس"]).split("/") if t.strip()]:
                                    summary_rows.append({"المعلم/ة": teacher, "الحصة": 1})
                            
                            if summary_rows:
                                df_summary = pd.DataFrame(summary_rows).groupby("المعلم/ة")["الحصة"].sum().reset_index()
                                df_summary.columns = ["المعلم/ة", "إجمالي الحصص الأسبوعية"]
                                
                                teacher_subjects = {}
                                for t in teachers:
                                    subs = df_result[df_result["المدرس"].apply(lambda x: t in [i.strip() for i in str(x).split("/")])]["المادة"].dropna().unique().tolist()
                                    teacher_subjects[t] = ", ".join(subs)
                                
                                df_summary["المواد الدراسية"] = df_summary["المعلم/ة"].map(teacher_subjects)
                                df_summary.to_excel(writer, sheet_name="كشف_المعلمين", index=False)

                                wb = writer.book
                                ws_teachers = wb["كشف_المعلمين"]
                                ws_teachers.sheet_view.rightToLeft = True

                            for c in classes:
                                df_c = df_result[df_result["الفصل"].apply(lambda x: c in [i.strip() for i in str(x).split(",")])]
                                if not df_c.empty:
                                    df_c_copy = df_c.copy()
                                    df_c_copy["عرض_الخلايا"] = df_c_copy["المادة"] + "\n(" + df_c_copy["المدرس"] + ")"
                                    pivot_c = df_c_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first").fillna("فراغ").reindex(index=days, columns=periods)
                                    pivot_c.to_excel(writer, sheet_name=f"فصل_{c}")

                            for t in teachers:
                                df_t = df_result[df_result["المدرس"].apply(lambda x: t in [i.strip() for i in str(x).split("/")])]
                                if not df_t.empty:
                                    df_t_copy = df_t.copy()
                                    df_t_copy["عرض_الخلايا"] = df_t_copy["الفصل"] + "\n(" + df_t_copy["المادة"] + ")"
                                    pivot_t = df_t_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first")
                                else:
                                    pivot_t = pd.DataFrame("راحة", index=days, columns=periods)
                                pivot_t = pivot_t.reindex(index=days, columns=periods).fillna("راحة")
                                for off_day in teacher_off_days.get(t, []):
                                    if off_day in pivot_t.index:
                                        pivot_t.loc[off_day, :] = "إجازة (OFF)"
                                pivot_t.to_excel(writer, sheet_name=f"مدرس_{t}")

                            for r in rooms:
                                df_r = df_result[df_result["القاعة"] == r]
                                if not df_r.empty:
                                    df_r_copy = df_r.copy()
                                    df_r_copy["عرض_الخلايا"] = df_r_copy["الفصل"] + "\n(" + df_r_copy["المادة"] + " - " + df_r_copy["المدرس"] + ")"
                                    pivot_r = df_r_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first")
                                else:
                                    pivot_r = pd.DataFrame("متاحة", index=days, columns=periods)
                                pivot_r = pivot_r.reindex(index=days, columns=periods).fillna("متاحة")
                                pivot_r.to_excel(writer, sheet_name=f"قاعة_{r}")

                        format_excel_workbook(out_file, school_input_name)
                        st.session_state.generated = True
                    else:
                        st.error("❌ لم يتمكن البرنامج من توليد الجدول بسبب وجود قيود متضاربة في ملف الإدخال.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء المعالجة: {e}")

if st.session_state.generated:
    st.markdown("<br>", unsafe_allow_html=True)
    st.success("🎉 تم توليد الجداول وملف الحصص الشامل بنجاح واحترافية عالية!")
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        with open("final_timetable.xlsx", "rb") as f:
            st.download_button(
                label="📥 تحميل الجداول النهائية",
                data=f,
                file_name="final_timetable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    with col2:
        with open("all_classes_master_table.xlsx", "rb") as f:
            st.download_button(
                label="📥 تحميل ملف الحصص الشامل",
                data=f,
                file_name="all_classes_master_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

st.markdown("""
    <div class='footer'>
        Code Wonders Academy &nbsp;|&nbsp; ☎️ 01060572506
    </div>
""", unsafe_allow_html=True)
