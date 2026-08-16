import streamlit as st
from supabase import create_client, Client
from datetime import datetime, date

@st.cache_resource
def get_supabase_client() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def get_school_by_email(email: str):
    supabase = get_supabase_client()
    res = supabase.table("Schools").select("*").eq("email", email.strip().lower()).execute()
    if res.data:
        return res.data[0]
    return None

def register_new_school(school_name: str, email: str, phone: str = ""):
    supabase = get_supabase_client()
    data = {
        "school_name": school_name.strip(),
        "email": email.strip().lower(),
        "phone": phone.strip(),
        "status": "pending"
    }
    res = supabase.table("Schools").insert(data).execute()
    return res.data

def get_all_schools():
    supabase = get_supabase_client()
    res = supabase.table("Schools").select("*").order("created_at", desc=True).execute()
    return res.data or []

def update_school_license(school_id: int, status: str, start_date: str = None, expiry_date: str = None, license_key: str = None):
    supabase = get_supabase_client()
    data = {"status": status}
    if start_date: data["start_date"] = str(start_date)
    if expiry_date: data["expiry_date"] = str(expiry_date)
    if license_key: data["license_key"] = license_key
    
    res = supabase.table("Schools").update(data).eq("id", school_id).execute()
    return res.data

def check_license_validity(school_data):
    if not school_data:
        return False, "❌ البريد الإلكتروني غير مسجل."
    
    status = school_data.get("status")
    if status == "pending":
        return False, "⏳ طلب تسجيل المدرسة قيد المراجعة والموافقة من الإدارة."
    elif status == "rejected" or status == "blocked":
        return False, "🚫 هذا الحساب معطل أو تم رفض الطلب."
    
    expiry_str = school_data.get("expiry_date")
    if not expiry_str:
        return False, "⚠️ لا يوجد ترخيص فعال مرتبط بهذا الحساب."
    
    try:
        expiry_date = datetime.strptime(str(expiry_str), "%Y-%m-%d").date()
        today = date.today()
        if today > expiry_date:
            return False, f"🔒 انتهت صلاحية اشتراك المدرسة في تاريخ ({expiry_str}). يرجى التواصل مع الدعم للتجديد."
    except Exception:
        pass
        
    return True, "✅ الاشتراك ساري ومفعل."
