import streamlit as st
from database import get_all_schools, update_school_license
import random
import string

def generate_license_key():
    part1 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    part2 = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    return f"CW-{part1}-{part2}"

def render_admin_panel():
    st.subheader("🛠️ لوحة تحكم المشرف (Admin Dashboard)")
    
    # حماية لوحة التحكم بكلمة مرور عبر secrets.toml
    admin_pass = st.text_input("أدخل كلمة مرور المشرف:", type="password")
    
    if "admin_logged_in" not in st.session_state:
        st.session_state["admin_logged_in"] = False

    if not st.session_state["admin_logged_in"]:
        if st.button("تسجيل دخول المشرف"):
            correct_pass = st.secrets.get("ADMIN_PASSWORD", "admin123")
            if admin_pass == correct_pass:
                st.session_state["admin_logged_in"] = True
                st.success("✅ تم تسجيل الدخول بنجاح!")
                st.rerun()
            else:
                st.error("❌ كلمة المرور غير صحيحة.")
        return

    st.success("أنت مسجل الدخول بصفتك المشرف الرئيسي.")
    
    if st.button("تسجيل خروج المشرف"):
        st.session_state["admin_logged_in"] = False
        st.rerun()

    st.markdown("---")
    st.subheader("📋 قائمة المدارس والاشتراكات")

    schools = get_all_schools()
    if not schools:
        st.info("لا توجد أي مدارس مسجلة حتى الآن.")
        return

    for school in schools:
        with st.expander(f"🏫 {school.get('school_name')} — الحالة: ({school.get('status').upper()})"):
            st.write(f"**البريد الإلكتروني:** {school.get('email')}")
            st.write(f"**الهاتف:** {school.get('phone', 'غير متوفر')}")
            st.write(f"**تاريخ التسجيل:** {school.get('created_at')}")
            st.write(f"**تاريخ بداية الاشتراك:** {school.get('start_date', 'غير محدد')}")
            st.write(f"**تاريخ انتهاء الاشتراك:** {school.get('expiry_date', 'غير محدد')}")
            st.write(f"**مفتاح الترخيص:** {school.get('license_key', 'لا يوجد')}")

            col1, col2 = st.columns(2)
            with col1:
                new_status = st.selectbox("تغيير الحالة:", ["pending", "approved", "rejected", "blocked"], index=["pending", "approved", "rejected", "blocked"].index(school.get("status", "pending")), key=f"status_{school.get('id')}")
            with col2:
                start_d = st.date_input("تاريخ البداية:", key=f"start_{school.get('id')}")
                expiry_d = st.date_input("تاريخ الانتهاء:", key=f"expiry_{school.get('id')}")

            if st.button(f"💾 حفظ التعديلات للمدرسة {school.get('school_name')}", key=f"save_{school.get('id')}"):
                lic_key = school.get('license_key') or generate_license_key()
                update_school_license(
                    school_id=school.get("id"),
                    status=new_status,
                    start_date=str(start_d),
                    expiry_date=str(expiry_d),
                    license_key=lic_key
                )
                st.success("✅ تم تحديث بيانات الترخيص بنجاح!")
                st.rerun()
