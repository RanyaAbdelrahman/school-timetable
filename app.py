import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# إعدادات الصفحة
st.set_page_config(
    page_title="⭐ نظام الإدارة الذكية للجداول المدرسية ⭐",
    page_icon="🏫",
    layout="centered"
)

# تنسيقات CSS احترافية
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
    .stButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        font-weight: 700;
        border-radius: 14px;
        border: none;
    }
    </style>
""", unsafe_allow_html=True)

def clean_off_days(value):
    if pd.isna(value): return []
    text = str(value).strip()
    if not text: return []
    text = text.replace("،", ",")
    result = []
    for day in text.split(","):
        day = day.strip()
        if day and day not in result: result.append(day)
    return result

def format_excel_workbook(file_path, school_name):
    wb = load_workbook(file_path)
    header_fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    empty_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    off_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="E2E8F0")
    
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        if sheetname in ["Master_Schedule", "كشف_المعلمين"]:
            ws.sheet_view.rightToLeft = True
            continue
        
        if sheetname != "Master_Schedule" and sheetname != "كشف_المعلمين":
            ws.sheet_view.rightToLeft = True
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = center_align
                    cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                    if cell.value in ["فراغ", "راحة", "متاحة", "None", ""]:
                        cell.fill = empty_fill
                    elif "إجازة" in str(cell.value) or "OFF" in str(cell.value):
                        cell.fill = off_fill
    wb.save(file_path)

# واجهة المستخدم
st.markdown("<div class='main-header'><h1>⭐ نظام الإدارة الذكية للجداول المدرسية ⭐</h1><p>Code Wonders Academy</p></div>", unsafe_allow_html=True)
school_input_name = st.text_input("📝 اسم المدرسة", value="")
uploaded_file = st.file_uploader("📂 اختر ملف البيانات (inputs.xlsx)", type=["xlsx"])

if "generated" not in st.session_state: 
    st.session_state.generated = False

if uploaded_file is not None and st.button("🚀 إنشاء الجدول المدرسي"):
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

        teacher_off_days = {str(row["Teacher"]).strip(): [d for d in clean_off_days(row.get("OffDays", [])) if d in days] 
                           for _, row in df_teachers.iterrows() if pd.notna(row["Teacher"])}

        clean_assignments = []
        for idx, row in df_assignments.iterrows():
            clean_assignments.append({
                "idx": idx, "c": str(row["ClassName"]).strip(), "s": str(row["Subject"]).strip(),
                "t": str(row["Teacher"]).strip(), "r": str(row.get("PreferredRoom", "Classroom")).strip(),
                "w": int(row["WeeklyLessons"])
            })

        rooms = list(set([item["r"] for item in clean_assignments if item["r"] != "Classroom"]))
        model = cp_model.CpModel()
        schedule = {}

        for item in clean_assignments:
            for d in range(num_days):
                for p in range(num_periods):
                    schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] = model.NewBoolVar(f"var_{item['idx']}_{d}_{p}")

        # القيود (Constraints)
        for item in clean_assignments:
            model.Add(sum(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] for d in range(num_days) for p in range(num_periods)) == item["w"])
        
        for c in classes:
            for d in range(num_days):
                for p in range(num_periods):
                    rel_vars = [schedule[(i["idx"], i["c"], i["s"], i["t"], i["r"], d, p)] for i in clean_assignments if c in [x.strip() for x in i["c"].split(",")]]
                    if rel_vars: 
                        model.Add(sum(rel_vars) <= 1)
        
        for t_name in teachers:
            for d in range(num_days):
                for p in range(num_periods):
                    t_vars = [schedule[(i["idx"], i["c"], i["s"], i["t"], i["r"], d, p)] for i in clean_assignments if t_name in [x.strip() for x in i["t"].split("/")]]
                    if t_vars: 
                        model.Add(sum(t_vars) <= 1)

        # منع الحصص في أيام الإجازة الخاصة بالمعلمين
        for item in clean_assignments:
            for t_name in [x.strip() for x in item["t"].split("/")]:
                for d, day_name in enumerate(days):
                    if day_name in teacher_off_days.get(t_name, []):
                        for p in range(num_periods):
                            model.Add(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)] == 0)

        # دالة الهدف (Objective) لتوزيع الحصص بشكل تفضيلي
        model.Maximize(sum(var * (num_periods - p) for (key, var) in schedule.items() for p in [key[-1]]))
        
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30.0
        
        if solver.Solve(model) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            output_data = []
            for item in clean_assignments:
                for d in range(num_days):
                    for p in range(num_periods):
                        if solver.Value(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]) == 1:
                            output_data.append({"الفصل": item["c"], "المادة": item["s"], "المدرس": item["t"], "القاعة": item["r"], "اليوم": days[d], "الحصة": f"الحصة {p + 1}"})
            
            df_result = pd.DataFrame(output_data)
            
            # إنشاء الجدول الشامل (Master Table)
            wb_master = Workbook()
            ws_master = wb_master.active
            ws_master.title = "الحصص_الشامل"
            ws_master.sheet_view.rightToLeft = True
            ws_master.views.sheetView[0].showGridLines = True

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

            # تعبئة الجدول الشامل مع معالجة وتدقيق النصوص والفراغات بنسبة 100%
            for idx, cls in enumerate(sorted(list(classes)), start=1):
                row_num = 5 + idx - 1
                ws_master.row_dimensions[row_num].height = 35
                row_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid") if idx % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                ws_master.cell(row=row_num, column=1, value=idx).alignment = center_align
                ws_master.cell(row=row_num, column=2, value=str(cls)).alignment = center_align

                col_cursor = 3
                thin_border = Border(left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"), top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
                
                for day in days:
                    for p in periods:
                        match = df_result[(df_result["الفصل"].astype(str).str.strip() == str(cls).strip()) & 
                                          (df_result["اليوم"].astype(str).str.strip() == str(day).strip()) & 
                                          (df_result["الحصة"].astype(str).str.strip() == str(p).strip())]
                        
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

            wb_master.save("all_classes_master_table.xlsx")
            df_result.to_excel("final_timetable.xlsx", index=False)
            st.session_state.generated = True
    except Exception as e:
        st.error(f"خطأ أثناء المعالجة: {e}")

if st.session_state.generated:
    st.success("تم التوليد بنجاح!")
    col1, col2 = st.columns(2)
    with col1:
        with open("final_timetable.xlsx", "rb") as f: 
            st.download_button("تحميل الجداول النهائية", f, "final_timetable.xlsx")
    with col2:
        with open("all_classes_master_table.xlsx", "rb") as f: 
            st.download_button("تحميل الجدول الشامل", f, "all_classes_master_table.xlsx")
