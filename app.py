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

# تنسيقات CSS احترافية بألوان مبهجة وعصرية
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%);
        font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif;
    }
    
    /* ترويسة رئيسية مبهجة ومميزة */
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

    /* حقول الإدخال */
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

    /* زر التوليد البارز والملون */
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

    /* تذييل الصفحة */
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
    school_title_font = Font(name="Segoe UI", size=15, bold=True, color="4F46E5")
    section_title_font = Font(name="Segoe UI", size=13, bold=True, color="1E293B")
    cell_font = Font(name="Segoe UI", size=10, bold=True, color="000000")
    day_font = Font(name="Segoe UI", size=11, bold=True, color="4F46E5")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="E2E8F0")
    med_side = Side(style="medium", color="6366F1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

        if sheetname in ["Master_Schedule", "كشف_المعلمين"]:
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.sheet_view.rightToLeft = True
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            continue

        if sheetname == "جداول_جميع_الفصول":
            ws.sheet_view.rightToLeft = True
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            continue

        df_sheet = pd.read_excel(file_path, sheet_name=sheetname)
        num_cols = len(df_sheet.columns)
        sheet_type, title_val = (sheetname.split("_", 1) if "_" in sheetname else ("", sheetname))
        
        if sheet_type == "فصل": sub_title = f"جدول حصص فصل: {title_val}"
        elif sheet_type == "مدرس": sub_title = f"جدول حصص المعلم/ة: {title_val}"
        elif sheet_type == "قاعة": sub_title = f"جدول إشغال قاعة / نشاط: {title_val}"
        else: sub_title = title_val

        ws.insert_rows(1, amount=2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        ws.cell(row=1, column=1, value=school_name).font = school_title_font
        ws.cell(row=1, column=1).alignment = center_align
        ws.row_dimensions[1].height = 25

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        ws.cell(row=2, column=1, value=sub_title).font = section_title_font
        ws.cell(row=2, column=1).alignment = center_align
        ws.row_dimensions[2].height = 22

        header_row_idx = 3
        ws.row_dimensions[header_row_idx].height = 26
        for col_idx in range(1, num_cols + 1):
            c = ws.cell(row=header_row_idx, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = center_align
            c.border = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

        max_row = ws.max_row
        for r_idx in range(4, max_row + 1):
            ws.row_dimensions[r_idx].height = 42
            for c_idx in range(1, num_cols + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = cell_font
                cell.alignment = center_align
                cell.border = cell_border
                cell.fill = data_fill
                if c_idx == 1:
                    cell.font = day_font
                    cell.fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
                val_str = str(cell.value or "")
                if val_str in ["فراغ", "راحة", "متاحة", "None", ""]:
                    cell.fill = empty_fill
                    cell.font = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
                elif "إجازة" in val_str or "OFF" in val_str:
                    cell.fill = off_fill
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")

        for col in ws.columns:
            max_len = max(len(str(line)) for cell in col for line in str(cell.value or "").split("\n"))
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 5, 18)
    wb.save(file_path)

# الترويسة
st.markdown("""
    <div class="main-header">
        <h1> ⭐ نظام الإدارة الذكية للجداول المدرسية ⭐ </h1>
        <p> Code Wonders Academy </p>
    </div>
""", unsafe_allow_html=True)

school_input_name = st.text_input("📝 اسم المدرسة", value="")
uploaded_file = st.file_uploader("📂 اختر ملف البيانات بصيغة Excel (inputs.xlsx)", type=["xlsx"])

# استخدام session_state لحفظ حالة النجاح لكي لا تختفي عند الضغط على التحميل
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
                    for _, row in df_teachers.iterrows():
                        if pd.isna(row["Teacher"]): continue
                        t_name = str(row["Teacher"]).strip()
                        off_days = clean_off_days(row.get("OffDays", []))
                        teacher_off_days[t_name] = [d for d in off_days if d in days]

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

                    # ========================================================
                    # ⭐ OBJECTIVE - تحسين جودة الجدول
                    # ========================================================

                    objective_terms = []

                    # 1. عدالة توزيع كل مادة على أيام الأسبوع داخل الفصل
                    SUBJECT_DISTRIBUTION_PENALTY = 40
                    class_subject_groups = {}

                    for item in clean_assignments:
                        class_names = [x.strip() for x in item["c"].split(",") if x.strip()]
                        for class_name in class_names:
                            key = (class_name, item["s"])
                            if key not in class_subject_groups:
                                class_subject_groups[key] = []
                            class_subject_groups[key].append(item)

                    for (class_name, subject), items in class_subject_groups.items():
                        daily_load = {}
                        for d in range(num_days):
                            lesson_vars = []
                            for item in items:
                                item_classes = [x.strip() for x in item["c"].split(",") if x.strip()]
                                if class_name not in item_classes:
                                    continue
                                for p in range(num_periods):
                                    lesson_vars.append(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)])

                            daily_load[d] = model.NewIntVar(0, len(lesson_vars), f"subject_load_{class_name}_{subject}_{d}")
                            if lesson_vars:
                                model.Add(daily_load[d] == sum(lesson_vars))
                            else:
                                model.Add(daily_load[d] == 0)

                        for d1 in range(num_days):
                            for d2 in range(d1 + 1, num_days):
                                difference = model.NewIntVar(0, num_periods * len(items), f"subject_diff_{class_name}_{subject}_{d1}_{d2}")
                                model.AddAbsEquality(difference, daily_load[d1] - daily_load[d2])
                                objective_terms.append(difference * (-SUBJECT_DISTRIBUTION_PENALTY))

                    # 2. تفضيل الحصص المبكرة
                    EARLY_PERIOD_WEIGHT = 10
                    for key, var in schedule.items():
                        p = key[-1]
                        weight = (num_periods - p) * EARLY_PERIOD_WEIGHT
                        objective_terms.append(var * weight)

                    # 3. تقليل الحصص المتتالية لنفس المادة/Assignment
                    CONSECUTIVE_PENALTY = 5
                    for item in clean_assignments:
                        idx = item["idx"]
                        c = item["c"]
                        s = item["s"]
                        t = item["t"]
                        r = item["r"]

                        for d in range(num_days):
                            for p in range(num_periods - 1):
                                current_var = schedule[(idx, c, s, t, r, d, p)]
                                next_var = schedule[(idx, c, s, t, r, d, p + 1)]

                                both_lessons = model.NewBoolVar(f"both_{idx}_{d}_{p}")
                                model.Add(both_lessons <= current_var)
                                model.Add(both_lessons <= next_var)
                                model.Add(both_lessons >= current_var + next_var - 1)

                                objective_terms.append(both_lessons * (-CONSECUTIVE_PENALTY))

                    # الهدف النهائي
                    model.Maximize(sum(objective_terms))

                    solver = cp_model.CpSolver()
                    solver.parameters.max_time_in_seconds = 60.0
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

# عرض رسالة النجاح وأزرار التحميل بشكل ثابت دائماً إذا تم التوليد بنجاح
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

# التذييل
st.markdown("""
    <div class='footer'>
        Code Wonders Academy &nbsp;|&nbsp; ☎️ 01060572506
    </div>
""", unsafe_allow_html=True)
