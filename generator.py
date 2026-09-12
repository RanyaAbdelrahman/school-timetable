import re
import os
import sys
import math
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model
import pandas as pd

# ============================================================
# تحديد المسار
# ============================================================

def get_path(filename):
    workdir = os.environ.get("TIMETABLE_WORKDIR")
    if workdir:
        os.makedirs(workdir, exist_ok=True)
        return os.path.join(workdir, filename)
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

SCHOOL_NAME = "مدرسة --------"

def get_school_name():
    return str(os.environ.get("SCHOOL_NAME", SCHOOL_NAME)).strip() or "مدرسة --------"

# ============================================================
# تنظيف أيام الإجازة
# ============================================================

def clean_off_days(value):
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    text = text.replace("،", ",")
    result = []
    for day in text.split(","):
        day = day.strip()
        if day and day not in result:
            result.append(day)
    return result

def clean_unwanted_periods(value):
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    text = text.replace("،", ",")
    result = []
    for p in text.split(","):
        p = p.strip()
        if p.isdigit():
            val = int(p)
            if val not in result:
                result.append(val)
    return result

def clean_supervision_days(value):
    """تنظيف أيام الإشراف المدخلة في عمود الإشراف."""
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    text = text.replace("،", ",")
    result = []
    for day in text.split(","):
        day = day.strip()
        if day and day not in result:
            result.append(day)
    return result

# ============================================================
# تنسيق ملف Excel
# ============================================================

def format_excel_workbook(file_path, school_name=None):
    wb = load_workbook(file_path)
    school_name = str(school_name if school_name is not None else get_school_name()).strip() or 'مدرسة --------'

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    empty_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    off_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    school_title_font = Font(name="Segoe UI", size=15, bold=True, color="1F4E78")
    section_title_font = Font(name="Segoe UI", size=13, bold=True, color="000000")
    cell_font = Font(name="Segoe UI", size=10, bold=True, color="000000")
    day_font = Font(name="Segoe UI", size=11, bold=True, color="1F4E78")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="BFBFBF")
    med_side = Side(style="medium", color="1F4E78")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True

        if sheetname in ["Master_Schedule", "كشف_المعلمين", "قائمة_الفصول"]:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            continue

        df_sheet = pd.read_excel(file_path, sheet_name=sheetname)
        num_cols = len(df_sheet.columns)

        if "_" in sheetname:
            sheet_type, title_val = sheetname.split("_", 1)
        else:
            sheet_type = ""
            title_val = sheetname

        if sheet_type == "فصل":
            sub_title = f"جدول حصص فصل: {title_val}"
        elif sheet_type == "مدرس":
            sub_title = f"جدول حصص المعلم/ة: {title_val}"
        elif sheet_type == "قاعة":
            sub_title = f"جدول إشغال قاعة / نشاط: {title_val}"
        else:
            sub_title = title_val

        ws.insert_rows(1, amount=2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        ws.cell(row=1, column=1, value=school_name)
        ws.cell(row=1, column=1).font = school_title_font
        ws.cell(row=1, column=1).alignment = center_align
        ws.row_dimensions[1].height = 25

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        ws.cell(row=2, column=1, value=sub_title)
        ws.cell(row=2, column=1).font = section_title_font
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
            ws.row_dimensions[r_idx].height = 36
            for c_idx in range(1, num_cols + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = cell_font
                cell.alignment = center_align
                cell.border = cell_border
                cell.fill = data_fill

                if c_idx == 1:
                    cell.font = day_font
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                    val_text = str(cell.value or "")
                    if any("\u0600" <= char <= "\u06ff" for char in val_text):
                        cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                val_str = str(cell.value or "")
                if val_str in ["فراغ", "راحة", "متاحة", "None", ""]:
                    cell.fill = empty_fill
                    cell.font = Font(name="Segoe UI", size=10, italic=True, color="7F7F7F")
                elif "إجازة" in val_str or "OFF" in val_str:
                    cell.fill = off_fill
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                for line in val.split("\n"):
                    if len(line) > max_len:
                        max_len = len(line)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    wb.save(file_path)

# ============================================================
# توليد الجدول
# ============================================================

def generate_timetable():
    excel_file = get_path("inputs.xlsx")

    if not os.path.exists(excel_file):
        print("❌ لم يتم العثور على ملف inputs.xlsx")
        print(f"المسار: {excel_file}")
        return

    try:
        df_teachers = pd.read_excel(excel_file, sheet_name="Teachers")
        df_classes = pd.read_excel(excel_file, sheet_name="Classes")
        df_assignments = pd.read_excel(excel_file, sheet_name="Assignments")
        df_settings = pd.read_excel(excel_file, sheet_name="Settings")
        df_days = pd.read_excel(excel_file, sheet_name="Days")
    except Exception as e:
        print("❌ حدث خطأ أثناء قراءة inputs.xlsx")
        print(e)
        return

    days = [str(d).strip() for d in df_days["DayName"].dropna().tolist()]
    num_days = len(days)
    num_periods = int(df_settings["PeriodsPerDay"].iloc[0])
    classes = [str(c).strip() for c in df_classes["ClassName"].dropna().tolist()]
    teachers = [str(t).strip() for t in df_teachers["Teacher"].dropna().tolist()]
    periods = [f"الحصة {p + 1}" for p in range(num_periods)]

    teacher_off_days = {}
    teacher_max_off_days = {}
    teacher_unwanted_periods = {}
    teacher_supervision_days = {}

    for _, row in df_teachers.iterrows():
        if pd.isna(row["Teacher"]):
            continue

        teacher_name = str(row["Teacher"]).strip()

        if "OffDays" in df_teachers.columns:
            off_days = clean_off_days(row["OffDays"])
        else:
            off_days = []

        valid_off_days = [d for d in off_days if d in days]
        teacher_off_days[teacher_name] = valid_off_days

        if "UnwantedPeriods" in df_teachers.columns:
            teacher_unwanted_periods[teacher_name] = clean_unwanted_periods(row["UnwantedPeriods"])
        else:
            teacher_unwanted_periods[teacher_name] = []

        if "الإشراف" in df_teachers.columns:
            supervision_days = clean_supervision_days(row["الإشراف"])
        else:
            supervision_days = []

        valid_supervision_days = [d for d in supervision_days if d in days]
        teacher_supervision_days[teacher_name] = valid_supervision_days

        if "MaxOffDays" in df_teachers.columns and pd.notna(row.get("MaxOffDays")):
            try:
                teacher_max_off_days[teacher_name] = int(row["MaxOffDays"])
            except (ValueError, TypeError):
                teacher_max_off_days[teacher_name] = len(valid_off_days)
        else:
            teacher_max_off_days[teacher_name] = len(valid_off_days)

    clean_assignments = []
    for idx, row in df_assignments.iterrows():
        c = str(row["ClassName"]).strip()
        s = str(row["Subject"]).strip()
        t = str(row["Teacher"]).strip()

        try:
            w = int(row["WeeklyLessons"])
        except (ValueError, TypeError):
            print(f"❌ WeeklyLessons غير صحيح في الصف {idx + 2}")
            return

        r_val = row.get("PreferredRoom", "Classroom")
        r = "Classroom" if pd.isna(r_val) or str(r_val).strip() in ["nan", "None", ""] else str(r_val).strip()

        clean_assignments.append({"idx": idx, "c": c, "s": s, "t": t, "r": r, "w": w})

    model = cp_model.CpModel()
    schedule = {}

    for item in clean_assignments:
        idx, c, s, t, r = item["idx"], item["c"], item["s"], item["t"], item["r"]
        for d in range(num_days):
            for p in range(num_periods):
                schedule[(idx, c, s, t, r, d, p)] = model.NewBoolVar(f"var_{idx}_{d}_{p}")

    for item in clean_assignments:
        idx, c, s, t, r, w = item["idx"], item["c"], item["s"], item["t"], item["r"], item["w"]
        model.Add(sum(schedule[(idx, c, s, t, r, d, p)] for d in range(num_days) for p in range(num_periods)) == w)

    for c in classes:
        for d in range(num_days):
            for p in range(num_periods):
                relevant_vars = [
                    schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                    for item in clean_assignments if c in [x.strip() for x in item["c"].split(",")]
                ]
                if relevant_vars:
                    model.Add(sum(relevant_vars) <= 1)

    for teacher_name in teachers:
        for d in range(num_days):
            for p in range(num_periods):
                teacher_vars = [
                    schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                    for item in clean_assignments
                    if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]
                ]
                if teacher_vars:
                    model.Add(sum(teacher_vars) <= 1)

    for item in clean_assignments:
        idx, c, s, t, r = item["idx"], item["c"], item["s"], item["t"], item["r"]
        assignment_teachers = [x.strip() for x in t.split("/") if x.strip()]
        for teacher_name in assignment_teachers:
            off_days = teacher_off_days.get(teacher_name, [])
            for d in range(num_days):
                if days[d] in off_days:
                    for p in range(num_periods):
                        model.Add(schedule[(idx, c, s, t, r, d, p)] == 0)

    all_rooms = set([item["r"] for item in clean_assignments if item["r"] not in ["Classroom", "nan", ""]])
    for room in all_rooms:
        for d in range(num_days):
            for p in range(num_periods):
                room_vars = [
                    schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                    for item in clean_assignments if item["r"] == room
                ]
                if room_vars:
                    model.Add(sum(room_vars) <= 1)

    target_period_idx = 6
    if num_periods > target_period_idx:
        for item in clean_assignments:
            subject = item["s"]
            if "PE" in subject.upper() or "تربيه رياضيه" in subject or "تربية رياضية" in subject:
                for d in range(num_days):
                    model.Add(schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, target_period_idx)] == 0)

    # قيد حصص القرآن الكريم - قيد صارم
    quran_by_class = {}
    def is_quran_subject(subject):
        text = str(subject or '').strip().lower()
        return 'قرآن' in text or 'قران' in text or 'quran' in text

    for item in clean_assignments:
        if not is_quran_subject(item['s']):
            continue
        item_classes = [x.strip() for x in str(item['c']).split(',') if x.strip()]
        for class_name in item_classes:
            quran_by_class.setdefault(class_name, []).append(item)

    for class_name, quran_items in quran_by_class.items():
        total_quran_lessons = sum(int(item['w']) for item in quran_items)
        if total_quran_lessons <= 0 or num_days <= 0:
            continue
        max_quran_per_day = min(math.ceil(total_quran_lessons / num_days), num_periods)
        quran_period_vars = {}

        for d in range(num_days):
            day_vars = []
            for p in range(num_periods):
                period_vars = []
                for item in quran_items:
                    item_classes = [x.strip() for x in str(item['c']).split(',') if x.strip()]
                    if class_name not in item_classes:
                        continue
                    period_vars.append(schedule[(item['idx'], item['c'], item['s'], item['t'], item['r'], d, p)])

                qvar = model.NewBoolVar(f'quran_{class_name}_{d}_{p}')
                if period_vars:
                    model.Add(qvar == sum(period_vars))
                else:
                    model.Add(qvar == 0)
                quran_period_vars[(d, p)] = qvar
                day_vars.append(qvar)

            mode_vars = []
            no_quran = model.NewBoolVar(f'quran_no_{class_name}_{d}')
            mode_vars.append(no_quran)

            start_modes, end_modes = {}, {}
            for k in range(1, max_quran_per_day + 1):
                start_var = model.NewBoolVar(f'quran_start_{class_name}_{d}_{k}')
                end_var = model.NewBoolVar(f'quran_end_{class_name}_{d}_{k}')
                start_modes[k] = start_var
                end_modes[k] = end_var
                mode_vars.extend([start_var, end_var])

            model.Add(sum(mode_vars) == 1)

            for p in range(num_periods):
                allowed_mode_vars = []
                for k in range(1, max_quran_per_day + 1):
                    if p < k:
                        allowed_mode_vars.append(start_modes[k])
                    if p >= num_periods - k:
                        allowed_mode_vars.append(end_modes[k])
                if allowed_mode_vars:
                    model.Add(day_vars[p] == sum(allowed_mode_vars))
                else:
                    model.Add(day_vars[p] == 0)

            model.Add(sum(day_vars) <= max_quran_per_day)

        all_quran_vars = [quran_period_vars[(d, p)] for d in range(num_days) for p in range(num_periods)]
        model.Add(sum(all_quran_vars) == total_quran_lessons)

    last_period = num_periods - 1
    for teacher_name in teachers:
        supervision_days = teacher_supervision_days.get(teacher_name, [])
        if not supervision_days:
            continue
        teacher_items = [item for item in clean_assignments if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]]
        for supervision_day in supervision_days:
            d = days.index(supervision_day)
            last_period_vars = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, last_period)]
                for item in teacher_items
                if (item["idx"], item["c"], item["s"], item["t"], item["r"], d, last_period) in schedule
            ]
            if last_period_vars:
                model.Add(sum(last_period_vars) >= 1)
            else:
                model.AddBoolOr([])

    late_periods = [p for p in range(num_periods) if p >= num_periods - 2]
    for item in clean_assignments:
        idx, c, s, t, r, w = item["idx"], item["c"], item["s"], item["t"], item["r"], item["w"]
        max_allowed_late = max(1, (w + 1) // 2)
        late_vars = [schedule[(idx, c, s, t, r, d, p)] for d in range(num_days) for p in late_periods]
        if late_vars:
            model.Add(sum(late_vars) <= max_allowed_late)

    objective_terms = []
    UNWANTED_PERIOD_PENALTY = 40
    TEACHER_ACTIVE_DAY_BONUS = 25000
    TEACHER_LOAD_BALANCE_PENALTY = 100000
    TEACHER_SUBJECT_DAY_BALANCE_PENALTY = 15000
    TEACHER_PERIOD_DIVERSITY_BONUS = 5000
    TEACHER_PERIOD_BALANCE_PENALTY = 5000
    CLASS_SUBJECT_DAY_BALANCE_PENALTY = 30000

    for (idx, c, s, t, r, d, p), var in schedule.items():
        weight = (num_periods - p) * 10
        objective_terms.append(var * weight)

    for item in clean_assignments:
        idx, c, s, t, r = item["idx"], item["c"], item["s"], item["t"], item["r"]
        assignment_teachers = [x.strip() for x in t.split("/") if x.strip()]
        for d in range(num_days):
            is_off_day = any(days[d] in teacher_off_days.get(tn, []) for tn in assignment_teachers)
            if not is_off_day:
                day_has_subject = model.NewBoolVar(f"day_has_{idx}_{d}")
                day_lessons = [schedule[(idx, c, s, t, r, d, p)] for p in range(num_periods)]
                model.Add(sum(day_lessons) >= 1).OnlyEnforceIf(day_has_subject)
                model.Add(sum(day_lessons) == 0).OnlyEnforceIf(day_has_subject.Not())
                objective_terms.append(day_has_subject * 80)

    for teacher_name in teachers:
        off_days = teacher_off_days.get(teacher_name, [])
        work_days_indices = [d for d in range(num_days) if days[d] not in off_days]
        if not work_days_indices:
            continue

        teacher_day_has_lessons = {}
        for d in work_days_indices:
            teacher_day_has_lessons[d] = model.NewBoolVar(f"teacher_active_{teacher_name}_{d}")
            teacher_lessons_on_day = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                for item in clean_assignments
                if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]
                for p in range(num_periods)
            ]
            if teacher_lessons_on_day:
                model.Add(sum(teacher_lessons_on_day) >= 1).OnlyEnforceIf(teacher_day_has_lessons[d])
                model.Add(sum(teacher_lessons_on_day) == 0).OnlyEnforceIf(teacher_day_has_lessons[d].Not())
                objective_terms.append(teacher_day_has_lessons[d] * TEACHER_ACTIVE_DAY_BONUS)

        teacher_total_lessons = sum(
            item["w"] for item in clean_assignments
            if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]
        )
        if teacher_total_lessons >= len(work_days_indices) and work_days_indices:
            for d in work_days_indices:
                model.Add(teacher_day_has_lessons[d] == 1)

    teacher_subject_groups = {}
    for item in clean_assignments:
        for teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]:
            teacher_subject_groups.setdefault((teacher_name, item["s"]), []).append(item)

    for (teacher_name, subject), items in teacher_subject_groups.items():
        off_days = teacher_off_days.get(teacher_name, [])
        work_days_indices = [d for d in range(num_days) if days[d] not in off_days]
        if len(work_days_indices) <= 1:
            continue

        subject_daily_loads = {}
        for d in work_days_indices:
            day_vars = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                for item in items for p in range(num_periods)
            ]
            subject_daily_loads[d] = model.NewIntVar(0, num_periods * max(1, len(items)), f"t_sub_load_{teacher_name}_{subject}_{d}")
            if day_vars:
                model.Add(subject_daily_loads[d] == sum(day_vars))
            else:
                model.Add(subject_daily_loads[d] == 0)

        for i in range(len(work_days_indices)):
            for j in range(i + 1, len(work_days_indices)):
                d1, d2 = work_days_indices[i], work_days_indices[j]
                diff = model.NewIntVar(0, num_periods * max(1, len(items)), f"t_sub_diff_{teacher_name}_{subject}_{d1}_{d2}")
                model.AddAbsEquality(diff, subject_daily_loads[d1] - subject_daily_loads[d2])
                objective_terms.append(diff * (-TEACHER_SUBJECT_DAY_BALANCE_PENALTY))

    class_subject_groups = {}
    for item in clean_assignments:
        for class_name in [x.strip() for x in str(item["c"]).split(",") if x.strip()]:
            class_subject_groups.setdefault((class_name, item["s"]), []).append(item)

    for (class_name, subject), items in class_subject_groups.items():
        if num_days <= 1:
            continue
        daily_loads = {}
        for d in range(num_days):
            day_vars = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                for item in items for p in range(num_periods)
            ]
            daily_loads[d] = model.NewIntVar(0, num_periods * max(1, len(items)), f"c_sub_load_{class_name}_{subject}_{d}")
            model.Add(daily_loads[d] == sum(day_vars) if day_vars else 0)

        for i in range(num_days):
            for j in range(i + 1, num_days):
                diff = model.NewIntVar(0, num_periods * max(1, len(items)), f"c_sub_diff_{class_name}_{subject}_{i}_{j}")
                model.AddAbsEquality(diff, daily_loads[i] - daily_loads[j])
                objective_terms.append(diff * (-CLASS_SUBJECT_DAY_BALANCE_PENALTY))

    for teacher_name in teachers:
        off_days = teacher_off_days.get(teacher_name, [])
        work_days_indices = [d for d in range(num_days) if days[d] not in off_days]
        if not work_days_indices:
            continue

        teacher_period_loads = {}
        teacher_period_used = {}
        for p in range(num_periods):
            period_vars = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                for item in clean_assignments
                if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]
                for d in work_days_indices
            ]
            teacher_period_loads[p] = model.NewIntVar(0, num_days * max(1, len(clean_assignments)), f"t_p_load_{teacher_name}_{p}")
            if period_vars:
                model.Add(teacher_period_loads[p] == sum(period_vars))
            else:
                model.Add(teacher_period_loads[p] == 0)

            teacher_period_used[p] = model.NewBoolVar(f"t_p_used_{teacher_name}_{p}")
            model.Add(teacher_period_loads[p] >= 1).OnlyEnforceIf(teacher_period_used[p])
            model.Add(teacher_period_loads[p] == 0).OnlyEnforceIf(teacher_period_used[p].Not())
            objective_terms.append(teacher_period_used[p] * TEACHER_PERIOD_DIVERSITY_BONUS)

        for p1 in range(num_periods):
            for p2 in range(p1 + 1, num_periods):
                diff = model.NewIntVar(0, num_days * max(1, len(clean_assignments)), f"t_p_diff_{teacher_name}_{p1}_{p2}")
                model.AddAbsEquality(diff, teacher_period_loads[p1] - teacher_period_loads[p2])
                objective_terms.append(diff * (-TEACHER_PERIOD_BALANCE_PENALTY))

        first_last_diff = model.NewIntVar(0, num_days * max(1, len(clean_assignments)), f"t_fl_diff_{teacher_name}")
        model.AddAbsEquality(first_last_diff, teacher_period_loads[0] - teacher_period_loads[num_periods - 1])
        objective_terms.append(first_last_diff * (-10000))

    for item in clean_assignments:
        idx, c, s, t, r = item["idx"], item["c"], item["s"], item["t"], item["r"]
        assignment_teachers = [x.strip() for x in t.split("/") if x.strip()]
        for teacher_name in assignment_teachers:
            unwanted_ps = teacher_unwanted_periods.get(teacher_name, [])
            if unwanted_ps:
                for d in range(num_days):
                    for p in range(num_periods):
                        if (p + 1) in unwanted_ps:
                            objective_terms.append(schedule[(idx, c, s, t, r, d, p)] * (-UNWANTED_PERIOD_PENALTY))

    TEACHER_MAX_DAILY_PREFERRED = 4
    TEACHER_OVERLOAD_PENALTY = 4000

    for teacher_name in teachers:
        off_days = teacher_off_days.get(teacher_name, [])
        work_days_indices = [d for d in range(num_days) if days[d] not in off_days]
        if len(work_days_indices) <= 1:
            continue

        teacher_daily_loads = {}
        for d in work_days_indices:
            day_lessons = [
                schedule[(item["idx"], item["c"], item["s"], item["t"], item["r"], d, p)]
                for item in clean_assignments
                if teacher_name in [x.strip() for x in item["t"].split("/") if x.strip()]
                for p in range(num_periods)
            ]
            teacher_daily_loads[d] = model.NewIntVar(0, num_periods, f"t_load_{teacher_name}_{d}")
            if day_lessons:
                model.Add(teacher_daily_loads[d] == sum(day_lessons))
            else:
                model.Add(teacher_daily_loads[d] == 0)

            overload = model.NewIntVar(0, num_periods, f"t_over_{teacher_name}_{d}")
            model.Add(overload >= teacher_daily_loads[d] - TEACHER_MAX_DAILY_PREFERRED)
            model.Add(overload >= 0)
            objective_terms.append(overload * (-TEACHER_OVERLOAD_PENALTY))

        for i in range(len(work_days_indices)):
            for j in range(i + 1, len(work_days_indices)):
                d1, d2 = work_days_indices[i], work_days_indices[j]
                diff_var = model.NewIntVar(0, num_periods, f"t_diff_{teacher_name}_{d1}_{d2}")
                model.AddAbsEquality(diff_var, teacher_daily_loads[d1] - teacher_daily_loads[d2])
                objective_terms.append(diff_var * (-TEACHER_LOAD_BALANCE_PENALTY))

    model.Maximize(sum(objective_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print()
        print("=" * 60)
        print("✅ تم توليد الجدول بنجاح!")
        print("=" * 60)

        output_data = []
        for item in clean_assignments:
            idx, c, s, t, r = item["idx"], item["c"], item["s"], item["t"], item["r"]
            for d in range(num_days):
                for p in range(num_periods):
                    if solver.Value(schedule[(idx, c, s, t, r, d, p)]) == 1:
                        output_data.append({
                            "الفصل": c,
                            "المادة": s,
                            "المدرس": t,
                            "القاعة": r,
                            "اليوم": days[d],
                            "الحصة": f"الحصة {p + 1}",
                        })

        df_result = pd.DataFrame(output_data)

        def safe_filename(value):
            value = re.sub(r'[\\/:*?"<>|]+', "_", str(value).strip())
            value = re.sub(r"\s+", " ", value).strip(" .")
            return value or "مدرسة"

        school_name = get_school_name()
        safe_school_name = safe_filename(school_name)

        out_file = get_path(f"{safe_school_name}_final_timetable.xlsx")
        master_table_file = get_path(f"{safe_school_name}_all_classes.xlsx")

        rooms = sorted({
            str(item.get("r", "")).strip()
            for item in clean_assignments
            if str(item.get("r", "")).strip() and str(item.get("r", "")).strip().lower() not in {"classroom", "nan"}
        })

        wb_master = Workbook()
        ws_master = wb_master.active
        ws_master.title = "الحصص_الشامل"
        ws_master.sheet_view.rightToLeft = True

        thin_border = Border(
            left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
        )
        header_fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
        sub_header_fill = PatternFill(start_color="818CF8", end_color="818CF8", fill_type="solid")
        title_fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        total_cols = 2 + len(days) * len(periods)
        ws_master.merge_cells(f"A1:{get_column_letter(total_cols)}1")
        title = ws_master.cell(1, 1, f"{school_name} - جدول الحصص المدرسي الشامل لجميع الفصول")
        title.font = Font(name="Segoe UI", size=14, bold=True, color="4F46E5")
        title.alignment = center
        title.fill = title_fill
        ws_master.row_dimensions[1].height = 40

        for col, value in ((1, "م"), (2, "اسم الفصل")):
            cell = ws_master.cell(3, col, value)
            cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            cell.alignment = center
            cell.fill = header_fill

        ws_master.merge_cells("A3:A4")
        ws_master.merge_cells("B3:B4")

        current_col = 3
        for day in days:
            start_col = current_col
            end_col = current_col + len(periods) - 1
            ws_master.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
            day_cell = ws_master.cell(3, start_col, day)
            day_cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            day_cell.alignment = center
            day_cell.fill = header_fill

            for p_idx, period in enumerate(periods):
                cell = ws_master.cell(4, start_col + p_idx, period)
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                cell.alignment = center
                cell.fill = sub_header_fill

            current_col += len(periods)

        for idx, cls in enumerate(sorted(classes), start=1):
            row_num = 4 + idx
            ws_master.cell(row_num, 1, idx).alignment = center
            ws_master.cell(row_num, 2, str(cls)).alignment = center
            ws_master.row_dimensions[row_num].height = 35

            col_cursor = 3
            for day in days:
                for period in periods:
                    match = df_result[
                        df_result["الفصل"].apply(lambda x: str(cls) in [part.strip() for part in str(x).split(",")])
                        & (df_result["اليوم"] == day)
                        & (df_result["الحصة"] == period)
                    ]

                    if not match.empty:
                        first = match.iloc[0]
                        subject = str(first.get("المادة", ""))
                        teacher = str(first.get("المدرس", ""))
                        value = f"{subject}\n({teacher})" if teacher else subject
                    else:
                        value = "متاحة"

                    cell = ws_master.cell(row_num, col_cursor, value)
                    cell.alignment = center
                    cell.border = thin_border
                    cell.font = Font(name="Segoe UI", size=8)
                    col_cursor += 1

        for column in ws_master.columns:
            max_len = max((len(str(c.value or "").split("\n")[0]) for c in column), default=10)
            ws_master.column_dimensions[get_column_letter(column[0].column)].width = max(max_len + 4, 15)

        ws_teachers = wb_master.create_sheet("جداول_المعلمين")
        ws_teachers.sheet_view.rightToLeft = True
        ws_teachers.sheet_view.showGridLines = False

        class_palette = [
            "FFF2CC", "D9EAD3", "CFE2F3", "F4CCCC", "D9D2E9",
            "FCE5CD", "D0E0E3", "EAD1DC", "DDEBF7", "E2F0D9",
            "FFF2F2", "EDE7F6", "FFF4CC", "DDEBF7", "FCE4D6",
            "E2EFDA", "F4CCCC", "D9E1F2", "E4DFEC", "FCE5CD"
        ]
        class_colors = {str(cls).strip(): class_palette[idx % len(class_palette)] for idx, cls in enumerate(sorted(classes, key=lambda x: str(x)))}

        all_teachers = set()
        for teacher_cell in df_result.get("المدرس", pd.Series(dtype=str)).astype(str):
            for teacher in teacher_cell.split("/"):
                teacher = teacher.strip()
                if teacher and teacher.lower() != "nan":
                    all_teachers.add(teacher)
        all_teachers = sorted(all_teachers, key=lambda x: str(x))

        fixed_cols = 3
        total_cols = fixed_cols + len(days) * len(periods)
        ws_teachers.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        title_cell = ws_teachers.cell(1, 1, f"{school_name} - الجدول المدرسي الشامل للمعلمين")
        title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
        title_cell.alignment = center
        title_cell.fill = header_fill
        ws_teachers.row_dimensions[1].height = 34

        for col in range(1, fixed_cols + 1):
            ws_teachers.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)

        fixed_headers = ["م", "المعلم/ة", "المادة"]
        for col, text in enumerate(fixed_headers, start=1):
            cell = ws_teachers.cell(2, col, text)
            cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            cell.alignment = center
            cell.fill = header_fill
            cell.border = thin_border

        current_col = fixed_cols + 1
        for day in days:
            start_col = current_col
            end_col = current_col + len(periods) - 1
            ws_teachers.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
            day_cell = ws_teachers.cell(2, start_col, day)
            day_cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            day_cell.alignment = center
            day_cell.fill = header_fill
            day_cell.border = thin_border

            for p_idx, period in enumerate(periods):
                cell = ws_teachers.cell(3, start_col + p_idx, period)
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
                cell.alignment = center
                cell.fill = sub_header_fill
                cell.border = thin_border
            current_col = end_col + 1

        teacher_data = {}
        for teacher in all_teachers:
            teacher_df = df_result[
                df_result["المدرس"].astype(str).apply(lambda x: teacher in [part.strip() for part in x.split("/")])
            ].copy()
            teacher_data[teacher] = teacher_df

        for teacher_idx, teacher in enumerate(all_teachers, start=1):
            row_num = 3 + teacher_idx
            teacher_df = teacher_data[teacher]

            subjects = []
            for value in teacher_df.get("المادة", pd.Series(dtype=str)).astype(str):
                value = value.strip()
                if value and value.lower() != "nan" and value not in subjects:
                    subjects.append(value)
            subject_text = " / ".join(subjects)

            ws_teachers.cell(row_num, 1, teacher_idx)
            ws_teachers.cell(row_num, 2, teacher)
            ws_teachers.cell(row_num, 3, subject_text)

            for col in range(1, fixed_cols + 1):
                cell = ws_teachers.cell(row_num, col)
                cell.font = Font(name="Segoe UI", size=9, bold=True)
                cell.alignment = center
                cell.border = thin_border
                if col == 2:
                    cell.fill = title_fill

            current_col = fixed_cols + 1
            for day in days:
                for period in periods:
                    match = teacher_df[(teacher_df["اليوم"] == day) & (teacher_df["الحصة"] == period)]
                    cell = ws_teachers.cell(row_num, current_col)
                    cell.alignment = center
                    cell.border = thin_border
                    cell.font = Font(name="Segoe UI", size=8, bold=True)

                    if not match.empty:
                        entries, first_class = [], None
                        for _, item in match.iterrows():
                            cls_text = str(item.get("الفصل", "")).strip()
                            cls_parts = [x.strip() for x in cls_text.split(",") if x.strip()]
                            entries.extend(cls_parts)
                            if first_class is None and cls_parts:
                                first_class = cls_parts[0]

                        unique_entries = list(dict.fromkeys(entries))
                        cell.value = "\n".join(unique_entries)

                        if len(unique_entries) == 1:
                            color = class_colors.get(unique_entries[0])
                            if color:
                                cell.fill = PatternFill(fill_type="solid", fgColor=color)
                        elif unique_entries:
                            color = class_colors.get(first_class)
                            if color:
                                cell.fill = PatternFill(fill_type="solid", fgColor=color)
                    else:
                        cell.value = ""
                    current_col += 1

            ws_teachers.row_dimensions[row_num].height = 38

        ws_teachers.column_dimensions["A"].width = 6
        ws_teachers.column_dimensions["B"].width = 22
        ws_teachers.column_dimensions["C"].width = 18
        for col_idx in range(fixed_cols + 1, total_cols + 1):
            ws_teachers.column_dimensions[get_column_letter(col_idx)].width = 10

        ws_teachers.freeze_panes = "D4"

        legend_start = 5 + len(all_teachers)
        ws_teachers.cell(legend_start, 1, "مفتاح ألوان الفصول")
        ws_teachers.cell(legend_start, 1).font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        ws_teachers.cell(legend_start, 1).fill = header_fill
        ws_teachers.cell(legend_start, 1).alignment = center
        ws_teachers.cell(legend_start, 1).border = thin_border

        legend_col = 2
        for cls in sorted(classes, key=lambda x: str(x)):
            cls_text = str(cls).strip()
            cell = ws_teachers.cell(legend_start, legend_col, cls_text)
            cell.font = Font(name="Segoe UI", size=9, bold=True)
            cell.alignment = center
            cell.border = thin_border
            color = class_colors.get(cls_text)
            if color:
                cell.fill = PatternFill(fill_type="solid", fgColor=color)
            legend_col += 1
            if legend_col > total_cols:
                legend_col = 2
                legend_start += 1

        wb_master.save(master_table_file)
        print(f"📘 تم إنشاء ملف All Classes: {master_table_file}")

        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="Master_Schedule", index=False)

            summary_rows = []
            for _, row in df_result.iterrows():
                teacher_cell = str(row["المدرس"])
                for teacher in [t.strip() for t in teacher_cell.split("/") if t.strip()]:
                    summary_rows.append({"المعلم/ة": teacher, "الحصة": 1})

            if summary_rows:
                df_temp = pd.DataFrame(summary_rows)
                df_summary = df_temp.groupby("المعلم/ة")["الحصة"].sum().reset_index()
                df_summary.columns = ["المعلم/ة", "إجمالي الحصص الأسبوعية"]
                df_summary = df_summary.sort_values(by="المعلم/ة").reset_index(drop=True)
            else:
                df_summary = pd.DataFrame(columns=["المعلم/ة", "إجمالي الحصص الأسبوعية"])

            df_summary.to_excel(writer, sheet_name="كشف_المعلمين", index=False)
            pd.DataFrame({"ClassName": classes}).to_excel(writer, sheet_name="قائمة_الفصول", index=False)

            for c in classes:
                df_c = df_result[df_result["الفصل"].apply(lambda x: c in [i.strip() for i in str(x).split(",")])]
                if not df_c.empty:
                    df_c_copy = df_c.copy()
                    df_c_copy["عرض_الخلايا"] = df_c_copy["المادة"] + "\n(" + df_c_copy["المدرس"] + ")"
                    pivot_c = df_c_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first").fillna("فراغ").reindex(index=days, columns=periods)
                    pivot_c.to_excel(writer, sheet_name=f"فصل_{c}")

            for t in teachers:
                df_t = df_result[df_result["المدرس"].apply(lambda x: t in [i.strip() for i in str(x).split("/")])]
                off_days_for_t = teacher_off_days.get(t, [])

                if not df_t.empty:
                    df_t_copy = df_t.copy()
                    df_t_copy["عرض_الخلايا"] = df_t_copy["الفصل"] + "\n(" + df_t_copy["المادة"] + ")"
                    pivot_t = df_t_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first")
                else:
                    pivot_t = pd.DataFrame(index=days, columns=periods)

                pivot_t = pivot_t.reindex(index=days, columns=periods).astype(object)
                for day_name in days:
                    if day_name in off_days_for_t:
                        for p_col in periods:
                            pivot_t.loc[day_name, p_col] = "إجازة"
                    else:
                        for p_col in periods:
                            if pd.isna(pivot_t.loc[day_name, p_col]):
                                pivot_t.loc[day_name, p_col] = "راحة"

                pivot_t.to_excel(writer, sheet_name=f"مدرس_{t}")

            for room in rooms:
                df_r = df_result[df_result["القاعة"].astype(str).str.strip() == str(room).strip()]
                if not df_r.empty:
                    df_r_copy = df_r.copy()
                    df_r_copy["عرض_الخلايا"] = df_r_copy["الفصل"].astype(str) + "\n(" + df_r_copy["المادة"].astype(str) + " - " + df_r_copy["المدرس"].astype(str) + ")"
                    pivot_r = df_r_copy.pivot_table(index="اليوم", columns="الحصة", values="عرض_الخلايا", aggfunc="first")
                else:
                    pivot_r = pd.DataFrame("متاحة", index=days, columns=periods)

                pivot_r = pivot_r.reindex(index=days, columns=periods).fillna("متاحة").astype(object)
                pivot_r.to_excel(writer, sheet_name=f"قاعة_{room}"[:31])

        print("📁 جارٍ تنسيق وتجميل ملف Excel النهائي...")
        format_excel_workbook(out_file, school_name)
        print(f"✨ تم الحفظ بنجاح في: {out_file}")
        print(f"📘 ملف الحصص الشامل: {master_table_file}")

    else:
        print("❌ لم يتم العثور على حل ممكن (Infeasible Model).")

if __name__ == "__main__":
    generate_timetable()
