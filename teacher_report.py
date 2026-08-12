import pandas as pd

# قراءة الملف الأصلي
excel_path = "inputs.xlsx"
df_assign = pd.read_excel(excel_path, sheet_name="Assignments")

# تجميع حصص المعلمين
teacher_summary = (
    df_assign.groupby("Teacher")["WeeklyLessons"].sum().reset_index()
)
teacher_summary.columns = ["المعلم", "إجمالي الحصص"]

# حفظ الملف باسم إنجليزي لضمان عدم حدوث تعليق
output_filename = "teachers_report.xlsx"

with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
  teacher_summary.to_excel(writer, sheet_name="Workload", index=False)
  df_assign.to_excel(writer, sheet_name="Assignments_Raw", index=False)

print("تم توليد الملف بنجاح تام باسم: teachers_report.xlsx")