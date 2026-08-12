import pandas as pd

# 1. شيت المعلمين
data_teachers = {
    'Teacher': ['ahmed', 'mona', 'sara', 'hassan', 'ali']
}

# 2. شيت الفصول
data_classes = {
    'ClassName': ['b1', 'b2']
}

# 3. شيت التكليفات (تم ضبط أعداد الحصص لمنع التضارب)
data_assignments = {
    'ClassName': ['b1', 'b1', 'b1', 'b1', 'b1', 'b2', 'b2', 'b2', 'b2', 'b2'],
    'Subject': ['arabic', 'english', 'math', 'science', 'social',
                'arabic', 'english', 'math', 'science', 'social'],
    'Teacher': ['ahmed', 'mona', 'sara', 'hassan', 'ali',
                'ahmed', 'mona', 'sara', 'hassan', 'ali'],
    'WeeklyLessons': [5, 4, 5, 3, 3, 5, 4, 5, 3, 3],
    'PreferredRoom': ['Classroom', 'Classroom', 'Classroom', 'Lab_1', 'Classroom',
                      'Classroom', 'Classroom', 'Classroom', 'Lab_2', 'Classroom']
}

# 4. شيت الإعدادات
data_settings = {
    'PeriodsPerDay': [7]
}

# 5. شيت الأيام
data_days = {
    'DayName': ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس']
}

with pd.ExcelWriter('inputs.xlsx', engine='openpyxl') as writer:
    pd.DataFrame(data_teachers).to_excel(writer, sheet_name='Teachers', index=False)
    pd.DataFrame(data_classes).to_excel(writer, sheet_name='Classes', index=False)
    pd.DataFrame(data_assignments).to_excel(writer, sheet_name='Assignments', index=False)
    pd.DataFrame(data_settings).to_excel(writer, sheet_name='Settings', index=False)
    pd.DataFrame(data_days).to_excel(writer, sheet_name='Days', index=False)

print("تم إنشاء ملف inputs.xlsx بنجاح!")