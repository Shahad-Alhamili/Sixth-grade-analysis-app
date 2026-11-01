# -*- coding: utf-8 -*-
"""
إنشاء ملفي بيانات تجريبية للاختبار:
  sample_grades.xlsx          شيت واحد بعمود الفصل
  sample_grades_sheets.xlsx   شيت لكل فصل (اسم الشيت = الفصل)
"""
import random

import pandas as pd

random.seed(7)
FIRST = ["محمد", "أحمد", "خالد", "سعد", "فهد", "عبدالله", "سلطان", "نواف", "تركي", "بندر",
         "يوسف", "ريان", "عمر", "حمد", "ماجد"]
LAST = ["العتيبي", "القحطاني", "الشمري", "الدوسري", "المطيري", "الغامدي", "الزهراني",
        "الحربي", "السبيعي", "العنزي"]
CLASSES = (("6-أ", 0), ("6-ب", -4), ("6-ج", 3), ("6-د", -8))  # (الفصل، انحراف المتوسط)


def make_class(bias, n=25):
    return [{"اسم الطالب": f"{random.choice(FIRST)} {random.choice(LAST)}",
             "الدرجة": max(8, min(50, round(random.gauss(38 + bias, 8), 1)))} for _ in range(n)]


rows = []
with pd.ExcelWriter("sample_grades_sheets.xlsx") as writer:
    for cls, bias in CLASSES:
        students = make_class(bias)
        pd.DataFrame(students).to_excel(writer, sheet_name=cls, index=False)
        rows += [{**s, "الفصل": cls} for s in students]

pd.DataFrame(rows)[["اسم الطالب", "الفصل", "الدرجة"]].to_excel("sample_grades.xlsx", index=False)
print("تم إنشاء sample_grades.xlsx و sample_grades_sheets.xlsx")
