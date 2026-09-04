# -*- coding: utf-8 -*-
"""ينشئ ملف بيانات تجريبي: sample_grades.xlsx"""
import random
import pandas as pd
random.seed(7)
first = ["محمد","أحمد","خالد","سعد","فهد","عبدالله","سلطان","نواف","تركي","بندر","يوسف","ريان","عمر","حمد","ماجد"]
last  = ["العتيبي","القحطاني","الشمري","الدوسري","المطيري","الغامدي","الزهراني","الحربي","السبيعي","العنزي"]
rows = []
for cls, bias in (("6/أ", 0), ("6/ب", -4), ("6/ج", 3), ("6/د", -8)):
    for _ in range(25):
        rows.append({"اسم الطالب": f"{random.choice(first)} {random.choice(last)}",
                     "الفصل": cls,
                     "الدرجة": max(8, min(50, round(random.gauss(38 + bias, 8), 1)))})
pd.DataFrame(rows).to_excel("sample_grades.xlsx", index=False)
print("تم إنشاء sample_grades.xlsx")
