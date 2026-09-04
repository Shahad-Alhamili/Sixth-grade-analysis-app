# -*- coding: utf-8 -*-
"""
نظام تحليل درجات الصف السادس - ثيم وزارة التعليم
تشغيل:  python app.py   ثم افتحي  http://127.0.0.1:5000
"""
import io
import re
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

LEVEL_NAMES = ["ممتاز", "جيد جدًا", "جيد", "مقبول", "غير مجتاز"]

# الأعمدة المقبولة في ملف الإكسل (عربي أو إنجليزي)
COLUMN_ALIASES = {
    "name":  ["اسم الطالب", "الاسم", "اسم", "الطالب", "name", "student", "student name"],
    "class": ["الفصل", "فصل", "الشعبة", "شعبة", "class", "section"],
    "score": ["الدرجة", "درجة", "الدرجه", "درجه", "المجموع", "score", "grade", "mark", "total"],
}

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def normalize(text):
    """توحيد النص العربي للمقارنة"""
    text = str(text).strip().lower()
    text = re.sub(r"[\u064B-\u0652\u0640]", "", text)          # تشكيل و تطويل
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ة", "ه").replace("ى", "ي")
    return text


def find_column(df, key):
    aliases = [normalize(a) for a in COLUMN_ALIASES[key]]
    for col in df.columns:
        if normalize(col) in aliases:
            return col
    for col in df.columns:  # مطابقة جزئية
        n = normalize(col)
        if any(a in n for a in aliases):
            return col
    return None


def to_number(value):
    if pd.isna(value):
        return None
    s = str(value).translate(ARABIC_DIGITS).replace("٫", ".").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def level_of(pct, thresholds):
    if pct >= thresholds[0]:
        return LEVEL_NAMES[0]
    if pct >= thresholds[1]:
        return LEVEL_NAMES[1]
    if pct >= thresholds[2]:
        return LEVEL_NAMES[2]
    if pct >= thresholds[3]:
        return LEVEL_NAMES[3]
    return LEVEL_NAMES[4]


def group_stats(df, pass_threshold):
    pct = df["pct"]
    levels = {name: int((df["level"] == name).sum()) for name in LEVEL_NAMES}
    return {
        "count": int(len(df)),
        "avg": round(float(pct.mean()), 2) if len(df) else 0,
        "median": round(float(pct.median()), 2) if len(df) else 0,
        "std": round(float(pct.std(ddof=0)), 2) if len(df) > 1 else 0,
        "max": round(float(pct.max()), 2) if len(df) else 0,
        "min": round(float(pct.min()), 2) if len(df) else 0,
        "pass_rate": round(float((pct >= pass_threshold).mean() * 100), 1) if len(df) else 0,
        "levels": levels,
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="ارفعي ملف الإكسل أولًا."), 400

    total = to_number(request.form.get("total", ""))
    if not total or total <= 0:
        return jsonify(error="الدرجة الكلية لازم تكون رقم أكبر من صفر."), 400

    try:
        thresholds = [to_number(request.form.get(k, d)) for k, d in
                      (("t_excellent", 90), ("t_vgood", 80), ("t_good", 65), ("t_pass", 50))]
        if any(t is None for t in thresholds) or thresholds != sorted(thresholds, reverse=True):
            raise ValueError
    except ValueError:
        return jsonify(error="حدود المستويات غير صحيحة (لازم تكون تنازلية: ممتاز > جيد جدًا > جيد > مقبول)."), 400

    support_threshold = to_number(request.form.get("support", 60)) or 60

    try:
        if file.filename.lower().endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
    except Exception:
        return jsonify(error="ما قدرت أقرأ الملف. تأكدي إنه ملف Excel (.xlsx) أو CSV."), 400

    df.columns = [str(c).strip() for c in df.columns]
    name_col, class_col, score_col = (find_column(df, k) for k in ("name", "class", "score"))
    missing = [label for label, col in
               (("اسم الطالب", name_col), ("الفصل", class_col), ("الدرجة", score_col)) if col is None]
    if missing:
        return jsonify(error=f"الأعمدة الناقصة في الملف: {'، '.join(missing)}. "
                             f"الأعمدة الموجودة: {'، '.join(df.columns)}"), 400

    data = pd.DataFrame({
        "name": df[name_col].astype(str).str.strip(),
        "class": df[class_col].astype(str).str.strip().str.translate(ARABIC_DIGITS),
        "score": df[score_col].map(to_number),
    })
    skipped = data[data["score"].isna() | (data["name"] == "") | (data["name"] == "nan")]
    data = data.drop(skipped.index).reset_index(drop=True)
    if data.empty:
        return jsonify(error="ما فيه صفوف صالحة (تأكدي إن عمود الدرجة أرقام)."), 400

    over = data[data["score"] > total]
    if not over.empty:
        return jsonify(error=f"فيه {len(over)} طالب درجته أعلى من الدرجة الكلية ({total:g})، "
                             f"مثال: {over.iloc[0]['name']} = {over.iloc[0]['score']:g}."), 400

    data["pct"] = (data["score"] / total * 100).round(2)
    data["level"] = data["pct"].map(lambda p: level_of(p, thresholds))
    data["rank_grade"] = data["pct"].rank(method="min", ascending=False).astype(int)
    data["rank_class"] = data.groupby("class")["pct"].rank(method="min", ascending=False).astype(int)
    data = data.sort_values(["rank_grade", "name"]).reset_index(drop=True)

    classes = sorted(data["class"].unique().tolist())
    per_class = {c: group_stats(data[data["class"] == c], thresholds[3]) for c in classes}

    # فجوة كل فصل عن متوسط الصف
    grade_stats = group_stats(data, thresholds[3])
    for c in classes:
        per_class[c]["gap"] = round(per_class[c]["avg"] - grade_stats["avg"], 2)

    # توزيع الدرجات (هستوجرام) لكل 10%
    bins = list(range(0, 101, 10))
    labels = [f"{b}–{b+9}" for b in bins[:-1]]
    labels[-1] = "90–100"
    def in_bin(series, b):
        return (series >= b) & ((series < b + 10) | (b == 90))

    histogram = {
        "labels": labels,
        "grade": [int(in_bin(data["pct"], b).sum()) for b in bins[:-1]],
        "classes": {c: [int(in_bin(data.loc[data["class"] == c, "pct"], b).sum()) for b in bins[:-1]]
                    for c in classes},
    }

    needs_support = data[data["pct"] < support_threshold]

    students = data[["name", "class", "score", "pct", "level", "rank_class", "rank_grade"]].to_dict("records")

    return jsonify({
        "total": total,
        "thresholds": thresholds,
        "support_threshold": support_threshold,
        "levels": LEVEL_NAMES,
        "classes": classes,
        "students": students,
        "grade_stats": grade_stats,
        "class_stats": per_class,
        "histogram": histogram,
        "top10": students[:10],
        "needs_support": needs_support.sort_values("pct")[["name", "class", "score", "pct", "level"]].to_dict("records"),
        "skipped": int(len(skipped)),
    })


@app.route("/export", methods=["POST"])
def export():
    payload = request.get_json(silent=True) or {}
    students = payload.get("students", [])
    classes = payload.get("classes", [])
    class_stats = payload.get("class_stats", {})
    grade_stats = payload.get("grade_stats", {})
    total = payload.get("total", "")
    if not students:
        return jsonify(error="ما فيه نتائج للتصدير."), 400

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="0B5D3B")
    head_font = Font(bold=True, color="FFFFFF", name="Arial")
    thin = Side(style="thin", color="C8D3CC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    def write_sheet(ws, headers, rows):
        ws.sheet_view.rightToLeft = True
        ws.append(headers)
        for cell in ws[1]:
            cell.fill, cell.font, cell.alignment, cell.border = head_fill, head_font, center, border
        for r in rows:
            ws.append(r)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment, cell.border = center, border
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(str(h)) + 6)
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "ترتيب الصف"
    write_sheet(ws, ["الترتيب على الصف", "اسم الطالب", "الفصل", "الدرجة", f"النسبة % (من {total})", "المستوى"],
                [[s["rank_grade"], s["name"], s["class"], s["score"], s["pct"], s["level"]] for s in students])

    for c in classes:
        ws = wb.create_sheet(re.sub(r'[\\/*?:\[\]]', "-", f"فصل {c}")[:31])
        rows = sorted([s for s in students if s["class"] == c], key=lambda s: s["rank_class"])
        write_sheet(ws, ["الترتيب على الفصل", "اسم الطالب", "الدرجة", "النسبة %", "المستوى", "الترتيب على الصف"],
                    [[s["rank_class"], s["name"], s["score"], s["pct"], s["level"], s["rank_grade"]] for s in rows])

    ws = wb.create_sheet("الإحصائيات")
    levels = LEVEL_NAMES
    headers = ["المجموعة", "عدد الطلاب", "المتوسط %", "الوسيط", "الأعلى", "الأدنى", "نسبة الاجتياز %"] + levels
    rows = []
    for c in classes:
        st = class_stats.get(c, {})
        rows.append([f"فصل {c}", st.get("count"), st.get("avg"), st.get("median"), st.get("max"), st.get("min"),
                     st.get("pass_rate")] + [st.get("levels", {}).get(l, 0) for l in levels])
    rows.append(["الصف كامل", grade_stats.get("count"), grade_stats.get("avg"), grade_stats.get("median"),
                 grade_stats.get("max"), grade_stats.get("min"), grade_stats.get("pass_rate")]
                + [grade_stats.get("levels", {}).get(l, 0) for l in levels])
    write_sheet(ws, headers, rows)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="نتائج_الصف_السادس.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/template")
def template():
    """نموذج إكسل فاضي بالأعمدة المطلوبة"""
    wb = Workbook()
    ws = wb.active
    ws.title = "الدرجات"
    ws.sheet_view.rightToLeft = True
    ws.append(["اسم الطالب", "الفصل", "الدرجة"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B5D3B")
        cell.alignment = Alignment(horizontal="center")
    ws.append(["محمد أحمد", "6/أ", 45])
    ws.append(["خالد سعد", "6/ج", 38])
    for col in "ABC":
        ws.column_dimensions[col].width = 20
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="نموذج_الدرجات.xlsx")


if __name__ == "__main__":
    import os
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
