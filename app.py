# -*- coding: utf-8 -*-
"""
نظام تحليل نتائج الاختبارات - ثيم وزارة التعليم
-------------------------------------------------
تطبيق Flask يقرأ ملف درجات (Excel/CSV)، يحسب النسب المئوية والتقديرات،
ويرتب الطلاب على مستوى الفصل وعلى مستوى الصف، مع إحصائيات وتصدير للنتائج.

صيغ الملف المدعومة:
  1. شيت واحد يحتوي أعمدة: اسم الطالب، الفصل، الدرجة
  2. شيت لكل فصل (اسم الشيت = اسم الفصل) يحتوي أعمدة: اسم الطالب، الدرجة
يمكن الجمع بين الصيغتين في الملف نفسه؛ إذا وُجد عمود الفصل داخل الشيت فهو الأولوية.

الترتيب: متسلسل (1، 2، 3...) حسب النسبة تنازليًا، وعند التساوي حسب الاسم أبجديًا.

التشغيل محليًا:
  python app.py
"""
import difflib
import io
import json
import os
import re
from collections import deque

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # الحد الأقصى لحجم الملف المرفوع: 10MB

# التقديرات حسب نموذج تحليل النتائج المعتمد في وزارة التعليم
LEVEL_NAMES = ["ممتاز", "جيد جدًا", "جيد", "مقبول", "ضعيف"]
DEFAULT_THRESHOLDS = (("t_excellent", 90), ("t_vgood", 80), ("t_good", 70), ("t_pass", 50))

# الأسماء المقبولة لكل عمود (عربي/إنجليزي)، تُطابق بعد التوحيد
COLUMN_ALIASES = {
    "name": ["اسم الطالب", "اسم الطالبة", "الاسم", "اسم", "الطالب", "الطالبة", "name", "student", "student name"],
    "class": ["الفصل", "فصل", "الشعبة", "شعبة", "class", "section"],
    "score": ["الدرجة", "درجة", "الدرجه", "درجه", "المجموع", "score", "grade", "mark", "total"],
}

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
SHEET_TITLE_INVALID = re.compile(r"[\\/*?:\[\]]")


# ---------------------------------------------------------------------------
# أدوات مساعدة
# ---------------------------------------------------------------------------
def normalize(text):
    """توحيد النص العربي لأغراض المقارنة (إزالة التشكيل، توحيد الهمزات والتاء المربوطة)."""
    text = str(text).strip().lower()
    text = re.sub(r"[\u064B-\u0652\u0640]", "", text)
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ة", "ه").replace("ى", "ي")
    return text


def find_column(df, key):
    """إرجاع اسم العمود المطابق للمفتاح المطلوب، أو None إذا لم يوجد."""
    aliases = [normalize(a) for a in COLUMN_ALIASES[key]]
    for col in df.columns:
        if normalize(col) in aliases:
            return col
    for col in df.columns:
        n = normalize(col)
        if any(a in n for a in aliases):
            return col
    return None


def to_number(value):
    """تحويل القيمة إلى رقم مع دعم الأرقام العربية والفواصل، أو None إذا تعذر."""
    if value is None or pd.isna(value):
        return None
    s = str(value).translate(ARABIC_DIGITS).replace("٫", ".").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def level_of(pct, thresholds):
    """تحديد التقدير بناءً على النسبة المئوية وحدود التقديرات (تنازلية)."""
    for threshold, name in zip(thresholds, LEVEL_NAMES):
        if pct >= threshold:
            return name
    return LEVEL_NAMES[-1]


def level_ranges(thresholds, total):
    """نطاق الدرجات الخام لكل تقدير، بالشكل المستخدم في نموذج الوزارة (مثال: 17.1 - 19)."""
    bounds = [round(t / 100 * total, 2) for t in thresholds]
    ranges = {}
    upper = total
    for name, low in zip(LEVEL_NAMES[:-1], bounds):
        ranges[name] = {"low": low, "high": upper}
        upper = round(low - 0.01, 2)
    ranges[LEVEL_NAMES[-1]] = {"low": 0, "high": upper}
    return ranges


def group_stats(df, pass_threshold, total):
    """إحصائيات وصفية لمجموعة من الطلاب (فصل أو الصف كامل)."""
    pct = df["pct"]
    score = df["score"]
    n = len(df)
    return {
        "count": int(n),
        "avg": round(float(pct.mean()), 2) if n else 0,
        "median": round(float(pct.median()), 2) if n else 0,
        "std": round(float(pct.std(ddof=0)), 2) if n > 1 else 0,
        "max": round(float(pct.max()), 2) if n else 0,
        "min": round(float(pct.min()), 2) if n else 0,
        "score_max": round(float(score.max()), 2) if n else 0,
        "score_min": round(float(score.min()), 2) if n else 0,
        "score_avg": round(float(score.mean()), 2) if n else 0,
        "score_sum": round(float(score.sum()), 2) if n else 0,
        # نسبة التحصيل = مجموع الدرجات ÷ (عدد الطلاب × الدرجة الكلية)
        "attainment": round(float(score.sum() / (n * total) * 100), 2) if n else 0,
        "pass_rate": round(float((pct >= pass_threshold).mean() * 100), 1) if n else 0,
        "levels": {name: int((df["level"] == name).sum()) for name in LEVEL_NAMES},
    }


HEADER_SCAN_ROWS = 30   # عدد الصفوف الأولى التي يُبحث فيها عن صف العناوين
MIN_DATA_ROWS = 1       # أقل عدد صفوف بيانات أسفل العنوان ليُقبل الصف كعنوان


def match_alias(cell, keys):
    """
    مطابقة نص خلية مع أسماء الأعمدة المعروفة.
    ترجع (المفتاح، نوع المطابقة) حيث النوع 2 لمطابقة تامة و1 لمطابقة جزئية، أو (None, 0).
    """
    text = normalize(cell)
    if not text or text in ("nan", "none"):
        return None, 0
    for key in keys:
        if any(text == normalize(a) for a in COLUMN_ALIASES[key]):
            return key, 2
    for key in keys:
        if any(normalize(a) in text for a in COLUMN_ALIASES[key]):
            return key, 1
    return None, 0


def detect_header(raw, keys, required=("name",)):
    """
    تحديد صف العناوين داخل ورقة خام (مقروءة بدون عناوين).

    يُفحص كل صف من الصفوف الأولى وتُحسب أعمدته المطابقة لأسماء معروفة؛ يفوز الصف
    صاحب أعلى عدد مطابقات تامة ثم جزئية، بشرط أن يحتوي الأعمدة المطلوبة وأن يليه
    صف بيانات واحد على الأقل. يتيح ذلك تجاوز الصفوف الفارغة أو العناوين التمهيدية
    (اسم المدرسة، عنوان التقرير) التي تسبق الجدول.

    ترجع (رقم صف العنوان، {المفتاح: رقم العمود}) أو (None, {}).
    """
    best_row, best_map, best_score = None, {}, (0, 0)
    limit = min(HEADER_SCAN_ROWS, len(raw))
    for i in range(limit):
        exact, partial, mapping = {}, {}, {}
        for j, cell in enumerate(raw.iloc[i]):
            key, kind = match_alias(cell, keys)
            if not key:
                continue
            if kind == 2 and key not in exact:
                exact[key] = j
            elif kind == 1 and key not in partial:
                partial[key] = j
        mapping = {**partial, **exact}
        if not all(k in mapping for k in required):
            continue
        # يجب أن يتبع صف العنوان بيانات فعلية
        below = raw.iloc[i + 1:]
        if below.dropna(how="all").shape[0] < MIN_DATA_ROWS:
            continue
        score = (len(exact), len(mapping))
        if score > best_score:
            best_row, best_map, best_score = i, mapping, score
    return best_row, best_map


def read_sheets(file):
    """
    قراءة الملف المرفوع وإرجاع قائمة (اسم الورقة، DataFrame خام بدون عناوين).
    تُقرأ الأوراق بلا عناوين ليجري تحديد صف العنوان لاحقًا، فتُتجاوز الصفوف
    التمهيدية أو الفارغة التي قد تسبق الجدول. ملفات CSV تُعامل كورقة واحدة.
    """
    name = (file.filename or "").lower()
    if name.endswith(".csv"):
        return [("الدرجات", pd.read_csv(file, header=None, dtype=object))]
    sheets = pd.read_excel(file, sheet_name=None, header=None, dtype=object)
    return [(n, df) for n, df in sheets.items() if not df.dropna(how="all").empty]


def build_dataset(sheets, need_score=True):
    """
    دمج أوراق الملف في جدول واحد بالأعمدة: name, class, score.

    لكل ورقة يُحدَّد صف العناوين تلقائيًا ثم تُقرأ البيانات أسفله. مصدر الفصل:
    عمود الفصل داخل الورقة إن وُجد، وإلا اسم الورقة نفسها. عند need_score=False
    لا تُشترط الدرجة (يُستخدم لشهادات لا تعتمد على درجات).

    ترجع (DataFrame، قائمة أخطاء، قائمة ملاحظات تشخيصية).
    """
    # عمود الدرجة يُكتشف دائمًا؛ need_score يحدد هل وجوده شرط لقبول الورقة
    keys = ["name", "class", "score"]
    required = ("name", "score") if need_score else ("name",)

    frames, errors, notes = [], [], []
    for sheet_name, raw in sheets:
        # يُحتفظ بالفهرس الأصلي ليعكس رقم الصف كما هو في ملف Excel
        raw = raw.dropna(how="all")
        if raw.empty:
            continue
        excel_rows = list(raw.index)
        raw = raw.reset_index(drop=True)

        header_row, mapping = detect_header(raw, keys, required)
        if header_row is None:
            preview = "، ".join(
                str(c).strip() for c in raw.iloc[0].tolist()[:6]
                if not pd.isna(c) and str(c).strip()
            ) or "صفوف فارغة"
            missing = "اسم الطالبة" + ("" if not need_score else " والدرجة")
            errors.append(f"الورقة «{sheet_name}»: لم يُعثر على جدول يحتوي عمود {missing}. "
                          f"أول صف فيها: {preview}")
            continue

        body = raw.iloc[header_row + 1:].reset_index(drop=True)
        name_col = mapping["name"]
        class_col = mapping.get("class")
        score_col = mapping.get("score")

        data = {"name": body[name_col].astype(str).str.strip()}
        if class_col is not None:
            class_values = body[class_col].astype(str).str.strip()
            class_values = class_values.replace({"nan": "", "None": ""})
            # الخلايا المدموجة تترك فراغات أسفل قيمة الفصل، فتُملأ بالقيمة السابقة
            class_values = class_values.replace("", pd.NA).ffill().fillna(str(sheet_name).strip())
            class_source = "عمود داخل الورقة"
        else:
            class_values = pd.Series([str(sheet_name).strip()] * len(body), index=body.index)
            class_source = "اسم الورقة"
        data["class"] = class_values.astype(str).str.translate(ARABIC_DIGITS)
        data["score"] = body[score_col].map(to_number) if score_col is not None else pd.NA

        frame = pd.DataFrame(data)
        frames.append(frame)

        header_cells = [str(raw.iloc[header_row, c]).strip()
                        for c in (name_col, class_col, score_col) if c is not None]
        notes.append(f"الورقة «{sheet_name}»: الجدول يبدأ من الصف {excel_rows[header_row] + 1} "
                     f"(الأعمدة: {'، '.join(header_cells)}) — الفصل من {class_source}")

    if not frames:
        return pd.DataFrame(columns=["name", "class", "score"]), errors, notes
    return pd.concat(frames, ignore_index=True), errors, notes


# كلمات نسب تُحذف قبل المقارنة لأنها تظهر في كشف دون آخر
NAME_FILLERS = {"بنت", "بن", "ابن", "ابنه", "الطالبه", "الطالب"}
FUZZY_THRESHOLD = 0.88   # أدنى تشابه يُقبل كتطابق تقريبي يحتاج مراجعة


def name_key(name):
    """مفتاح مقارنة للاسم: توحيد الحروف، إزالة الرموز وكلمات النسب، وضغط المسافات."""
    text = normalize(name)
    text = re.sub(r"[^\w\s]", " ", text)
    parts = [w for w in text.split() if w and w not in NAME_FILLERS]
    return " ".join(parts)


def class_key(value):
    """مفتاح مقارنة للفصل يتجاهل الفواصل وصيغة الأرقام (6/أ و6-أ سواء)."""
    text = normalize(str(value)).translate(ARABIC_DIGITS)
    return re.sub(r"[\s/\-_.]", "", text)


def is_subsequence(short_tokens, full_tokens):
    """هل تظهر كلمات الاسم المختصر داخل الاسم الكامل بالترتيب نفسه؟"""
    it = iter(full_tokens)
    return all(any(tok == f for f in it) for tok in short_tokens)


def expand_names(entries, roster):
    """
    استكمال الأسماء المختصرة (ثنائية غالبًا) من كشف الطالبات إلى أسمائها الكاملة.

    يُبحث أولًا داخل فصل الطالبة إن ذُكر، فإن لم يُذكر أو لم يُعثر على مرشح يُبحث
    في الكشف كله. يُقبل الاسم الكامل إذا وردت كلمات الاسم المختصر داخله بالترتيب
    نفسه، فيغطي ذلك «نورة محمد» ← «نورة محمد العتيبي» و«نورة العتيبي» ← الاسم ذاته.

    لا يُختار اسم عند تعدد المرشحين: يُترك الاسم كما هو ويُعلَّم للمراجعة، لأن
    الاختيار العشوائي بين طالبتين يطبع شهادة باسم خاطئ.

    ترجع قائمة مدخلات، لكل منها name وclass وoriginal وstatus وchoices.
    """
    by_class, everyone = {}, []
    for _, row in roster.iterrows():
        full = str(row["name"]).strip()
        tokens = name_key(full).split()
        if not tokens:
            continue
        record = (tokens, full, str(row["class"]).strip())
        everyone.append(record)
        by_class.setdefault(class_key(row["class"]), []).append(record)

    def search(tokens, pool):
        exact = [r for r in pool if r[0] == tokens]
        if exact:
            return exact, True
        return [r for r in pool if is_subsequence(tokens, r[0])], False

    results = []
    for entry in entries:
        original = str(entry.get("name", "")).strip()
        cls = str(entry.get("class", "")).strip()
        tokens = name_key(original).split()
        item = {"name": original, "class": cls, "original": original,
                "status": "notfound", "choices": []}
        if not tokens:
            results.append(item)
            continue

        pool = by_class.get(class_key(cls), []) if cls else []
        found, was_exact = search(tokens, pool) if pool else ([], False)
        scope = "class"
        if not found:
            found, was_exact = search(tokens, everyone)
            scope = "all"

        if len(found) == 1:
            matched_tokens, full, full_cls = found[0]
            item["name"] = full
            if not cls:
                item["class"] = full_cls
            if was_exact or matched_tokens == tokens:
                item["status"] = "exact"
            elif scope == "class":
                item["status"] = "expanded"
            else:
                item["status"] = "expanded_noclass"
                item["matched_class"] = full_cls
        elif len(found) > 1:
            item["status"] = "ambiguous"
            item["choices"] = [f"{full} ({c})" for _, full, c in found[:6]]
        results.append(item)
    return results


def match_students(roster, grades):
    """
    مطابقة كشف الطالبات بملف الدرجات على ثلاث مراحل: تطابق تام للمفتاح، ثم تطابق
    مجموعة الكلمات بصرف النظر عن ترتيبها، ثم تشابه نصي يتجاوز الحد المعتمد.
    المرحلتان الأخيرتان تُسجَّلان للمراجعة لأنهما اجتهاد لا يقين.

    ترجع (قائمة نتائج الكشف، مجموعة فهارس صفوف الدرجات المستخدمة).
    """
    by_key, by_tokens = {}, {}
    for idx, row in grades.iterrows():
        key = name_key(row["name"])
        by_key.setdefault(key, []).append(idx)
        by_tokens.setdefault(frozenset(key.split()), []).append(idx)

    results, used = [], set()

    def pick(candidates, cls):
        """اختيار المرشح الأنسب: من نفس الفصل أولًا، ثم أول مرشح لم يُستخدم."""
        same = [i for i in candidates if class_key(grades.at[i, "class"]) == cls]
        for pool in (same, candidates):
            for i in pool:
                if i not in used:
                    return i
        return candidates[0] if candidates else None

    for _, student in roster.iterrows():
        key = name_key(student["name"])
        cls = class_key(student["class"])
        match, how = None, None

        if key in by_key:
            match, how = pick(by_key[key], cls), "exact"
        else:
            tokens = frozenset(key.split())
            if tokens and tokens in by_tokens:
                match, how = pick(by_tokens[tokens], cls), "tokens"
            else:
                best, ratio = None, 0.0
                for cand_key, idxs in by_key.items():
                    r = difflib.SequenceMatcher(None, key, cand_key).ratio()
                    if r > ratio:
                        best, ratio = idxs, r
                if best and ratio >= FUZZY_THRESHOLD:
                    match, how = pick(best, cls), "fuzzy"

        if match is not None:
            used.add(match)
        results.append({"student": student, "match": match, "how": how})
    return results, used


def style_sheet(ws, headers, rows):
    """كتابة جدول منسق في ورقة عمل: رؤوس ملونة، حدود، محاذاة، اتجاه من اليمين لليسار."""
    head_fill = PatternFill("solid", fgColor="1E9E8A")
    head_font = Font(bold=True, color="FFFFFF", name="Arial")
    thin = Side(style="thin", color="C8D3CC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.sheet_view.rightToLeft = True
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment, cell.border = head_fill, head_font, center, border
    for row in rows:
        ws.append(row)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment, cell.border = center, border
    for i, header in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(str(header)) + 6)
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# المسارات
# ---------------------------------------------------------------------------
LOGO_DIRS = ("static", ".")
LOGO_EXTENSIONS = (".png", ".jpg", ".jpeg", ".svg", ".webp", ".gif")
# أسماء الملفات المقبولة لكل نوع صورة
IMAGE_STEMS = {
    "logo": ("logo",),                              # شعار وزارة التعليم
    "emblem": ("emblem", "شعار", "نخلة"),            # شعار المملكة: النخلة والسيفان
    "flag": ("flag", "علم", "saudi-flag", "saudiflag"),  # علم المملكة
    "stamp": ("stamp", "ختم", "طبعة"),               # ختم المعلمة بجانب اسمها
}


def find_image(kind):
    """البحث عن ملف صورة بالأسماء المقبولة للنوع المطلوب، في مجلد static أو مجلد المشروع."""
    stems = IMAGE_STEMS[kind]
    base = os.path.dirname(os.path.abspath(__file__))
    for folder in LOGO_DIRS:
        path = os.path.join(base, folder)
        if not os.path.isdir(path):
            continue
        for entry in sorted(os.listdir(path)):
            stem, ext = os.path.splitext(entry)
            if stem.lower() in stems and ext.lower() in LOGO_EXTENSIONS:
                return os.path.join(path, entry)
    return None


def find_logo():
    return find_image("logo")


@app.route("/")
def index():
    return render_template("index.html")


# ذاكرة مؤقتة للشعار بعد إزالة الخلفية، مفتاحها المسار وزمن آخر تعديل
_logo_cache = {}

# حدود اعتبار البكسل جزءًا من الخلفية البيضاء
BG_MIN_BRIGHTNESS = 232   # أقل قناة لونية تعتبر بعدها البكسل فاتحًا
BG_MAX_SATURATION = 22    # أقصى فرق بين القنوات ليعتبر البكسل رماديًا/أبيض
EDGE_SOFT_START = 190     # بداية التدرج لتنعيم حواف الشعار


def strip_background(path):
    """
    إزالة الخلفية البيضاء من صورة الشعار وإرجاعها بصيغة PNG شفافة.

    تُزال فقط المساحات الفاتحة المتصلة بحواف الصورة (تعبئة انتشارية)، حتى لا
    تُفقد المساحات البيضاء داخل الشعار نفسه. حواف الرسم تُنعَّم بشفافية متدرجة.
    تُرجع None إذا تعذّرت المعالجة، ليُقدَّم الملف الأصلي كما هو.
    """
    try:
        from PIL import Image
    except ImportError:
        return None

    try:
        img = Image.open(path).convert("RGBA")
    except Exception:
        return None

    # صورة تحمل شفافية فعلية تُقدَّم كما هي؛ معالجتها قد تقضم حوافها
    alpha = img.getchannel("A")
    if alpha.getextrema()[0] < 250:
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        return buf.getvalue()

    w, h = img.size
    if w * h > 4_000_000:  # صورة كبيرة جدًا للمعالجة البكسلية
        return None

    px = img.load()

    def is_light(x, y):
        r, g, b, a = px[x, y]
        if a == 0:
            return True
        return min(r, g, b) >= BG_MIN_BRIGHTNESS and max(r, g, b) - min(r, g, b) <= BG_MAX_SATURATION

    # تعبئة انتشارية من حواف الصورة لتحديد بكسلات الخلفية
    background = bytearray(w * h)
    queue = deque()
    for x in range(w):
        for y in (0, h - 1):
            if not background[y * w + x] and is_light(x, y):
                background[y * w + x] = 1
                queue.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            if not background[y * w + x] and is_light(x, y):
                background[y * w + x] = 1
                queue.append((x, y))

    while queue:
        x, y = queue.popleft()
        for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            if 0 <= nx < w and 0 <= ny < h and not background[ny * w + nx] and is_light(nx, ny):
                background[ny * w + nx] = 1
                queue.append((nx, ny))

    span = BG_MIN_BRIGHTNESS - EDGE_SOFT_START
    for y in range(h):
        row = y * w
        for x in range(w):
            r, g, b, a = px[x, y]
            if background[row + x]:
                px[x, y] = (r, g, b, 0)
            else:
                low = min(r, g, b)
                if low > EDGE_SOFT_START and max(r, g, b) - low <= BG_MAX_SATURATION:
                    # حافة فاتحة ملاصقة للرسم: شفافية متدرجة بدل قطع حاد
                    alpha = int(a * max(0.0, (BG_MIN_BRIGHTNESS - low) / span))
                    px[x, y] = (r, g, b, alpha)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


@app.route("/logo")
@app.route("/emblem")
@app.route("/flag")
@app.route("/stamp")
def logo():
    """
    إرجاع صورة بخلفية شفافة حسب المسار:
      /logo   شعار وزارة التعليم، يظهر في ترويسة التقرير والشهادات
      /stamp  ختم المعلمة، يظهر بجانب اسمها في جميع الشهادات
      /emblem شعار المملكة (النخلة والسيفان)، يظهر في شهادات ثيم اليوم الوطني
      /flag   علم المملكة، يظهر بجوار الشعار في ثيم اليوم الوطني
    المعامل raw=1 يُرجع الملف الأصلي دون معالجة.
    تُرجع 404 عند غياب الملف، فيُخفي القالب العنصر أو يعرض نصًا بديلًا.
    """
    path = find_image(request.path.lstrip("/"))
    if not path:
        return "", 404
    if request.args.get("raw") or path.lower().endswith(".svg"):
        return send_file(path)

    key = (path, os.path.getmtime(path))
    if key not in _logo_cache:
        _logo_cache.clear()
        _logo_cache[key] = strip_background(path)
    data = _logo_cache[key]
    if data is None:
        return send_file(path)
    return send_file(io.BytesIO(data), mimetype="image/png", download_name="logo.png")


@app.route("/analyze", methods=["POST"])
def analyze():
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="لم يتم رفع ملف الدرجات."), 400

    total = to_number(request.form.get("total", ""))
    if not total or total <= 0:
        return jsonify(error="الدرجة الكلية يجب أن تكون رقمًا أكبر من صفر."), 400

    thresholds = [to_number(request.form.get(k, d)) for k, d in DEFAULT_THRESHOLDS]
    if any(t is None for t in thresholds) or thresholds != sorted(thresholds, reverse=True):
        return jsonify(error="حدود التقديرات غير صحيحة؛ يجب أن تكون تنازلية "
                             "(ممتاز > جيد جدًا > جيد > مقبول)."), 400

    support_threshold = to_number(request.form.get("support", 60)) or 60

    try:
        sheets = read_sheets(file)
    except Exception:
        return jsonify(error="تعذّرت قراءة الملف. الصيغ المدعومة: Excel (.xlsx) أو CSV."), 400
    if not sheets:
        return jsonify(error="الملف لا يحتوي على أي بيانات."), 400

    data, errors, notes = build_dataset(sheets)
    if data.empty:
        return jsonify(error=" | ".join(errors) or "لا توجد بيانات صالحة في الملف."), 400

    # استبعاد الصفوف بدون اسم أو بدرجة غير رقمية
    invalid_mask = data["score"].isna() | data["name"].isin(["", "nan", "None"])
    skipped = int(invalid_mask.sum())
    data = data[~invalid_mask].reset_index(drop=True)
    if data.empty:
        return jsonify(error="لا توجد صفوف صالحة؛ يجب أن يحتوي عمود الدرجة على أرقام."), 400

    over = data[data["score"] > total]
    if not over.empty:
        return jsonify(error=f"يوجد {len(over)} طالب درجته أعلى من الدرجة الكلية ({total:g})، "
                             f"مثال: {over.iloc[0]['name']} = {over.iloc[0]['score']:g}."), 400

    # الحسابات الأساسية
    data["pct"] = (data["score"] / total * 100).round(2)
    data["level"] = data["pct"].map(lambda p: level_of(p, thresholds))

    # ترتيب متسلسل بدون تكرار: النسبة تنازليًا ثم الاسم أبجديًا
    data = data.sort_values(["pct", "name"], ascending=[False, True]).reset_index(drop=True)
    data["rank_grade"] = range(1, len(data) + 1)
    data["rank_class"] = data.groupby("class").cumcount() + 1

    classes = sorted(data["class"].unique().tolist())
    grade_stats = group_stats(data, thresholds[3], total)
    class_stats = {}
    for c in classes:
        st = group_stats(data[data["class"] == c], thresholds[3], total)
        st["gap"] = round(st["avg"] - grade_stats["avg"], 2)  # الفرق عن متوسط الصف
        class_stats[c] = st

    # توزيع النسب على فئات بعرض 10% (الفئة الأخيرة تشمل 100)
    bins = list(range(0, 100, 10))
    labels = [f"{b}–{b + 9}" for b in bins]
    labels[-1] = "90–100"

    def in_bin(series, b):
        return (series >= b) & ((series < b + 10) | (b == 90))

    histogram = {
        "labels": labels,
        "grade": [int(in_bin(data["pct"], b).sum()) for b in bins],
        "classes": {c: [int(in_bin(data.loc[data["class"] == c, "pct"], b).sum()) for b in bins]
                    for c in classes},
    }

    needs_support = data[data["pct"] < support_threshold].sort_values("pct")
    students = data[["name", "class", "score", "pct", "level", "rank_class", "rank_grade"]].to_dict("records")

    return jsonify({
        "total": total,
        "thresholds": thresholds,
        "ranges": level_ranges(thresholds, total),
        "support_threshold": support_threshold,
        "levels": LEVEL_NAMES,
        "classes": classes,
        "students": students,
        "grade_stats": grade_stats,
        "class_stats": class_stats,
        "histogram": histogram,
        "top10": students[:10],
        "needs_support": needs_support[["name", "class", "score", "pct", "level"]].to_dict("records"),
        "skipped": skipped,
        "warnings": errors,
        "detection": notes,
    })


@app.route("/names", methods=["POST"])
def names():
    """
    استخراج أسماء الطالبات من ملف Excel أو CSV دون اشتراط وجود درجات.
    يُستخدم في شهادات المشاركة والتقدير التي لا تعتمد على اختبار.
    """
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(error="لم يتم رفع ملف."), 400

    try:
        sheets = read_sheets(file)
    except Exception:
        return jsonify(error="تعذّرت قراءة الملف. الصيغ المدعومة: Excel (.xlsx) أو CSV."), 400
    if not sheets:
        return jsonify(error="الملف لا يحتوي على أي بيانات."), 400

    data, errors, notes = build_dataset(sheets, need_score=False)
    if data.empty:
        return jsonify(error=" | ".join(errors)
                       or "لم يُعثر على عمود باسم الطالبات في الملف."), 400

    data = data[~data["name"].isin(["", "nan", "None"])]
    data = data.drop_duplicates(subset=["name", "class"]).reset_index(drop=True)
    if data.empty:
        return jsonify(error="لا توجد أسماء صالحة في الملف."), 400

    return jsonify({
        "students": data[["name", "class"]].to_dict("records"),
        "warnings": errors,
        "detection": notes,
    })


def read_roster(file, label):
    """قراءة ملف أسماء وإرجاع (DataFrame, أخطاء, ملاحظات) أو رفع ValueError برسالة عربية."""
    try:
        sheets = read_sheets(file)
    except Exception:
        raise ValueError(f"تعذّرت قراءة {label}. الصيغ المدعومة: Excel (.xlsx) أو CSV.")
    if not sheets:
        raise ValueError(f"{label} لا يحتوي على أي بيانات.")

    data, errors, notes = build_dataset(sheets, need_score=False)
    data = data[~data["name"].isin(["", "nan", "None"])].reset_index(drop=True)
    if data.empty:
        raise ValueError(" | ".join(errors) or f"لم يُعثر على عمود أسماء في {label}.")
    return data, [f"{label}: {e}" for e in errors], [f"{label} — {n}" for n in notes]


@app.route("/expand_names", methods=["POST"])
def expand_names_route():
    """
    استكمال أسماء الطالبات من كشف الطالبات، لاستخدامها في الشهادات.

    يُرفع كشف الطالبات، وتُرسل معه قائمة الأسماء الحالية بصيغة JSON في الحقل
    names. يُرجَع لكل اسم صيغته الكاملة إن وُجد مرشح واحد، وإلا يُترك كما هو
    مع بيان سبب التعذّر.
    """
    try:
        roster, errors, notes = read_roster(request.files.get("roster"), "كشف الطالبات")
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    except (AttributeError, TypeError):
        return jsonify(error="يجب رفع كشف الطالبات."), 400

    try:
        entries = json.loads(request.form.get("names", "[]"))
        if not isinstance(entries, list):
            raise ValueError
    except (ValueError, TypeError):
        return jsonify(error="قائمة الأسماء غير صالحة."), 400
    if not entries:
        return jsonify(error="لا توجد أسماء لاستكمالها."), 400

    results = expand_names(entries, roster)
    counts = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    return jsonify({
        "students": results,
        "counts": counts,
        "roster_size": int(len(roster)),
        "warnings": errors,
        "detection": notes,
    })


@app.route("/absent", methods=["POST"])
def absent():
    """
    مقارنة كشف الطالبات الكامل بملف الدرجات لتحديد من لم تختبر.

    تُعدّ الطالبة مختبِرة إذا طوبقت بصف في ملف الدرجات يحمل درجة رقمية. وما عدا
    ذلك يُصنَّف: غير مدرجة في ملف الدرجات، أو مدرجة بلا درجة. كما تُرجَع أسماء
    وردت في ملف الدرجات ولا مقابل لها في الكشف، لأنها غالبًا اختلاف إملائي.
    """
    try:
        roster, r_err, r_notes = read_roster(request.files.get("roster"), "كشف الطالبات")
        grades, g_err, g_notes = read_roster(request.files.get("grades"), "ملف الدرجات")
    except ValueError as exc:
        return jsonify(error=str(exc)), 400
    except (AttributeError, TypeError):
        return jsonify(error="يجب رفع الملفين: كشف الطالبات وملف الدرجات."), 400

    roster = roster.drop_duplicates(subset=["name", "class"]).reset_index(drop=True)
    results, _ = match_students(roster, grades)

    absent_rows, review_rows = [], []
    tested_idx = set()
    for item in results:
        student, idx, how = item["student"], item["match"], item["how"]
        entry = {"name": student["name"], "class": student["class"]}
        if idx is None:
            absent_rows.append({**entry, "reason": "غير مدرجة في ملف الدرجات"})
            continue
        score = grades.at[idx, "score"]
        if score is None or pd.isna(score):
            absent_rows.append({**entry, "reason": "مدرجة في ملف الدرجات بلا درجة"})
        else:
            tested_idx.add(idx)
        if how in ("tokens", "fuzzy"):
            review_rows.append({**entry, "matched": str(grades.at[idx, "name"]),
                                "how": "ترتيب مختلف للكلمات" if how == "tokens" else "تشابه إملائي"})

    matched_any = {it["match"] for it in results if it["match"] is not None}
    extra_rows = [
        {"name": grades.at[i, "name"], "class": grades.at[i, "class"]}
        for i in grades.index if i not in matched_any
    ]

    absent_keys = {(r["name"], r["class"]) for r in absent_rows}
    classes = sorted({str(c).strip() for c in roster["class"] if str(c).strip()})
    stats = {}
    for cls in classes:
        rows = roster[roster["class"].astype(str).str.strip() == cls]
        total = int(len(rows))
        missing = sum(1 for _, r in rows.iterrows() if (r["name"], r["class"]) in absent_keys)
        stats[cls] = {"total": total, "absent": missing, "tested": total - missing}

    absent_rows.sort(key=lambda r: (str(r["class"]), str(r["name"])))
    return jsonify({
        "classes": classes,
        "absent": absent_rows,
        "extra": extra_rows,
        "review": review_rows,
        "stats": stats,
        "totals": {"roster": int(len(roster)), "graded": int(len(tested_idx)),
                   "absent": len(absent_rows)},
        "warnings": r_err + g_err,
        "detection": r_notes + g_notes,
    })


@app.route("/absent_export", methods=["POST"])
def absent_export():
    """تصدير نتيجة المقارنة إلى Excel: ورقة لمن لم تختبر، وورقة إحصائيات، وورقتا مراجعة."""
    payload = request.get_json(silent=True) or {}
    absent_rows = payload.get("absent", [])
    if not absent_rows:
        return jsonify(error="لا توجد نتائج للتصدير."), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "لم تختبر"
    style_sheet(ws, ["م", "اسم الطالبة", "الفصل", "الحالة"],
                [[i, r["name"], r["class"], r["reason"]] for i, r in enumerate(absent_rows, 1)])

    stats = payload.get("stats", {})
    style_sheet(wb.create_sheet("إحصائية الفصول"),
                ["الفصل", "عدد الطالبات", "اختبرن", "لم تختبر", "نسبة الغياب %"],
                [[c, v["total"], v["tested"], v["absent"],
                  round(v["absent"] / v["total"] * 100, 1) if v["total"] else 0]
                 for c, v in stats.items()])

    review = payload.get("review", [])
    if review:
        style_sheet(wb.create_sheet("تطابق يحتاج مراجعة"),
                    ["اسم الطالبة في الكشف", "الفصل", "الاسم في ملف الدرجات", "سبب المراجعة"],
                    [[r["name"], r["class"], r["matched"], r["how"]] for r in review])

    extra = payload.get("extra", [])
    if extra:
        style_sheet(wb.create_sheet("في الدرجات وليست بالكشف"),
                    ["اسم الطالبة", "الفصل"], [[r["name"], r["class"]] for r in extra])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="من_لم_تختبر.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/export", methods=["POST"])
def export():
    """تصدير النتائج إلى ملف Excel: ترتيب الصف، شيت لكل فصل، الإحصائيات، الأوائل، والطلاب الذين يحتاجون متابعة."""
    payload = request.get_json(silent=True) or {}
    students = payload.get("students", [])
    classes = payload.get("classes", [])
    class_stats = payload.get("class_stats", {})
    grade_stats = payload.get("grade_stats", {})
    total = payload.get("total", "")
    if not students:
        return jsonify(error="لا توجد نتائج للتصدير."), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "ترتيب الصف"
    style_sheet(ws,
                ["م", "اسم الطالب", "الفصل", "الدرجة", f"النسبة % (من {total})", "التقدير"],
                [[s["rank_grade"], s["name"], s["class"], s["score"], s["pct"], s["level"]] for s in students])

    for c in classes:
        ws = wb.create_sheet(SHEET_TITLE_INVALID.sub("-", f"فصل {c}")[:31])
        rows = sorted((s for s in students if s["class"] == c), key=lambda s: s["rank_class"])
        style_sheet(ws,
                    ["م", "اسم الطالب", "الدرجة", "النسبة %", "التقدير", "الترتيب على الصف"],
                    [[s["rank_class"], s["name"], s["score"], s["pct"], s["level"], s["rank_grade"]] for s in rows])

    ws = wb.create_sheet("الإحصائيات")
    headers = ["المجموعة", "عدد الطلاب", "أعلى درجة", "أقل درجة", "متوسط الدرجات", "نسبة التحصيل %",
               "مجموع الدرجات", "نسبة الاجتياز %"] + LEVEL_NAMES
    rows = []
    groups = [(f"فصل {c}", class_stats.get(c, {})) for c in classes] + [("الصف كامل", grade_stats)]
    for label, st in groups:
        rows.append([label, st.get("count"), st.get("score_max"), st.get("score_min"), st.get("score_avg"),
                     st.get("attainment"), st.get("score_sum"), st.get("pass_rate")]
                    + [st.get("levels", {}).get(l, 0) for l in LEVEL_NAMES])
    style_sheet(ws, headers, rows)

    top10 = payload.get("top10", [])
    style_sheet(wb.create_sheet("الأوائل"),
                ["م", "اسم الطالب", "الفصل", "الدرجة", "النسبة %"],
                [[s["rank_grade"], s["name"], s["class"], s["score"], s["pct"]] for s in top10])

    support = payload.get("needs_support", [])
    threshold = payload.get("support_threshold", "")
    style_sheet(wb.create_sheet("يحتاجون متابعة"),
                ["اسم الطالب", "الفصل", "الدرجة", f"النسبة % (أقل من {threshold})", "التقدير"],
                [[s["name"], s["class"], s["score"], s["pct"], s["level"]] for s in support])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="نتائج_التحليل.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/template")
def template():
    """نموذج Excel جاهز: شيت لكل فصل، واسم الشيت هو اسم الفصل."""
    wb = Workbook()
    wb.remove(wb.active)
    for cls, rows in (("6-أ", [["نورة محمد", 45], ["سارة فهد", 42]]),
                      ("6-ب", [["ريم خالد", 38], ["هند عمر", 30]]),
                      ("6-ج", [["لمى ناصر", 47], ["جود علي", 29]])):
        style_sheet(wb.create_sheet(cls), ["اسم الطالب", "الدرجة"], rows)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="نموذج_الدرجات.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
