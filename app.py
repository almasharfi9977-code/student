from pathlib import Path

from flask import Flask, abort, render_template, request
from openpyxl import load_workbook


app = Flask(__name__)

# مسار مجلد المشروع
BASE_DIR = Path(__file__).resolve().parent

# يجب أن يكون ملف Excel بجانب app.py
EXCEL_FILE = BASE_DIR / "excel.xlsx"


def normalize_text(value):
    """توحيد النص العربي لتسهيل البحث."""
    if value is None:
        return ""

    text = str(value).strip().lower()

    # إزالة الحركات والتنوين والتطويل
    arabic_marks = "ًٌٍَُِّْـ"

    for mark in arabic_marks:
        text = text.replace(mark, "")

    # توحيد بعض الحروف العربية
    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",
        "ى": "ي",
        "ؤ": "و",
        "ئ": "ي",
    }

    for old_character, new_character in replacements.items():
        text = text.replace(old_character, new_character)

    # توحيد المسافات
    return " ".join(text.split())


def read_students():
    """
    قراءة الطلاب من ملف Excel.

    يجب أن يكون الصف الأول عناوين الأعمدة، مثل:
    اسم الطالب | الصف | الولاية
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            "لم يتم العثور على ملف excel.xlsx. "
            "تأكد من وضعه في نفس مجلد app.py."
        )

    try:
        workbook = load_workbook(
            EXCEL_FILE,
            data_only=True
        )
    except Exception as exception:
        raise RuntimeError(
            f"تعذر فتح ملف Excel: {exception}"
        ) from exception

    # استخدام أول ورقة في ملف Excel
    worksheet = workbook.active

    # قراءة جميع الصفوف
    rows = list(
        worksheet.iter_rows(values_only=True)
    )

    if not rows:
        return []

    # قراءة عناوين الأعمدة من الصف الأول
    headers = []

    for index, value in enumerate(rows[0], start=1):
        if value is None or str(value).strip() == "":
            headers.append(f"البيان {index}")
        else:
            headers.append(str(value).strip())

    students = []

    # بيانات الطلاب تبدأ من الصف الثاني
    for excel_row_number, row in enumerate(rows[1:], start=2):

        # تجاهل الصفوف الفارغة بالكامل
        if not any(
            value is not None and str(value).strip() != ""
            for value in row
        ):
            continue

        student_data = {}

        for column_index, header in enumerate(headers):
            if column_index < len(row):
                value = row[column_index]
            else:
                value = ""

            student_data[header] = value

        students.append(
            {
                # رقم الصف في ملف Excel
                "id": excel_row_number,

                # بيانات الطالب
                "data": student_data,
            }
        )

    return students


def search_students(search_text):
    """
    البحث في جميع خلايا صف الطالب داخل ملف Excel.

    يمكن البحث بالاسم أو جزء من الاسم أو الصف أو الولاية.
    """

    normalized_search = normalize_text(search_text)

    if not normalized_search:
        return []

    matched_students = []

    students = read_students()

    for student in students:
        student_data = student["data"]

        # جمع جميع خلايا الصف في نص واحد
        searchable_text = " ".join(
            normalize_text(value)
            for value in student_data.values()
            if value is not None
        )

        # البحث بجزء من النص
        if normalized_search in searchable_text:
            matched_students.append(student)

    return matched_students


def get_student_by_id(student_id):
    """الحصول على طالب حسب رقم الصف في ملف Excel."""

    students = read_students()

    for student in students:
        if student["id"] == student_id:
            return student

    return None


@app.route("/", methods=["GET", "POST"])
def index():
    """الصفحة الرئيسية والبحث عن الطلاب."""

    results = []
    searched_name = ""
    error = ""

    if request.method == "POST":

        searched_name = request.form.get(
            "student_name",
            ""
        ).strip()

        if not searched_name:
            error = "يرجى إدخال اسم الطالب أولًا."

        else:
            try:
                results = search_students(searched_name)

                if not results:
                    error = "لم يتم العثور على طالب بهذا الاسم."

            except FileNotFoundError as exception:
                error = str(exception)

            except RuntimeError as exception:
                error = str(exception)

            except Exception as exception:
                error = (
                    "حدث خطأ أثناء البحث في ملف Excel: "
                    f"{exception}"
                )

    return render_template(
        "index.html",
        results=results,
        searched_name=searched_name,
        error=error,
    )


@app.route("/student/<int:student_id>")
def student_details(student_id):
    """عرض جميع تفاصيل طالب محدد."""

    try:
        student = get_student_by_id(student_id)

    except Exception as exception:
        abort(
            404,
            description=(
                f"تعذر قراءة بيانات الطالب: {exception}"
            ),
        )

    if student is None:
        abort(
            404,
            description="لم يتم العثور على بيانات الطالب."
        )

    return render_template(
        "student_details.html",
        student=student["data"],
    )


@app.errorhandler(404)
def page_not_found(error):
    """عرض رسالة مفهومة عند عدم العثور على صفحة أو طالب."""

    return render_template(
        "student_details.html",
        student=None,
        error_message=getattr(
            error,
            "description",
            "الصفحة غير موجودة.",
        ),
    ), 404


if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True,
    )
