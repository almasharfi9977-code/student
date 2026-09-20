from pathlib import Path

from flask import Flask, abort, render_template, request
from openpyxl import load_workbook

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "excel.xlsx"

NAME_COLUMN_NAMES = {
    "اسم الطالب",
    "الاسم",
    "اسم",
    "name",
    "student name",
}


def normalize_text(value):
    """توحيد النص العربي لتسهيل مطابقة الأسماء."""

    if value is None:
        return ""

    text = str(value).strip().lower()

    # إزالة الحركات والتنوين والتطويل
    for mark in "ًٌٍَُِّْـ":
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

    return " ".join(text.split())


def read_students():
    """قراءة بيانات الطلاب من ملف excel.xlsx."""

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            "لم يتم العثور على ملف excel.xlsx. "
            "ضعه في نفس مجلد app.py."
        )

    try:
        workbook = load_workbook(EXCEL_FILE, data_only=True)
    except Exception as exception:
        raise RuntimeError(
            f"تعذر فتح ملف Excel: {exception}"
        ) from exception

    worksheet = workbook.active
    rows = list(worksheet.iter_rows(values_only=True))

    if not rows:
        return []

    headers = []

    for index, value in enumerate(rows[0], start=1):
        if value is None or str(value).strip() == "":
            headers.append(f"البيان {index}")
        else:
            headers.append(str(value).strip())

    students = []

    for excel_row_number, row in enumerate(rows[1:], start=2):

        if not any(
            value is not None and str(value).strip() != ""
            for value in row
        ):
            continue

        student_data = {}

        for column_index, header in enumerate(headers):
            value = row[column_index] if column_index < len(row) else ""
            student_data[header] = value

        students.append({
            "id": excel_row_number,
            "data": student_data,
        })

    return students


def find_name_column(student_data):
    """العثور على عمود اسم الطالب."""

    normalized_possible_names = {
        normalize_text(name)
        for name in NAME_COLUMN_NAMES
    }

    for column_name in student_data.keys():

        normalized_column = normalize_text(column_name)

        if (
            normalized_column in normalized_possible_names
            or (
                "اسم" in normalized_column
                and "طالب" in normalized_column
            )
        ):
            return column_name

    if student_data:
        return next(iter(student_data.keys()))

    return None


def search_students(search_text):
    """
    البحث المرن عن الطالب.
    يسمح بالأسماء الطويلة وبالأسماء التي تحتوي على (بن).
    """

    search_text = normalize_text(search_text)

    if len(search_text.split()) < 3:
        return []

    matched_students = []

    for student in read_students():

        student_data = student["data"]

        name_column = find_name_column(student_data)

        if not name_column:
            continue

        student_name = normalize_text(
            student_data.get(name_column, "")
        )

        # جميع كلمات البحث يجب أن تكون موجودة في الاسم
        if all(
            word in student_name
            for word in search_text.split()
        ):
            matched_students.append(student)

    return matched_students


def get_student_by_id(student_id):
    """الحصول على طالب حسب رقم صفه في Excel."""

    for student in read_students():
        if student["id"] == student_id:
            return student

    return None


@app.route("/", methods=["GET", "POST"])
def index():
    """الصفحة الرئيسية."""

    results = []
    searched_name = ""
    error = ""

    if request.method == "POST":

        searched_name = request.form.get(
            "student_name",
            ""
        ).strip()

        search_words = normalize_text(
            searched_name
        ).split()

        if len(search_words) < 3:
            error = (
                "يرجى كتابة ثلاثة أسماء على الأقل."
            )

        else:
            try:

                results = search_students(
                    searched_name
                )

                if not results:
                    error = (
                        "لم يتم العثور على طالب مطابق للاسم المدخل."
                    )

            except FileNotFoundError as exception:
                error = str(exception)

            except RuntimeError as exception:
                error = str(exception)

            except Exception as exception:
                error = (
                    f"حدث خطأ أثناء البحث في ملف Excel: "
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
    """عرض تفاصيل الطالب."""

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
