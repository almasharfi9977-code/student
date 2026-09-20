from pathlib import Path

from flask import Flask, abort, render_template, request
from openpyxl import load_workbook


app = Flask(__name__)

# تحديد مكان المشروع الحالي
BASE_DIR = Path(__file__).resolve().parent

# ملف قاعدة بيانات الطلاب
EXCEL_FILE = BASE_DIR / "excel.xlsx"


def normalize_text(value):
    """
    توحيد النص العربي لتسهيل البحث.
    يعالج بعض الاختلافات بين الحروف العربية.
    """
    if value is None:
        return ""

    text = str(value).strip().lower()

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",
        "ى": "ي",
        "ؤ": "و",
        "ئ": "ي",
        "ـ": "",
    }

    for old_character, new_character in replacements.items():
        text = text.replace(old_character, new_character)

    # إزالة المسافات الزائدة
    return " ".join(text.split())


def read_students():
    """
    قراءة بيانات الطلاب من ملف Excel.

    الصف الأول في ملف Excel يجب أن يحتوي على أسماء الأعمدة،
    مثل:
    اسم الطالب | الصف | الولاية
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            "لم يتم العثور على ملف excel.xlsx. "
            "تأكد من وضعه في نفس مجلد app.py"
        )

    # data_only=True لقراءة القيم بدلًا من معادلات Excel
    workbook = load_workbook(EXCEL_FILE, data_only=True)

    # استخدام أول ورقة في ملف Excel
    sheet = workbook.active

    # قراءة جميع الصفوف
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        return []

    # الصف الأول هو عناوين الأعمدة
    headers = []

    for index, value in enumerate(rows[0]):
        if value is None or str(value).strip() == "":
            headers.append(f"البيان {index + 1}")
        else:
            headers.append(str(value).strip())

    students = []

    # بداية قراءة بيانات الطلاب من الصف الثاني
    for excel_row_number, row in enumerate(rows[1:], start=2):

        # تجاهل الصفوف الفارغة
        if not any(
            value is not None and str(value).strip()
            for value in row
        ):
            continue

        student_data = {}

        for index, header in enumerate(headers):
            if index < len(row):
                value = row[index]
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


def find_name_column(student_data):
    """
    تحديد العمود الذي يحتوي على اسم الطالب.

    يقبل أسماء أعمدة مثل:
    اسم الطالب
    الاسم
    اسم
    name
    student name
    """

    possible_name_columns = [
        "اسم الطالب",
        "الاسم",
        "اسم",
        "name",
        "student name",
    ]

    normalized_possible_names = [
        normalize_text(name)
        for name in possible_name_columns
    ]

    for column_name in student_data.keys():
        normalized_column_name = normalize_text(column_name)

        if normalized_column_name in normalized_possible_names:
            return column_name

    # إذا لم يتم العثور على عمود معروف، يستخدم أول عمود
    if student_data:
        return next(iter(student_data.keys()))

    return None


def get_student_by_id(student_id):
    """
    الحصول على طالب محدد باستخدام رقم صفه في Excel.
    """

    students = read_students()

    for student in students:
        if student["id"] == student_id:
            return student

    return None


@app.route("/", methods=["GET", "POST"])
def index():
    """
    الصفحة الرئيسية والبحث عن الطلاب.
    """

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
                students = read_students()

                # النص الذي أدخله المستخدم بعد التوحيد
                search_value = normalize_text(searched_name)

                for student in students:

                    student_data = student["data"]

                    name_column = find_name_column(student_data)

                    if name_column:
                        student_name = student_data.get(
                            name_column,
                            ""
                        )
                    else:
                        student_name = ""

                    # البحث بجزء من الاسم
                    if search_value in normalize_text(student_name):
                        results.append(student)

                if not results:
                    error = "لم يتم العثور على طالب بهذا الاسم."

            except FileNotFoundError as exception:
                error = str(exception)

            except Exception as exception:
                error = (
                    "حدث خطأ أثناء قراءة ملف Excel: "
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
    """
    صفحة عرض جميع تفاصيل الطالب.
    """

    try:
        student = get_student_by_id(student_id)

    except FileNotFoundError as exception:
        abort(404, description=str(exception))

    except Exception as exception:
        abort(
            404,
            description=(
                "حدث خطأ أثناء قراءة بيانات الطالب: "
                f"{exception}"
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
    """
    صفحة الخطأ عند عدم العثور على الطالب أو الصفحة.
    """

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
        debug=True,
        host="127.0.0.1",
        port=5000,
    )
