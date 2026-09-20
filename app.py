from flask import Flask, render_template, request, abort
from openpyxl import load_workbook
from pathlib import Path

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "excel.xlsx"


def normalize_text(value):
    """توحيد النص لتسهيل البحث باللغة العربية."""
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

    for old, new in replacements.items():
        text = text.replace(old, new)

    return " ".join(text.split())


def read_students():
    """قراءة الطلاب من أول ورقة في ملف Excel وإعطاء كل طالب رقمًا خاصًا."""
    if not EXCEL_FILE.exists():
        raise FileNotFoundError("لم يتم العثور على ملف excel.xlsx")

    workbook = load_workbook(EXCEL_FILE, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        return []

    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    students = []

    for excel_row_number, row in enumerate(rows[1:], start=2):
        if not any(value is not None and str(value).strip() for value in row):
            continue

        student = {}
        for index, header in enumerate(headers):
            if not header:
                header = f"البيان {index + 1}"
            student[header] = row[index] if index < len(row) else ""

        students.append({
            "id": excel_row_number,
            "data": student,
        })

    return students


def find_name_column(student):
    possible_names = [
        "اسم الطالب",
        "الاسم",
        "اسم",
        "student name",
        "name",
    ]

    normalized_names = [normalize_text(name) for name in possible_names]

    for column in student.keys():
        if normalize_text(column) in normalized_names:
            return column

    return next(iter(student.keys()), None)


def get_student_by_id(student_id):
    for student in read_students():
        if student["id"] == student_id:
            return student
    return None


@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    searched_name = ""
    error = ""

    if request.method == "POST":
        searched_name = request.form.get("student_name", "").strip()

        if not searched_name:
            error = "يرجى إدخال اسم الطالب أولًا."
        else:
            try:
                students = read_students()
                search_value = normalize_text(searched_name)

                for student in students:
                    name_column = find_name_column(student["data"])
                    student_name = student["data"].get(name_column, "") if name_column else ""

                    if search_value in normalize_text(student_name):
                        results.append(student)

                if not results:
                    error = "لم يتم العثور على طالب بهذا الاسم."

            except FileNotFoundError as exception:
                error = str(exception)
            except Exception as exception:
                error = f"حدث خطأ أثناء قراءة ملف Excel: {exception}"

    return render_template(
        "index.html",
        results=results,
        searched_name=searched_name,
        error=error,
    )


@app.route("/student/<int:student_id>")
def student_details(student_id):
    try:
        student = get_student_by_id(student_id)
    except FileNotFoundError as exception:
        abort(404, description=str(exception))

    if student is None:
        abort(404, description="لم يتم العثور على بيانات الطالب.")

    return render_template(
        "student_details.html",
        student=student["data"],
    )


@app.errorhandler(404)
def page_not_found(error):
    return render_template(
        "student_details.html",
        student=None,
        error_message=getattr(error, "description", "الصفحة غير موجودة."),
    ), 404


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
