from flask import Flask, render_template, request
import pandas as pd
import os
import re

app = Flask(__name__)

EXCEL_FILE = "excel.xlsx"


def normalize_text(value):
    """
    توحيد النص لتسهيل البحث:
    - إزالة المسافات الزائدة
    - إزالة التشكيل
    - توحيد بعض الحروف العربية
    - تحويل الأرقام العربية إلى إنجليزية
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    arabic_numbers = str.maketrans(
        "٠١٢٣٤٥٦٧٨٩",
        "0123456789"
    )
    value = value.translate(arabic_numbers)

    value = re.sub(r"[\u064B-\u065F\u0670]", "", value)

    value = value.replace("أ", "ا")
    value = value.replace("إ", "ا")
    value = value.replace("آ", "ا")
    value = value.replace("ى", "ي")
    value = value.replace("ؤ", "و")
    value = value.replace("ئ", "ي")
    value = value.replace("ة", "ه")

    value = re.sub(r"\s+", " ", value)

    return value.lower().strip()


def normalize_class(value):
    """
    توحيد كتابة الصفوف، مثل:
    10\\1 أو 10-1 أو 10 / 1
    لتصبح: 10/1
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    arabic_numbers = str.maketrans(
        "٠١٢٣٤٥٦٧٨٩",
        "0123456789"
    )
    value = value.translate(arabic_numbers)

    value = value.replace("\\", "/")
    value = value.replace("-", "/")
    value = value.replace("–", "/")
    value = value.replace("ـ", "/")

    value = re.sub(r"\s*/\s*", "/", value)
    value = re.sub(r"\s+", "", value)

    # إزالة .0 إذا قرأ Excel الرقم بهذه الصورة
    value = re.sub(r"\.0$", "", value)

    return value


def find_column(dataframe, possible_names):
    """
    البحث عن اسم العمود حتى لو وُجدت مسافات زائدة.
    """
    columns_map = {
        normalize_text(column): column
        for column in dataframe.columns
    }

    for name in possible_names:
        normalized_name = normalize_text(name)

        if normalized_name in columns_map:
            return columns_map[normalized_name]

    return None


def load_students():
    """
    قراءة بيانات الطلاب من ملف Excel.
    """
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(), None, None, "ملف excel.xlsx غير موجود."

    try:
        dataframe = pd.read_excel(
            EXCEL_FILE,
            engine="openpyxl",
            dtype=str
        )

        dataframe.columns = [
            str(column).strip()
            for column in dataframe.columns
        ]

        name_column = find_column(
            dataframe,
            [
                "اسم الطالب",
                "اسم الطالب الثلاثي",
                "اسم الطالب الرباعي",
                "الطالب",
                "الاسم"
            ]
        )

        class_column = find_column(
            dataframe,
            [
                "الصف",
                "الصف الدراسي",
                "الفصل",
                "الشعبة",
                "الصف والشعبة"
            ]
        )

        if name_column is None:
            return (
                pd.DataFrame(),
                None,
                None,
                "لم يتم العثور على عمود اسم الطالب في ملف Excel."
            )

        if class_column is None:
            return (
                pd.DataFrame(),
                None,
                None,
                "لم يتم العثور على عمود الصف في ملف Excel."
            )

        dataframe[name_column] = (
            dataframe[name_column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        dataframe[class_column] = (
            dataframe[class_column]
            .apply(normalize_class)
        )

        # حذف الصفوف التي لا تحتوي على اسم طالب
        dataframe = dataframe[
            dataframe[name_column] != ""
        ].copy()

        dataframe["_normalized_name"] = (
            dataframe[name_column]
            .apply(normalize_text)
        )

        dataframe["_normalized_class"] = (
            dataframe[class_column]
            .apply(normalize_class)
        )

        return dataframe, name_column, class_column, None

    except Exception as error:
        return (
            pd.DataFrame(),
            None,
            None,
            f"حدث خطأ أثناء قراءة ملف Excel: {error}"
        )


def class_sort_key(class_name):
    """
    ترتيب الصفوف ترتيبًا منطقيًا:
    10/1 ثم 10/2 ثم 11/1...
    """
    numbers = re.findall(r"\d+", str(class_name))

    if numbers:
        return tuple(int(number) for number in numbers)

    return (999, str(class_name))


@app.route("/", methods=["GET", "POST"])
def index():
    dataframe, name_column, class_column, error = load_students()

    search_results = []
    selected_class = ""
    selected_class_students = []
    search_query = ""
    message = ""

    if error:
        return render_template(
            "index.html",
            error=error,
            classes=[],
            search_results=[],
            selected_class_students=[],
            selected_class="",
            search_query="",
            message="",
            name_column=None,
            class_column=None
        )

    classes = [
        class_name
        for class_name in dataframe[class_column].dropna().unique()
        if str(class_name).strip()
    ]

    classes = sorted(classes, key=class_sort_key)

    # البحث باسم الطالب
    if request.method == "POST":
        search_query = request.form.get(
            "student_name",
            ""
        ).strip()

        if search_query:
            normalized_query = normalize_text(search_query)

            matched_students = dataframe[
                dataframe["_normalized_name"].str.contains(
                    normalized_query,
                    case=False,
                    na=False,
                    regex=False
                )
            ]

            search_results = matched_students.to_dict("records")

            if not search_results:
                message = "لم يتم العثور على طالب مطابق لعملية البحث."
        else:
            message = "يرجى إدخال اسم الطالب."

    # عرض قائمة طلاب صف معين
    selected_class = request.args.get(
        "class_name",
        ""
    ).strip()

    if selected_class:
        normalized_selected_class = normalize_class(selected_class)

        class_students = dataframe[
            dataframe["_normalized_class"] == normalized_selected_class
        ].copy()

        class_students = class_students.sort_values(
            by=name_column,
            key=lambda column: column.map(normalize_text)
        )

        selected_class_students = class_students.to_dict("records")

        if not selected_class_students:
            message = "لا توجد أسماء طلاب في هذا الصف."

    return render_template(
        "index.html",
        error=None,
        classes=classes,
        search_results=search_results,
        selected_class_students=selected_class_students,
        selected_class=selected_class,
        search_query=search_query,
        message=message,
        name_column=name_column,
        class_column=class_column
    )


if __name__ == "__main__":
    app.run(debug=True)
