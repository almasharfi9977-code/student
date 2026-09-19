from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "excel.xlsx"


def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE)

        df.columns = df.columns.str.strip()

        return df

    except Exception:
        return pd.DataFrame()


@app.route("/", methods=["GET", "POST"])
def index():

    df = load_data()

    student_data = []
    error_message = ""

    search_query = ""

    classes = []
    selected_class = ""
    class_students = []

    name_column = "اسم الطالب"
    class_column = "الصف"

    if not df.empty:

        if class_column in df.columns:

            classes = sorted(
                df[class_column]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        selected_class = request.args.get(
            "class_name",
            ""
        )

        if selected_class:

            class_students = (
                df[
                    df[class_column]
                    .astype(str)
                    == selected_class
                ]
                .fillna("")
                .to_dict("records")
            )

    if request.method == "POST":

        search_query = request.form.get(
            "student_name",
            ""
        ).strip()

        if search_query:

            try:

                results = df[
                    df[name_column]
                    .astype(str)
                    .str.contains(
                        search_query,
                        case=False,
                        na=False
                    )
                ]

                if not results.empty:
                    student_data = (
                        results
                        .fillna("")
                        .to_dict("records")
                    )
                else:
                    error_message = (
                        "لم يتم العثور على الطالب."
                    )

            except Exception:
                error_message = (
                    "تعذر تنفيذ البحث."
                )
        else:
            error_message = (
                "يرجى إدخال اسم الطالب."
            )

    return render_template(
        "index.html",
        student_data=student_data,
        error_message=error_message,
        search_query=search_query,
        classes=classes,
        selected_class=selected_class,
        class_students=class_students,
        name_column=name_column,
        class_column=class_column
    )


if __name__ == "__main__":
    app.run(debug=True)
