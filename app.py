from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = 'excel.xlsx'

@app.route('/', methods=['GET', 'POST'])
def index():
    student_data = None
    error_message = None
    search_query = ''

    if request.method == 'POST':
        search_query = request.form.get('student_name', '').strip()

        if search_query:
            if os.path.exists(EXCEL_FILE):
                try:
                    df = pd.read_excel(EXCEL_FILE)
                    df = df.fillna('—')

                    search_words = search_query.strip().split()

                    # يجب إدخال 3 كلمات على الأقل
                    if len(search_words) < 3:
                        error_message = "يرجى كتابة ثلاثة أجزاء من اسم الطالب على الأقل."
                    else:
                        results = df[
                            df['اسم الطالب'].astype(str).apply(
                                lambda name: all(
                                    word.lower() in str(name).lower()
                                    for word in search_words
                                )
                            )
                        ]

                        if not results.empty:
                            student_data = results.to_dict(orient='records')
                        else:
                            error_message = 'لم يتم العثور على طالب بهذا الاسم.'

                except Exception as e:
                    error_message = f'حدث خطأ أثناء قراءة البيانات: {e}'

            else:
                error_message = 'ملف excel.xlsx غير موجود.'

        else:
            error_message = 'يرجى إدخال اسم الطالب.'

    return render_template(
        'index.html',
        student_data=student_data,
        error_message=error_message,
        search_query=search_query
    )

if __name__ == '__main__':
    app.run(debug=True)
