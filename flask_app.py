from pathlib import Path

from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

import quizmaker

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1000 # max uploaded file size

HOMEDIR = Path(__file__).parent

@app.route('/')
def form():
    return render_template("home.html")

@app.route('/download', methods=["POST"])
def convert():
    #pattern from https://flask.palletsprojects.com/en/2.0.x/patterns/fileuploads/
    raw_file = request.files['data_file']
    if raw_file and raw_file.filename.endswith(".csv"):
        pass
    else:
        "Invalid file; please select a .csv file generated by Google Sheets."
    filename = secure_filename(raw_file.filename)
    raw_file.save(HOMEDIR / "uploads" / filename)

    quizfile = f"{filename[:filename.find('.csv')]}.pptx"
    quiz = quizmaker.build_quiz(HOMEDIR / "static/pptx_template.pptx", HOMEDIR / "uploads" / filename)
    quiz.save(HOMEDIR / "output" / quizfile)

    return send_from_directory(HOMEDIR / "output", quizfile, as_attachment=True)