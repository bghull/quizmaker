from pathlib import Path

from flask import Flask, g, json, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

import quizmaker

HOMEDIR = Path(__file__).parent

try:
    import gspread
except ImportError:
    import sys
    sp = HOMEDIR.parent / ".local/lib/python3.9/site-packages"
    sys.path.insert(0, str(sp))
    import gspread

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1000 # max uploaded file size

@app.before_request
def get_sheets():
    gc = gspread.service_account(filename=HOMEDIR / "service_account.json")
    sh = gc.open("BWPT Quiz Data")
    g.sh = sh
    return None

@app.route("/")
def home():
    sheets = []
    for sheet in g.sh.worksheets():
        if ("/" not in sheet.title) and (sheet.title not in ("How To Use This Document", "Template")):
            sheets.append(sheet.title)
    sheets.sort()
    changelog = json.load(open(HOMEDIR / "CHANGELOG"))
    return render_template("home.html", sheets=sheets, changelog=changelog)

@app.route("/sheet", methods=["POST"])
def convert_sheet():
    sheet = g.sh.worksheet(request.form.get("sheet"))
    data = sheet.get_values()
    audio = bool(request.form.get("r3audio"))
    quizfile = f"{sheet.title}.pptx"
    quiz = quizmaker.build_quiz(HOMEDIR / "static/pptx_template.pptx", data, audio=audio, sheet=True)
    quiz.save(HOMEDIR / "output" / quizfile)
    return send_from_directory(HOMEDIR / "output", quizfile, as_attachment=True)

@app.route("/download", methods=["POST"])
def convert():
    #pattern from https://flask.palletsprojects.com/en/2.0.x/patterns/fileuploads/
    raw_file = request.files["data_file"]
    audio = bool(request.form.get("r3audio"))
    if raw_file and raw_file.filename.endswith(".csv"):
        pass
    else:
        "Invalid file; please select a .csv file generated by Google Sheets."
    filename = secure_filename(raw_file.filename)
    raw_file.save(HOMEDIR / "uploads" / filename)

    quizfile = f"{filename[:filename.find('.csv')]}.pptx"
    quiz = quizmaker.build_quiz(HOMEDIR / "static/pptx_template.pptx", HOMEDIR / "uploads" / filename, audio=audio)
    quiz.save(HOMEDIR / "output" / quizfile)

    return send_from_directory(HOMEDIR / "output", quizfile, as_attachment=True)