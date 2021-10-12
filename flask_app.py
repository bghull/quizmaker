from pathlib import Path

from flask import Flask, g, json, render_template, request, send_from_directory

import quizmaker

HOMEDIR = Path(__file__).parent

try:
    import gspread
except ImportError:
    import sys
    sys.path.insert(0, str(HOMEDIR.parent / ".local/lib/python3.9/site-packages"))
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
        if all(("/" not in sheet.title,
        "\\" not in sheet.title,
        sheet.title not in ("How To Use This Document", "Template"))):
            sheets.append(sheet.title)
    sheets.sort()
    changelog = json.load(open(HOMEDIR / "CHANGELOG"))
    return render_template("home.html", sheets=sheets, changelog=changelog)

@app.route("/download", methods=["POST"])
def convert_sheet():
    sheet = g.sh.worksheet(request.form.get("sheet"))
    data = sheet.get_all_records()
    audio = bool(request.form.get("r3audio"))
    outfile = f"{sheet.title}.pptx"
    quiz = quizmaker.build_quiz(HOMEDIR / "static/pptx_template.pptx", data, audio=audio)
    quiz.save(HOMEDIR / "output" / outfile)
    return send_from_directory(HOMEDIR / "output", outfile, as_attachment=True)