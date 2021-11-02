# Quizmaker

A workflow automation tool for the creation of pub trivia slide decks. Quizmaker inserts text from a Google Sheet into a PowerPoint template file while handling slide creation (round headers, questions/answers, promotional bumpers), text formatting, and file management, all behind a dead-simple web interface.

A test environment (including dummy data) can be found at https://vml.pythonanywhere.com. Test env may be non-functional during active development.

Built on:
- Python 3.9 with `gspread` and `python-pptx` libraries
- Flask 2.0
- PythonAnywhere web hosting (free)

## Requirements
This tool relies on a few proprietary components not stored in this repo, namely:
- A .pptx file with explicit slide template layouts already defined (see `quizmaker.py`).
- A Google Sheets template document with round category/description, questions, and answers.
- A `gspread` service account with proper credentials/access to said Google Sheet.

A dummy sheet has been linked to [the test environment](https://vml.pythonanywhere.com) for demo/development purposes.

## How to use

1. Staff quiz writers insert their round categories/descriptions, questions, and answers into the Google Sheet.
2. In the Quizmaker webapp, select the sheet name from the dropdown and ensure that the "Audio Round 3?" checkbox is accurate.
3. Save the file presented by your browser. It's that easy!

## License

The code in this repo is available under GPL 3.0. Quizmaker was originally created for use by [Buzz Worthy Pub Trivia](https://www.buzzworthypubtrivia.com/) -- check them out if you're in the Pittsburgh or Cleveland areas!
