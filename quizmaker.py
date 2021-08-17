import csv
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor

TITLE = 0
RULES = 1
ROUND_START = 2
QUESTION = 3
ANSWERS_1 = 4
ANSWERS_2 = 5
ANSWERS_AUDIO = 6
BUMPER_1 = 7
BUMPER_2 = 8
BUMPER_3 = 9
BUMPER_4 = 10
CLOSER = 11


def import_questions(infile):
    with open(infile, newline="", encoding="utf-8-sig") as csvfile:
        csvdict = csv.DictReader(csvfile)
        return {"R" + d["Round"] + d["Type"][:1] + d["Number"]: d["Text"] for d in csvdict}


def build_round(prs, number, data, audio=False):
    global bumpers
    slide = prs.slides.add_slide(prs.slide_layouts[ROUND_START])
    slide.placeholders[0].text = f"Round {number}"
    slide.placeholders[10].text = data[f"R{number}C"]
    slide.placeholders[11].text = data[f"R{number}D"]
    if audio: # No question slides for audio rounds
        prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_AUDIO])
        slide.placeholders[0].text = f"Round {number} Answers"
        for a in range(1, 8):
            slide.placeholders[a + 9].text = f"A{a}: " + data[f"R{number}A{a}"]
    else:
        # Questions
        for q in range(1, 8):
            slide = prs.slides.add_slide(prs.slide_layouts[QUESTION])
            slide.placeholders[0].text = f"Question {q}"
            slide.placeholders[10].text = data[f"R{number}Q{q}"]
            #slide.placeholders[10].text_frame.fit_text(max_size=60)
        prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
        # Answers 1-4
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_1])
        slide.placeholders[0].text = f"Round {number} Answers"
        tf = slide.placeholders[10].text_frame
        tf.paragraphs[0].text = f"Q1: {data[f'R{number}Q1']}"
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = f"A1: {data[f'R{number}A1']}"
        run.font.bold = True
        run.font.color.rgb = RGBColor(25, 217, 203)
        for q in range(2, 5):
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"Q{q}: {data[f'R{number}Q{q}']}"
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"A{q}: {data[f'R{number}A{q}']}"
            run.font.bold = True
            run.font.color.rgb = RGBColor(25, 217, 203)
        # Answers 5-7
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_2])
        slide.placeholders[0].text = f"Round {number} Answers (cont.)"
        tf = slide.placeholders[10].text_frame
        tf.paragraphs[0].text = f"Q5: {data[f'R{number}Q5']}"
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = f"A5: {data[f'R{number}A5']}"
        run.font.bold = True
        run.font.color.rgb = RGBColor(25, 217, 203)
        for q in range(6, 8):
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"Q{q}: {data[f'R{number}Q{q}']}"
            p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"A{q}: {data[f'R{number}A{q}']}"
            run.font.bold = True
            run.font.color.rgb = RGBColor(25, 217, 203)


def build_quiz(template, data):
    global bumpers
    bumpers = [BUMPER_1, BUMPER_4, BUMPER_3, BUMPER_2, BUMPER_1]
    prs = Presentation(template)
    all_data = import_questions(data)
    prs.slides.add_slide(prs.slide_layouts[RULES])
    build_round(prs, 1, all_data)
    build_round(prs, 2, all_data)
    build_round(prs, 3, all_data, audio=True)
    build_round(prs, 4, all_data)
    build_round(prs, 5, all_data)
    prs.slides.add_slide(prs.slide_layouts[CLOSER])
    return prs


if __name__ == "__main__":
    template = sys.argv[1]
    infile = sys.argv[2]
    outfile = infile[: infile.find(".")]
    quiz = build_quiz(template, infile)
    quiz.save(f"{outfile}.pptx")
    print(f"File saved as {outfile}.pptx")
