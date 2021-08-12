import csv
import sys

from pptx import Presentation

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

def start_round(prs, number, data):
    slide = prs.slides.add_slide(prs.slide_layouts[ROUND_START])
    slide.placeholders[0].text = f"Round {number}"
    slide.placeholders[10].text = data[f"R{number}C"]
    slide.placeholders[11].text = data[f"R{number}D"]
    #slide.placeholders[11].text_frame.fit_text(max_size=28)
    return None


def build_normal_round(prs, number, data):
    global bumpers
    start_round(prs, number, data)
    # Questions
    for q in range(1, 8):
        slide = prs.slides.add_slide(prs.slide_layouts[QUESTION])
        slide.placeholders[0].text = f"Question {q}"
        slide.placeholders[10].text = data[f"R{number}Q{q}"]
        #slide.placeholders[10].text_frame.fit_text(max_size=60)
    prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
    # Answer 1-4
    slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_1])
    slide.placeholders[0].text = f"Round {number} Answers"
    p = 10
    for q in range(1, 5):
        q_pholder = p
        a_pholder = p + 1
        slide.placeholders[q_pholder].text = f"Q{q}: " + data[f"R{number}Q{q}"]
        slide.placeholders[a_pholder].text = f"A{q}: " + data[f"R{number}A{q}"]
        p += 2
    # Answers 5-7
    slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_2])
    slide.placeholders[0].text = f"Round {number} Answers (cont.)"
    p = 10
    for q in range(5, 8):
        question = p
        answer = p + 1
        slide.placeholders[question].text = f"Q{q}: " + data[f"R{number}Q{q}"]
        slide.placeholders[answer].text = f"A{q}: " + data[f"R{number}A{q}"]
        p += 2


def build_audio_round(prs, number, data):
    global bumpers
    start_round(prs, number, data)
    # No question slides for audio round
    prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
    slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_AUDIO])
    slide.placeholders[0].text = f"Round {number} Answers"
    for a in range(1, 8):
        slide.placeholders[a + 9].text = f"A{a}: " + data[f"R{number}A{a}"]


def build_quiz(template, data):
    global bumpers
    bumpers = [BUMPER_1, BUMPER_4, BUMPER_3, BUMPER_2, BUMPER_1]
    prs = Presentation(template)
    all_data = import_questions(data)
    prs.slides.add_slide(prs.slide_layouts[RULES])
    build_normal_round(prs, 1, all_data)
    build_normal_round(prs, 2, all_data)
    build_audio_round(prs, 3, all_data)
    build_normal_round(prs, 4, all_data)
    build_normal_round(prs, 5, all_data)
    prs.slides.add_slide(prs.slide_layouts[CLOSER])
    return prs


if __name__ == "__main__":
    template = sys.argv[1]
    infile = sys.argv[2]
    outfile = infile[: infile.find(".")]
    quiz = build_quiz(template, infile)
    quiz.save(f"{outfile}.pptx")
    print(f"File saved as {outfile}.pptx")
