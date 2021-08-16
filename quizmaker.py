import csv
import sys
from pprint import pprint

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
        reader = csv.reader(csvfile)
        data = [row for row in reader]
        res = {}
        pprint(data[:2])
        # for r in range(1, 6):
        #     data[f"R{r}"] = {"C": data
        #     for q in range(1, 8):
        #         data[f"R{r}"] = {


d = import_questions("/home/vml/mysite/uploads/08-08.csv")
try:
    print(d["R2"]["Category"])
    print(d["R2"]["Q"][4]["Text"] == "Named by a savvy sales manager who wanted something easily pronounced by children, this small lollipop is a favorite of bank tellers and doctor's receptionists.")
except Exception as exc:
    print("Incorrect format:", exc)


def build_round(prs, r, data, audio=False):
    global bumpers
    slide = prs.slides.add_slide(prs.slide_layouts[ROUND_START])
    slide.placeholders[0].text = f"Round {r}"
    slide.placeholders[10].text = data[f"R{r}C"]["Text"]
    slide.placeholders[11].text = data[f"R{r}D"]["Text"]
    #slide.placeholders[11].text_frame.fit_text(max_size=28)
    if audio: # No questions on audio rounds
        prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_AUDIO])
        slide.placeholders[0].text = f"Round {r} Answers"
        for a in range(1, 8):
            slide.placeholders[a + 9].text = f"A{a}: " + data[f"R{r}A{a}"]["Text"]
    # Normal round questions
    else:
        for q in range(1, 8):
            slide = prs.slides.add_slide(prs.slide_layouts[QUESTION])
            slide.placeholders[0].text = f"Question {q}"
            slide.placeholders[10].text = data[f"R{r}Q{q}"]["Text"]
            slide.notes_slide.notes_text_frame.text = data[f"R{r}Q{q}"]["Notes"]
            #slide.placeholders[10].text_frame.fit_text(max_size=60)
        prs.slides.add_slide(prs.slide_layouts[bumpers.pop()])
        # Answer 1-4
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_1])
        slide.placeholders[0].text = f"Round {r} Answers"
        p = 10
        for q in range(1, 5):
            q_pholder = p
            a_pholder = p + 1
            slide.placeholders[q_pholder].text = f"Q{q}: " + data[f"R{r}Q{q}"]["Text"]
            slide.placeholders[a_pholder].text = f"A{q}: " + data[f"R{r}A{q}"]["Text"]
            slide.notes_slide.notes_text_frame.text = None
            p += 2
        # Answers 5-7
        slide = prs.slides.add_slide(prs.slide_layouts[ANSWERS_2])
        slide.placeholders[0].text = f"Round {r} Answers (cont.)"
        p = 10
        for q in range(5, 8):
            question = p
            answer = p + 1
            slide.placeholders[question].text = f"Q{q}: " + data[f"R{r}Q{q}"]["Text"]
            slide.placeholders[answer].text = f"A{q}: " + data[f"R{r}A{q}"]["Text"]
            p += 2


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


# if __name__ == "__main__":
#     template = sys.argv[1]
#     infile = sys.argv[2]
#     outfile = infile[: infile.find(".")]
#     quiz = build_quiz(template, infile)
#     quiz.save(f"{outfile}.pptx")
#     print(f"File saved as {outfile}.pptx")
