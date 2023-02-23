import pprint
from pptx import Presentation

quizes = []

def read_data():
    global quizes
    filename = "data.txt"
    f = open(filename, "r", encoding="utf-8")

    lines = f.readlines()
    for line in lines:
        quiz = line.split(". ")[1]
        quizes.append(quiz)
    
def ppt_gen():
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])

    # add some text to the slide
    title = slide.shapes.title
    title.text = "My Title"
    subtitle = slide.placeholders[1]
    subtitle.text = "My Subtitle"

    # save the presentation to a file
    prs.save("output.pptx")

read_data()
print(quizes)