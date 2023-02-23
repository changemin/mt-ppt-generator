import pprint
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os, random

img_file_names = []
quizes = []


def get_img_list():
    global img_file_names
    for filename in os.listdir("four-letter-game/data/imgs"):
        if filename not in img_file_names:
            img_file_names.append(filename)


def read_data():
    global quizes
    filename = "four-letter-game/data/data.txt"
    f = open(filename, "r", encoding="utf-8")

    lines = f.readlines()
    for line in lines:
        quizes.append(line)

def merge_quizes():
    global quizes
    quizes += img_file_names
    random.shuffle(quizes)

def ppt_gen():
    prs = Presentation()
    for quiz in list(quizes):
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)

        if quiz[-3:] == 'png' or quiz[-3:] == 'jpg':
            try:
                pic = slide.shapes.add_picture("four-letter-game/data/imgs/"+quiz, Inches(1), Inches(1), height=Inches(5))
            except:
                print(quiz)
            pic.left = int((prs.slide_width - pic.width) / 2)
            pic.top = int((prs.slide_height - pic.height) / 2)
            quizes.remove(quiz)

        else:
            # add 4글자 문제
            txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
            tf = txBox.text_frame
            tf.text = quiz[0:2]
            
            tf.paragraphs[0].font.size = Pt(100)
            tf.paragraphs[0].font.bold = True

            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            txBox.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            quizes.remove(quiz)  
        
    
    prs.save("4글자 이어말하기+인물퀴즈.pptx")



read_data()
get_img_list()

merge_quizes()
print("총", len(img_file_names),"개의 인물사진")
print("총", len(quizes),"개의 Slides")
ppt_gen()


