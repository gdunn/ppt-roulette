#!/usr/bin/python3

from pptx import Presentation


presentation = Presentation()

title_layout = presentation.slide_layouts[0]
blank_layout = presentation.slide_layouts[6]

first_slide = presentation.slides.add_slide(title_layout)
second_slide = presentation.slides.add_slide(blank_layout)

first_slide.shapes.title.text = "Random topic title!"
first_slide.placeholders[1].text = "Powerpoint Roulette"

presentation.save("First_presentation.pptx")
