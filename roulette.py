#!/usr/bin/python3

from pptx import Presentation
from glob import glob
import random


presentation = Presentation('template_presentation.pptx')

title_layout = presentation.slide_layouts[0]
blank_layout = presentation.slide_layouts[6]

first_slide = presentation.slides.add_slide(title_layout)
second_slide = presentation.slides.add_slide(blank_layout)

first_slide.shapes.title.text = "Random topic title!"
first_slide.placeholders[1].text = "Powerpoint Roulette"

# Select a random image
image_filenames = glob('images/*')
image_filename = random.choice(image_filenames)

# Will strecth or shrink the image to the slide size
pic = second_slide.shapes.add_picture(image_filename, 0, 0,
                                      presentation.slide_width, presentation.slide_height)

presentation.save("random_presentation.pptx")
