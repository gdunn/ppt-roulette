#!/usr/bin/python3

from pptx import Presentation
from glob import glob
import random


SLIDES_COUNT = 10


presentation = Presentation('template_presentation.pptx')

title_layout = presentation.slide_layouts[0]
blank_layout = presentation.slide_layouts[6]

first_slide = presentation.slides.add_slide(title_layout)


first_slide.shapes.title.text = "Random topic title!"
first_slide.placeholders[1].text = "Powerpoint Roulette"

# Select a random images
all_image_filenames = glob('images/*')
image_filenames = random.sample(all_image_filenames, k=SLIDES_COUNT)

for image_filename in image_filenames:
    image_slide = presentation.slides.add_slide(blank_layout)

    # Will strecth or shrink the image to the slide size
    pic = image_slide.shapes.add_picture(
        image_filename, 0, 0, 
        presentation.slide_width, presentation.slide_height)

last_slide = presentation.slides.add_slide(title_layout)
last_slide.shapes.title.text = "Random topic title!"
last_slide.placeholders[1].text = "End. Well done!"

presentation.save("random_presentation.pptx")
