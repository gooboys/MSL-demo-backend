from pptx import Presentation
from pptx.slide import Slide, Slides, SlideLayout
import os

def pptx_maker(input):
  data = input["data"]
  firstSlide = data[0]["output"]
  secondSlide = data[1]["output"]
  thirdSlide = data[2]["output"]
  return [firstSlide, secondSlide, thirdSlide]