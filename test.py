from lessonplancreator import *
from docx import *
from docx.shared import Cm

name = str(input('Class Name?'))
time = str(input('Time?'))
filename = '\\' + name
location = 'C:\\Users\Mike\Documents\Python\Creating text'
extension = '.docx'
save = location + filename + extension

newclass = LessonPlanCreator(name, time, save, 'lessonplanedt.docx')

newclass.page_one_creator()
newclass.page_two_creator()
newclass.page_one_fill()
newclass.page_two_fill()
newclass.savefile()
