from lessonplancreator import *
from docx import *
from docx.shared import Cm
import sys
from PyQt4 import QtGui, QtCore


def lessonplanfunction():
    
    name = str(input('Class name?'))
    time = str(input('Time?'))
    filename = '\\' + name + '-' + time
    location = 'C:\\Users\Mike\Documents\lessonplan\lessonplancode\document_output'
    extension = '.docx'
    template = 'C:\\Users\Mike\Documents\lessonplan\lessonplancode\document_templates\lessonplantemplate.docx'
    save = location + filename + extension
    
    newclass = LessonPlanCreator(name, time, save, template)
    
    newclass.page_one_creator()
    newclass.page_two_creator()
    newclass.page_one_fill()
    newclass.page_two_fill()
    newclass.savefile()
    
lessonplanfunction()