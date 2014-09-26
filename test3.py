
from lessonplancreator import *
from docx import *
from docx.shared import Cm
import sys
from PyQt4 import QtGui, QtCore



class ComboBoxBasic(QtGui.QWidget):
    """
    An basic example combo box application
    """

    def __init__(self):
        # create GUI
        QtGui.QMainWindow.__init__(self)
        self.setWindowTitle('Combo Box Basic')
        # Set the window dimensions
        self.resize(250,50)
        
        # vertical layout for widgets
        self.vbox = QtGui.QVBoxLayout()
        self.setLayout(self.vbox)

        # Create a combo box and add it to our layout
        self.namecombo = QtGui.QComboBox()
        self.vbox.addWidget(self.namecombo)
        self.timecombo = QtGui.QComboBox()
        self.vbox.addWidget(self.timecombo)        
        self.button = QtGui.QPushButton('OK')
        self.vbox.addWidget(self.button)

        # A label to display our selection
        self.lbl = QtGui.QLabel('Class')
        
        # Center align text
        self.lbl.setAlignment(QtCore.Qt.AlignHCenter)
        self.vbox.addWidget(self.lbl)

        # Add lists to comboboxes
        classname = ['Select Class...', '10YS', '7YS', '8XS']
        self.namecombo.addItems(classname)
        time = ['Select Time...', '0900h', '1000h', '1130h', '1300h', '1400h']
        self.timecombo.addItems(time)
                       
        # Connect the activated signal on the combo box to our handler.
        # This is an overloaded signal, meaning there are variants of it, for
        # example the activated(int) variant emits the index of the chosen
        # option, rather than it's text
        self.namecombo.currentIndexChanged[str].connect(self.namecombo_chosen)
        self.timecombo.currentIndexChanged[str].connect(self.timecombo_chosen)

    def namecombo_chosen(self, text):
        
        self.name = text
        self.button.clicked.connect(self.final)
        
    def timecombo_chosen(self, text):
                
        self.time = text
        self.button.clicked.connect(self.final)    
        
    def final(self):
        
        self.template = 'C:\\Users\Mike\Documents\lessonplan\lessonplancode\document_templates\lessonplantemplate.docx'
        self.location = 'C:\\Users\Mike\Documents\lessonplan\lessonplancode\document_output\\'
        self.extension = '.docx'
        self.save = self.location + self.name + '-' + self.time + self.extension
        self.newclass = LessonPlanCreator(self.name, self.time, self.save, self.template)
        self.newclass.page_one_creator()
        self.newclass.page_two_creator()
        self.newclass.page_one_fill()
        self.newclass.page_two_fill()
        self.newclass.savefile()        
        self.close()


# If the program is run directly or passed as an argument to the python
# interpreter then create a ComboBoxBasic instance and show it
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    gui = ComboBoxBasic()
    gui.show()
    app.exec_()
    
    
