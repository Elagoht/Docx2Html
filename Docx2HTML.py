from PyQt5.QtWidgets import QApplication, QLineEdit, QMainWindow, QSpinBox, QRadioButton, QWidget, QGroupBox, QPushButton, QCheckBox, QFileDialog, QLabel, QGridLayout
from PyQt5.QtGui import QIcon, QKeySequence
import sys
import mammoth
from bs4 import BeautifulSoup as bs
from io import StringIO as io
import re

messages={
    "select":"Select a file to convert.",
    "ready":"Ready to convert.",
    "success":"Successfully converted into {}.",
    "unselected":"Output file not selected.",
    "notdocx":"File extension must be docx."
}

def prettify(self,encoding=None,formatter="minimal",indent_width=4):
    return re.compile(r"^(\s*)",re.MULTILINE).sub(r"\1"*indent_width,bs.prettify(self,encoding,formatter))

def convert(input:str,output:str,oneline:bool,pretty:bool,indent:int):
    with open(input,"rb") as file:
        html=mammoth.convert_to_html(file).value
        if not oneline:
            html=html.replace("><",">\n<")
    with open(output,"w",encoding="UTF-8") as f:
        f.write(prettify(bs(io(html),"html.parser"),indent_width=indent) if pretty else html)

class MainWin(QMainWindow):
    def __init__(self):
        super(MainWin,self).__init__()
        self.show()
        self.setFixedSize(360,290 if "win" in sys.platform else 320)
        self.central=Central()
        self.setCentralWidget(self.central)

class Central(QWidget):
    def __init__(self):
        super(Central,self).__init__()
        self.layout=QGridLayout(self)

        self.boxFileSelect=QGroupBox("File Selection",self)
        self.layFileSelect=QGridLayout(self.boxFileSelect)

        self.boxSettings=QGroupBox("Settings",self)
        self.laySettings=QGridLayout(self.boxSettings)

        self.lDescription=QLabel("Pick or drag and drop a file to start.")
        self.lDescription.setWordWrap(True)
        self.bSelect=Button("Select File",self)
        self.eFileName=QLineEdit("Not Selected.")
        self.eFileName.setDisabled(True)
        self.eFileName.setStyleSheet("color:black;")

        self.layFileSelect.addWidget(self.lDescription,0,0,1,3)
        self.layFileSelect.addWidget(self.bSelect,1,0)
        self.layFileSelect.addWidget(self.eFileName,1,1,1,2)

        self.rOneLine=QRadioButton("One-line",self)
        self.rMultiLine=QRadioButton("Multi-line",self)
        self.cIndent=QCheckBox("Add indent",self)
        self.lIndent=QLabel("Indent Spaces",self)
        self.sIndent=QSpinBox(self)
        self.sIndent.setRange(1,8)
        self.sIndent.setValue(4)
        self.sIndent.setFixedWidth(35)

        self.laySettings.addWidget(self.rOneLine,0,0,1,10)
        self.laySettings.addWidget(self.rMultiLine,1,0,1,10)
        self.laySettings.addWidget(self.cIndent,2,1,1,2)
        self.laySettings.addWidget(self.lIndent,3,1)
        self.laySettings.addWidget(self.sIndent,3,2)

        self.bConvert=QPushButton("Convert",self)
        self.lStatus=QLabel()

        self.layout.addWidget(self.boxFileSelect,0,0,1,3)
        self.layout.addWidget(self.boxSettings,1,0,1,3)
        self.layout.addWidget(self.bConvert,2,1)
        self.layout.addWidget(self.lStatus,3,0,1,3)

        self.file=""

        self.rMultiLine.setChecked(True)
        self.cIndent.setChecked(True)

        self.bSelect.clicked.connect(self.selectFile)
        self.bSelect.setShortcut(QKeySequence("Ctrl+O"))
        self.bConvert.clicked.connect(self.saveConverted)
        self.bConvert.setShortcut(QKeySequence("Ctrl+Enter"))
        self.rOneLine.clicked.connect(self.lineDecided)
        self.rMultiLine.clicked.connect(self.lineDecided)
        self.cIndent.clicked.connect(self.isIndented)
        self.eFileName.textChanged.connect(self.canConvert)

        self.canConvert()

    def selectFile(self):
        file=QFileDialog.getOpenFileName(self,"Select DOCX File","","Microsoft Word Document (*.docx)",options=QFileDialog.DontUseNativeDialog)
        self.file=file[0]
        self.eFileName.setText(file[0].split("/")[-1] if file[0] else "Not Selected.")

    def saveConverted(self):
        exportFile=QFileDialog.getSaveFileName(self,"Save Converted HTML File",filter="Hyper Text Markup Language (*.html)",options=QFileDialog.DontUseNativeDialog)[0]
        if exportFile!="":
            if exportFile[-5:]!=".html":
                exportFile+=".html"
            oneLine=self.rOneLine.isChecked()
            convert(self.file,exportFile,oneLine,self.cIndent.isChecked() if not oneLine else False,self.sIndent.value())
            self.lStatus.setText(messages["success"].format(exportFile.split("/")[-1]))
        else:
            self.lStatus.setText(messages["unselected"])

    def isIndented(self):
        self.sIndent.setDisabled(not self.cIndent.isChecked())

    def lineDecided(self):
        state=self.rOneLine.isChecked()
        self.cIndent.setDisabled(state)
        self.lIndent.setDisabled(state)
        if state:
            self.sIndent.setDisabled(True)
        else:
            self.isIndented()

    def canConvert(self):
        state=self.eFileName.text()=="Not Selected."
        self.bConvert.setDisabled(state)
        if state:
            self.lStatus.setText(messages["select"])
        else:
            self.lStatus.setText(messages["ready"])

class Button(QPushButton):
    def __init__(self, title, parent):
        super().__init__(title, parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self,e):

        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self,e):
        print(e.mimeData().text().strip())
        if e.mimeData().text().strip()[-5:]==".docx":
            main.central.eFileName.setText(e.mimeData().text().strip().split("/")[-1])
            main.central.file=e.mimeData().text().strip().split("//"+("/" if "win" in sys.platform else ""))[1]
        else:
            main.central.lStatus.setText(messages["notdocx"])
app=QApplication(sys.argv)
app.setApplicationName("Docx2Html Convertor")
app.setStyle("Fusion")
app.setWindowIcon(QIcon("docx2html.png"))
main=MainWin()
sys.exit(app.exec_())