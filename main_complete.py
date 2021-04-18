""" Docx

author: ashraf minhaj
mail  : ashraf_minhaj@yahoo.com
"""

"""
install -
$ pip install pyqt5
$ pip install docx2txt
"""

import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
import docx2txt


class MainApp(QMainWindow):
    """ the main class of our app """
    def __init__(self):
        """ init things here """
        super().__init__()                         # parent class initializer

        # window title
        self.title = "Google Doc Clone"
        self.setWindowTitle(self.title)
        
        # editor section
        self.editor = QTextEdit(self) 
        self.setCentralWidget(self.editor)

        # create menubar and toolbar first
        self.create_menu_bar()
        self.create_toolbar()

        # after craeting toolabr we can call and select font size
        font = QFont('Times', 12)
        self.editor.setFont(font)
        self.editor.setFontPointSize(12)

        # stores path
        self.path = ''

    def create_menu_bar(self):
        menuBar = QMenuBar(self)

        """ add elements to the menubar """
        # App icon will go here
        app_icon = menuBar.addMenu(QIcon("doc_icon.png"), "icon")

        # file menu **
        file_menu = QMenu("File", self)
        menuBar.addMenu(file_menu)

        save_action = QAction('Save', self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        open_action = QAction('Open', self)
        open_action.triggered.connect(self.file_open)
        file_menu.addAction(open_action)

        rename_action = QAction('Rename', self)
        rename_action.triggered.connect(self.file_saveas)
        file_menu.addAction(rename_action)

        pdf_action = QAction("Save as PDF", self)
        pdf_action.triggered.connect(self.save_pdf)
        file_menu.addAction(pdf_action)
        

        # edit menu **
        edit_menu = QMenu("Edit", self)
        menuBar.addMenu(edit_menu)

        # paste
        paste_action = QAction('Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)

        # clear 
        clear_action = QAction('Clear', self)
        clear_action.triggered.connect(self.editor.clear)
        edit_menu.addAction(clear_action)

        # select all
        select_action = QAction('Select All', self)
        select_action.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(select_action)

        # view menu **
        view_menu = QMenu("View", self)
        menuBar.addMenu(view_menu)

        # fullscreen
        fullscr_action = QAction('Full Screen View', self)
        fullscr_action.triggered.connect(lambda : self.showFullScreen())
        view_menu.addAction(fullscr_action)

        # normal screen
        normscr_action = QAction('Normal View', self)
        normscr_action.triggered.connect(lambda : self.showNormal())
        view_menu.addAction(normscr_action)

        # minimize
        minscr_action = QAction('Minimize', self)
        minscr_action.triggered.connect(lambda : self.showMinimized())
        view_menu.addAction(minscr_action)

        self.setMenuBar(menuBar)

    def create_toolbar(self):
        # Using a title
        ToolBar = QToolBar("Tools", self)

        # undo
        undo_action = QAction(QIcon("undo.png"), 'Undo', self)
        undo_action.triggered.connect(self.editor.undo)
        ToolBar.addAction(undo_action)

        # redo
        redo_action = QAction(QIcon("redo.png"), 'Redo', self)
        redo_action.triggered.connect(self.editor.redo)
        ToolBar.addAction(redo_action)

        # adding separator
        ToolBar.addSeparator()

        # copy
        copy_action = QAction(QIcon("copy.png"), 'Copy', self)
        copy_action.triggered.connect(self.editor.copy)
        ToolBar.addAction(copy_action)

        # cut 
        cut_action = QAction(QIcon("cut.png"), 'Cut', self)
        cut_action.triggered.connect(self.editor.cut)
        ToolBar.addAction(cut_action)

        # paste
        paste_action = QAction(QIcon("paste.png"), 'Paste', self)
        paste_action.triggered.connect(self.editor.paste)
        ToolBar.addAction(paste_action)

        # adding separator
        ToolBar.addSeparator()
        ToolBar.addSeparator()

        # fonts
        self.font_combo = QComboBox(self)
        self.font_combo.addItems(["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"])
        self.font_combo.activated.connect(self.set_font)      # connect with function
        ToolBar.addWidget(self.font_combo) 

        # font size
        self.font_size = QSpinBox(self)   
        self.font_size.setValue(12)  
        self.font_size.valueChanged.connect(self.set_font_size)      # connect with funcion
        ToolBar.addWidget(self.font_size)

        # separator
        ToolBar.addSeparator()

        # bold
        bold_action = QAction(QIcon("bold.png"), 'Bold', self)
        bold_action.triggered.connect(self.bold_text)
        ToolBar.addAction(bold_action)

        # underline
        underline_action = QAction(QIcon("underline.png"), 'Underline', self)
        underline_action.triggered.connect(self.underline_text)
        ToolBar.addAction(underline_action)

        # italic
        italic_action = QAction(QIcon("italic.png"), 'Italic', self)
        italic_action.triggered.connect(self.italic_text)
        ToolBar.addAction(italic_action)

        # separator
        ToolBar.addSeparator()

        # text alignment
        right_alignment_action = QAction(QIcon("right-align.png"), 'Align Right', self)
        right_alignment_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignRight))
        ToolBar.addAction(right_alignment_action)

        left_alignment_action = QAction(QIcon("left-align.png"), 'Align Left', self)
        left_alignment_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignLeft))
        ToolBar.addAction(left_alignment_action)

        justification_action = QAction(QIcon("justification.png"), 'Center/Justify', self)
        justification_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignCenter))
        ToolBar.addAction(justification_action)

        # separator
        ToolBar.addSeparator()

        # zoom in
        zoom_in_action = QAction(QIcon("zoom-in.png"), 'Zoom in', self)
        zoom_in_action.triggered.connect(self.editor.zoomIn)
        ToolBar.addAction(zoom_in_action)

        # zoom out
        zoom_out_action = QAction(QIcon("zoom-out.png"), 'Zoom out', self)
        zoom_out_action.triggered.connect(self.editor.zoomOut)
        ToolBar.addAction(zoom_out_action)


        # separator
        ToolBar.addSeparator()
        
        self.addToolBar(ToolBar)

    def italic_text(self):
        # if already italic, change into normal, else italic
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not(state))

    def underline_text(self):
        # if already underlined, change into normal, else underlined
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))

    def bold_text(self):
        # if already bold, make normal, else make bold
        if self.editor.fontWeight() != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)

    def set_font(self):
        font = self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(font))

    def set_font_size(self):
        value = self.font_size.value()
        self.editor.setFontPointSize(value)

        # we can also make it one liner without writing such function.
        # by using lamba function -
        # self.font_size.valueChanged.connect(self.editor.setFontPointSize(self.font_size.value()))  


    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "", "Text documents (*.text);Text documents (*.txt);All files (*.*)")

        try:
            #with open(self.path, 'r') as f:
            #    text = f.read()
            text = docx2txt.process(self.path) # docx2txt
            #doc = Document(self.path)         # if using docx
            #text = ''
            #for line in doc.paragraphs:
            #    text += line.text
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        print(self.path)
        if self.path == '':
            # If we do not have a path, we need to use Save As.
            self.file_saveas()

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "text documents (*.text);Text documents (*.txt);All files (*.*)")

        if self.path == '':
            return   # If dialog is cancelled, will return ''

        text = self.editor.toPlainText()

        try:
            with open(path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)

    def save_pdf(self):
        f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "PDF files (.pdf);;All files()")
        print(f_name)

        if f_name != '':  # if name not empty
           printer = QPrinter(QPrinter.HighResolution)
           printer.setOutputFormat(QPrinter.PdfFormat)
           printer.setOutputFileName(f_name)
           self.editor.document().print_(printer)
    

app = QApplication(sys.argv)
window = MainApp()
window.show()
sys.exit(app.exec_())