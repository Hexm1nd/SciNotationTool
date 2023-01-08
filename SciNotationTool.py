import sys
from os import path
import win32com.client as client
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt

from ui_MainWindow import Ui_MainWindow

#Constants
wdFindContinue = 1
wdCharacter = 1
wdToggle = 9999998
wdCollapseEnd = 0

class EtoSciApp(QMainWindow):
    def __init__(self):
        super(EtoSciApp, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.DoBtn.clicked.connect(self.EtoSciNotation)
        self.ui.OnTopCheckBox.stateChanged.connect(self.SetOnTop)
        self.ui.DocumentsList.currentRowChanged.connect(self.ShowDetails)
        app.focusWindowChanged.connect(self.on_FocusChange)
        
        self.word = client.Dispatch("Word.Application")
        self.DocList = self.word.Documents
        self.FindDocs()        

    def on_FocusChange(self):
        if self.isActiveWindow():
            self.FindDocs()

    def FindDocs(self):
        self.ui.DocumentsList.blockSignals(True)
        self.ui.FNameLabel.setText("...")
        self.ui.PathLabel.setText("...")
        self.ui.DocumentsList.clear()
        if self.DocList:
            for doc in self.DocList:
                self.ui.DocumentsList.addItem(path.splitext(doc.Name)[0])
        self.ui.DocumentsList.blockSignals(False)

    def ShowDetails(self):
        selnum = self.ui.DocumentsList.currentRow()
        self.ui.FNameLabel.setText(self.DocList[selnum].Name)
        self.ui.PathLabel.setText(self.DocList[selnum].FullName)
        pass
        

    def EtoSciNotation(self):        
        selnum = self.ui.DocumentsList.currentRow()
        self.word.Documents[selnum].Activate()
        self.word.Selection.Find.ClearFormatting()
        self.word.Selection.WholeStory()

        #First of all trying to find sequences:  (digit)E+(any number of digits)
        #                                        (digit)E-(any number of digits)
        #                                        (digit)E(any number of digits)
        #using a regular expression. Case insensitive.
        sel = self.word.Selection.Find
        sel.Text = "[0-9][Ee][-+0-9]{1;}"
        sel.Replacement.Text = ""
        sel.Forward = True
        sel.Wrap = wdFindContinue
        sel.Format = False
        sel.MatchCase = False
        sel.MatchWholeWord = False
        sel.MatchWildcards = True
        sel.MatchSoundsLike = False
        sel.MatchAllWordForms = False

        sel.Execute()
        while sel.Found:
            #When found the sequence is selected
            #Reducing the selection by 1 left symbol therefore deselecting the last digit before symbol E
            self.word.Selection.MoveStart(wdCharacter, 1)

            #Deleting the symbol E. Case insensitive
            self.word.Selection.Text = self.word.Selection.Text.upper().replace( "E", "")

            #Converting the selection into float and back to string. Therefore deleting all heading zeros and symbol +
            self.word.Selection.Text = float(self.word.Selection.Text)

            #Converting selected text into superscript            
            self.word.Selection.Font.Superscript = wdToggle

            self.word.Selection.InsertBefore ("·10")
            self.word.Selection.Collapse (wdCollapseEnd)

            #Trying to find next sequence
            sel.Execute()


    def SetOnTop(self):         
        flags = self.windowFlags()
        flags |= Qt.CustomizeWindowHint            
        if self.ui.OnTopCheckBox.checkState():
            flags |= Qt.WindowStaysOnTopHint    
        else:
            flags &= ~Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.show()
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EtoSciApp()
    window.show()
    sys.exit(app.exec())