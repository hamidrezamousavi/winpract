from win32com.client import Dispatch
import os

class WordWarp:
    def __init__(self,templatefile=None):
        self.wordapp = Dispatch('Word.Application')
        if templatefile == None:
            self.wordoc = self.wordapp.Documents.Add()
        else:
            self.wordoc = self.wordapp.Documents.Add(templatefile)
        self.wordoc.Range(0,0).Select()
        self.wordsel = self.wordapp.selection
        # self.getStyleDictionary()

    def show(self):
        self.wordapp.Visible = 1

    def save_as(self,filename):
        self.worddoc.SaveAs(filename)

    def printout(self):
        self.wordoc.PrintOut()

    def select_end(self):
        self.wordsel.Collapse(0)

    def add_text(self,text):
        self.wordsel.InsertAfter(text)
        self.select_end()

filename = os.getcwd()+'\\t.docx'
mydoc = WordWarp(filename)
mydoc.show()
mydoc.add_text("jhgsadfjhgaskdfj")



