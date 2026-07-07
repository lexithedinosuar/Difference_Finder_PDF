import os
from pathlib import Path

class paragraph:
    def __init__(self, textFile):
        self.file = open(textFile, "rt", encoding = "utf-8", errors="replace")
        self.text = self.file.read()

    def splitParagraph(self):
        paragraphs = self.text.split("\n")
        index=0
        for p in paragraphs:
            if p == "":
                paragraphs.pop(index)
            index += 1
        return paragraphs
    
    def splitWords(self):
        words = self.split()
        return words
    
    def splitEverything(self):
        nestedWords = []
        paragraphList=self.splitParagraph()

        for p in paragraphList:
            wordsList=p.split()
            nestedWords.append(wordsList)

        return nestedWords



