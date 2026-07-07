"""
Creator: Alexia Piunno
Title: Difference Finder
Purpose: takes two txt files of different verions of the same document and compares all
text within to find items that are different in each file. putting the text back into a word
document with different formatting. text is compaired paragraph by paragraph, then word by word.  
 
input: two txt files
Output: new Docx file with different formatting to represent text comparison
 - black text = text that is the same in both files
 - red text (old word/new word) = text that changes
 - Blue text = additional text (not reviewed because there was nothing to compair to)
 - Green text = number of errors found in the comparison

Next steps:
 - compair sentence by sentence, page by page 
    - straigth from word file?
        - pdf split page by page, each pdf converted to text, iterate through all page files?
 - find matching sentence before comparing text
        -use list indexing methods?
 - add a new function that checks that the word document is not open in word before editing it.
 - make new functions to make future codeing easier.
    -find longest/shotest lists
    -matching/extra list items and doc formating
 - check the page/paragraph/sentence before checking the words to reduce run time.
"""
# import Libraries/modules
from wordDocClass import new_document
from paragraphClass import paragraph
from docx import Document
from docx.shared import Pt
import os
from PDFsplitter import *
from pathlib import Path

FileName = "EDS0027DifferenceReport"
fullFileName = FileName + ".docx"
diffDoc = new_document("0028 EDS Difference report", FileName)
folder_1 = FullPDF_to_SplitTXT("0028_V1.pdf", "EDS_0027_V1_folder")
folder_2 = FullPDF_to_SplitTXT("0028_V2.pdf", "EDS_0027_V2_folder")

Files_1 = [file for file in folder_1.iterdir() if file.is_file()]
Files_2 = [file for file in folder_2.iterdir() if file.is_file()]

print(Files_1)
print(Files_2)

longest_folder = 0
shortest_folder = 0
longest_folder_len = 0

if len(Files_1)>=len(Files_2):
    longest_folder=Files_1
    shortest_folder =Files_2
    longest_folder_len =len(Files_1)
else:
    longest_folder=Files_2
    shortest_folder =Files_1
    longest_folder_len =len(Files_2)


f_index = 0
while f_index < longest_folder_len:
    
    if f_index < len(shortest_folder):
        diffDoc.addHeading("Page" + str(f_index+1))
        # variable definition
        f1 = paragraph(Files_1[f_index])
        f2 = paragraph(Files_2[f_index])

        longest_p_list = 0
        shortest_p_list = 0
        longest_p_len = 0
        ErrorCount = 0

        # split test into nested list [[words, in, paragraph, 1.], [words, in paragraph, 2.], ... ,[words, in, paragraph, n.]]
        f1_list = f1.splitEverything()
        f2_list = f2.splitEverything()


        # find Longest list of Paragraphs
        if len(f1_list) >= len(f2_list):
            longest_p_len=len(f1_list)
            longest_p_list = f1_list
            shortest_p_list = f2_list
        else:
            longest_p_len=len(f2_list)
            longest_p_list = f2_list
            shortest_p_list = f1_list

        # iterate through paragraphs
        p_index = 0
        while p_index < longest_p_len:
            longest_w_list = 0
            shortest_w_list = 0
            w_list_len = 0
            w_index = 0
            
            # prevent indexing error, compare test in paragraphes with index available in both list
            if p_index < len(shortest_p_list):

                f1_paragraph_words = f1_list[p_index]
                f2_paragraph_words = f2_list[p_index]
                
                # start a new paragraph on Difference Report Doc
                diffDoc.newParagraph("")

                # find the longest word list
                if len(f1_paragraph_words) >= len(f2_paragraph_words):
                    longest_w_list = f1_paragraph_words
                    shortest_w_list = f2_paragraph_words
                    w_list_len = len(f1_paragraph_words)
                else:
                    longest_w_list = f2_paragraph_words
                    shortest_w_list = f1_paragraph_words
                    w_list_len = len(f2_paragraph_words)

                # iterate through words
                while w_index < w_list_len:
                    # prevent indexing error: compare words with the same index, add all extra words in blue.
                    if w_index < len(shortest_w_list):
                        w1 = f1_paragraph_words[w_index]
                        w2 = f2_paragraph_words[w_index]
                        
                        #formate matching/different words 
                        if w1 == w2:
                            diffDoc.add2Paragraph(w1 + " ")
                        else:
                            diffDoc.addRedText(w1 + "/" + w2 + " ")
                            ErrorCount += 1
                    else:
                        extra_word = longest_w_list[w_index]
                        diffDoc.addBlueText(extra_word + " ")
                        ErrorCount += 1
                
                    w_index += 1
            else: # all paragraphs in the longer list get added in blue without comparison
                diffDoc.newParagraph("")
                diffDoc.addBlueText(longest_p_list[p_index])
                ErrorCount += 1

            p_index += 1

        # start a new page
        diffDoc.newParagraph("")
        diffDoc.addGreenText("Number of Errors: " + str(ErrorCount))
        diffDoc.addPageBreak()
        print("Page: "+ str(f_index+1)+ "--> Number of errors: " + str(ErrorCount))
    else: 
        ErrorCount+=1
        page_text = read_txt_file(longest_folder[f_index])
        diffDoc.addHeading("Page" + str(f_index+1))
        diffDoc.addBlueText(page_text)
        diffDoc.newParagraph("")
        diffDoc.addGreenText("Number of Errors: " + str(ErrorCount))
        diffDoc.addPageBreak()
        print("Page: "+ str(f_index+1)+ "--> Number of errors: " + str(ErrorCount))
        pass
    
    f_index += 1



#Save the document for the difference report
diffDoc.saveFile()
#open the file in word
os.startfile(fullFileName)

print("You have reached the end of your script!")