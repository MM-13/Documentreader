#modules
from gtts import gTTS

import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

import docx2txt
import PyPDF2

import os
import time

#created
print('Created by MM-13')
print('Visit https://github.com/MM-13 for more programs!')

#language list
print('\nLanguages:')
print('- nl = Dutsch') 
print('- en = English')
print('- fr = French')
print('- de = Deutsch')
print('- es = Spanisch')
print('- ...')

#input Language
taal = simpledialog.askstring("Language", "Languege:")
print('\n- {}'.format(taal))
time.sleep(1.5)

#open explorertab
exp = tk.Tk()
exp.withdraw()

file = filedialog.askopenfilename()

#filetype
extension = os.path.splitext(file)[1]

#filetype process
if extension == '.docx':
    txt = docx2txt.process('{}'.format(file))
    # Language in which you want to convert
    language = ('{}'.format(taal))
    # Passing the text and language to the engine
    print('\n- converting file...')
    myobj = gTTS(text=txt, lang=language, slow=False)
    # Saving the converted audio in a mp3 file named
    print('\n- saving...')
    myobj.save("FileReader.mp3")
    print('\n- Conversion compleet!')
    
elif extension == '.txt':
    with open('{}'.format(file), 'r') as f:
        txt = f.read()
        # Language in which you want to convert
        language = ('{}'.format(taal))
        # Passing the text and language to the engine
        print('\n- converting file...')
        myobj = gTTS(text=txt, lang=language, slow=False)
        # Saving the converted audio in a mp3 file named
        print('\n- saving...')
        myobj.save("FileReader.mp3")
        print('\n- Conversion compleet!')

elif extension == '.pdf':
    #Open file Path
    pdf_File = open('{}'.format(file), 'rb') 

    #Create PDF Reader Object
    pdf_Reader = PyPDF2.PdfFileReader(pdf_File)
    count = pdf_Reader.numPages # counts number of pages in pdf
    textList = []

    #Extracting text data from each page of the pdf file
    for i in range(count):
        try:
            page = pdf_Reader.getPage(i)    
            textList.append(page.extractText())
        except:
            pass
    #Converting multiline text to single line text
    txt = " ".join(textList)
    # Language in which you want to convert
    language = ('{}'.format(taal))
    # Passing the text and language to the engine
    print('\n- converting file...')
    myobj = gTTS(text=txt, lang=language, slow=False)
    # Saving the converted audio in a mp3 file named
    print('\n- saving...')
    myobj.save("FileReader.mp3")
    print('\n- Conversion compleet!')

time.sleep(2)