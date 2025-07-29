import speech_recognition as sr 
import pandas as pd 
import numpy as np 
from openpyxl import Workbook 
from xlutils.copy import copy 
from xlrd import open_workbook 
import xlrd
from gtts import gTTS
import win32com.client
#import xlrd
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
# get audio from the microphone 
r = sr.Recognizer()
with sr.Microphone() as source: 
	 print("Speak something:") 
	 audio = r.listen(source) 
	 try: 
		 f = open("speechtext.txt", "w") 
		 f.write(r.recognize_google(audio)) 
	 except sr.UnknownValueError: 
			 print("Could not understand audio") 
	 except sr.RequestError as e: 
				 print("Could not request results; {0}".format(e)) 
# Convert Speech to Text 
Mssg = r.recognize_google(audio) 
print("You said " + Mssg)
book = open_workbook(r"C:\Users\ssp1_\OneDrive\Desktop\DESKTOP\destop\NAC DMT\dmt projects\books.xlsx")
for sheet in book.sheets(): 
			for rowidx in range(sheet.nrows): 
				row = sheet.row(rowidx) 
				for colidx, cell in enumerate(row): 
					if cell.value == Mssg : 
						print(sheet.name) 
						print(colidx) 
						print(rowidx) 
						i=rowidx 
						book = xlrd.open_workbook(r'C:\Users\ssp1_\Downloads\books.xls') 
						first_sheet = book.sheet_by_index(0) 
						print("The searched book was found") 
						print("The details are :") 
						print("Book Id Title Author Description Type of book cost rating ") 
						print(first_sheet.row_values(i))
						speaker = win32com.client.Dispatch("SAPI.SpVoice") 
  
						while 1: 
								s = first_sheet.row_values(i) 
								speaker.Speak(s) 
								break