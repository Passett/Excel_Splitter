#Import dependencies
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox


#Functions to be used in the tkinter gui below, triggered by button clicks

#Create functions to get the folderpath from the 'Browse' buttons
def getFolderPath():
	folder_selected = filedialog.askdirectory()
	folderPath.set(folder_selected)


#function to split the workbooks
def split():

	#Get folderpath for documents that need to be split
	folder = folderPath.get()

	#Create a new folder to put the results in
	newfolder=folder+"\\"+"Split_Results"
	os.mkdir(newfolder)

	#Create a list of files that need to be split
	files=os.listdir(folder)

	#Filter out any files that are not .xlsx files
	files_xlsx=[f for f in files if f[-4:]=='xlsx']

	#Loop over the list of files, adding each file to a pandas dataframe
	for file in files_xlsx:
		xl=pd.ExcelFile(folder+"\\"+file)

		#loop over each sheet in each file and save each sheet to your newfolder using the naming convention old-file-name_sheet-name
		for sheet in xl.sheet_names:
			df=pd.read_excel(xl,sheet_name=sheet)
			fname= newfolder+"\\"+file[:-5]+"_{}.xlsx".format(sheet)
			with pd.ExcelWriter(fname) as writer:
				df.to_excel(writer, sheet_name='Sheet1', index=False)


	#Update tkinter status
	New_Status='All files have been successfully split. Results are in '+newfolder
	Status_Update.set(New_Status)


#Tkinter GUI for users, it should ask for folderpath and have a go button


#Create a GUI window
window = tk.Tk()
window.geometry('550x400')
window.configure(bg="blue")
window.title("Excel Workbook Splitter")

#Set data from entry boxes to variables
folderPath=tk.StringVar()
Status_Update=tk.StringVar()

#Create the top label
label = tk.Label(text="Enter the filepath to your documents and click 'Split'", bg='blue', fg='white')
label.grid(row=0,column=0, pady=10, padx=10, sticky='W')

#Create the entry box and focus it
entry1=tk.Entry(window, width=65, textvariable=folderPath)
entry1.grid(row=1,column=0, padx=10, sticky='W')
entry1.focus()

#Create the 'Browse Folder' button
btnFind = tk.ttk.Button(width=15, text="Browse Folder",command=getFolderPath)
btnFind.grid(row=1,column=1, sticky='W', padx=5)

#Create 'Split' button
button1=tk.Button(width=15, text='Split', command =split)
button1.grid(row=2,column=0, padx=10, sticky='W', pady=5)

#Create a button to close the program
button2=tk.Button(width=15, text='Close Program', command=window.quit)
button2.grid(row=2, column=0)

#Create the status label
Status_label=tk.Label(textvariable=Status_Update, bg='blue', fg='white')
Status_label.grid(row=6, column=0, columnspan=20, padx=10, pady=10, sticky='W')

#Run the gui
window.mainloop()