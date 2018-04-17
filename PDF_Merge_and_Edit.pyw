import PyPDF2
import os, time
import tkinter as tk
from tkinter import filedialog

#Written by Simon Wong
#https://github.com/simonwongwong

def popup(text):
	textwindow = tk.Tk()
	textwindow.title('Error!')
	textwindow.minsize(width=300, height=50)
	label = tk.Label(textwindow, text=text)
	label.pack()
	label.configure(pady=20)
	textwindow.mainloop()

def finished(file, operation, window):
	finishPrompt = tk.Tk()
	finishPrompt.title(operation + " completed!")
	tk.Label(finishPrompt, text=operation + " finished. Would you like to open the file now?").grid(row=0,column=0, columnspan=2, pady=5, padx=5)
	tk.Button(finishPrompt, text="Open file", command=lambda: os.startfile(file)).grid(row=1, column=0, pady=5, padx=5, sticky=stickyFill)
	tk.Button(finishPrompt, text="Finished", command= lambda: finishPrompt.destroy()).grid(row=1,column=1, pady=5, padx=5, sticky=stickyFill)
	finishPrompt.mainloop()


def filePicker(entry, window):
	file = filedialog.askopenfilename(title="Choose PDF file",filetypes=(("PDF","*.pdf"),),initialdir=os.getcwd())
	
	if entry.get() == "":
		entry.insert(0, file)
	else:
		entry.delete(0, 'end')
		entry.insert(0, file)

	window.lift()


def merge():

	mergeWindow = tk.Tk()
	mergeWindow.title("PDF merger")

	tk.Label(mergeWindow, text="Creates a new PDF file with the second file added to the end of the first file").grid(row=0,column=0, columnspan=3, padx=10, pady=3, sticky=stickyFill)

	tk.Label(mergeWindow, text="Select first PDF file:").grid(row=1, column=0, padx=10, pady=3)
	firstFile = tk.Entry(mergeWindow)
	firstFile.grid(row=1, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(mergeWindow, text="Browse...", command=lambda entry=firstFile, window=mergeWindow: filePicker(entry, window)).grid(row=1, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(mergeWindow, text="Select second PDF file:").grid(row=2, column=0, padx=10, pady=3)
	secondFile = tk.Entry(mergeWindow)
	secondFile.grid(row=2, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(mergeWindow, text="Browse...", command=lambda entry=secondFile, window=mergeWindow: filePicker(entry, window)).grid(row=2, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(mergeWindow, text="Desired name of merged file:").grid(row=3, column=0)
	mergedFile = tk.Entry(mergeWindow)
	mergedFile.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky=stickyFill)

	tk.Button(mergeWindow, text="Merge!", command=lambda: mergeWindow.quit()).grid(row=4,column=0, columnspan=3, padx=5, pady=10, sticky=stickyFill)

	mergeWindow.mainloop()

	firstFile = checkExist(firstFile.get())
	secondFile = checkExist(secondFile.get())
	mergedFile = mergedFile.get()
	

	# print("Creates a new PDF file with the second file added to the end of the first file")
	# firstFile = checkExist(input("Enter the first file name (without the .pdf): "))
	# secondFile = checkExist(input("Enter the second file name (without the .pdf): "))
	# mergedFile = input("Enter the desired name of the merged file (must be different from other files): ")

	mergedBook = PyPDF2.PdfFileMerger()
	mergedBook.append(firstFile)
	mergedBook.append(secondFile)

	fullbook = open(mergedFile+'.pdf','wb')
	mergedBook.write(fullbook)
	mergedBook.close()
	fullbook.close()
	mergeWindow.destroy()
	finished(mergedFile+".pdf", "Merge", mergeWindow)


def pageUpdate():
	updaterWindow = tk.Tk()
	updaterWindow.title("PDF page updater")

	tk.Label(updaterWindow, text="Updates a single page inside an existing PDF").grid(row=0,column=0, columnspan=3, padx=10, pady=3, sticky=stickyFill)

	tk.Label(updaterWindow, text="Select PDF file to edit:").grid(row=1, column=0, padx=10, pady=3)
	updateFile = tk.Entry(updaterWindow)
	updateFile.grid(row=1, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(updaterWindow, text="Browse...", command=lambda entry=updateFile, window=updaterWindow: filePicker(entry, window)).grid(row=1, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(updaterWindow, text="Page to update:").grid(row=2,column=0, padx=10, pady=3)
	pageToUpdate = tk.Entry(updaterWindow)
	pageToUpdate.grid(row=2, column=1, sticky=stickyFill, pady=5, padx=5)

	tk.Label(updaterWindow, text="Select PDF file with updated page:").grid(row=3, column=0, padx=10, pady=3)
	updatedPage = tk.Entry(updaterWindow)
	updatedPage.grid(row=3, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(updaterWindow, text="Browse...", command=lambda entry=updatedPage, window=updaterWindow: filePicker(entry, window)).grid(row=3, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(updaterWindow, text="Page number of updated page:").grid(row=4,column=0, padx=10, pady=3)
	pageWithUpdate = tk.Entry(updaterWindow)
	pageWithUpdate.grid(row=4, column=1, sticky=stickyFill, pady=5, padx=5)
	pageWithUpdate.insert(0, "1")

	tk.Button(updaterWindow, text="Update!", command=lambda: updaterWindow.quit()).grid(row=5,column=0, columnspan=3, padx=5, pady=10, sticky=stickyFill)

	updaterWindow.mainloop()

	filename = updateFile.get()
	filename = filename[:-4] + '-updated.pdf'
	updateFile = checkExist(updateFile.get())
	pageToUpdate = int(pageToUpdate.get())
	updatedPage = checkExist(updatedPage.get())
	pageWithUpdate = int(pageWithUpdate.get())


	if pageToUpdate == 0 or pageWithUpdate == 0:
		popup("invalid page number, must be greater than 0")

	originalPDF = PyPDF2.PdfFileReader(updateFile)
	updatedPagePDF = PyPDF2.PdfFileReader(updatedPage)

	updatedPDF = PyPDF2.PdfFileWriter()
	updatedPDF.cloneDocumentFromReader(originalPDF)
	try:
		updatedPDF.insertPage(updatedPagePDF.getPage(pageWithUpdate-1),pageToUpdate-1)
	except IndexError:
		popup("Please check if page number is within range")
	
	outputFile = open(filename,'wb')

	pdfOut = PyPDF2.PdfFileWriter()

	for i in range(updatedPDF.getNumPages()):
		if i != pageToUpdate:
			pdfOut.addPage(updatedPDF.getPage(i))

	pdfOut.write(outputFile)
	outputFile.close()

	updaterWindow.destroy()
	finished(filename, "Page update", updaterWindow)

def insertPage():
	inserterWindow = tk.Tk()
	inserterWindow.title("PDF page inserter")

	tk.Label(inserterWindow, text="Inserts a single page inside an existing PDF").grid(row=0,column=0, columnspan=3, padx=10, pady=3, sticky=stickyFill)

	tk.Label(inserterWindow, text="Select PDF file to edit:").grid(row=1, column=0, padx=10, pady=3)
	updateFile = tk.Entry(inserterWindow)
	updateFile.grid(row=1, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(inserterWindow, text="Browse...", command=lambda entry=updateFile, window=inserterWindow: filePicker(entry, window)).grid(row=1, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(inserterWindow, text="Page number where new page will be inserted:").grid(row=2,column=0, padx=10, pady=3)
	pageToInsert = tk.Entry(inserterWindow)
	pageToInsert.grid(row=2, column=1, sticky=stickyFill, pady=5, padx=5)

	tk.Label(inserterWindow, text="Select PDF file with page to be inserted:").grid(row=3, column=0, padx=10, pady=3)
	fileWithInsert = tk.Entry(inserterWindow)
	fileWithInsert.grid(row=3, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(inserterWindow, text="Browse...", command=lambda entry=fileWithInsert, window=inserterWindow: filePicker(entry, window)).grid(row=3, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(inserterWindow, text="Page number of page to be inserted:").grid(row=4,column=0, padx=10, pady=3)
	pageWithInsert = tk.Entry(inserterWindow)
	pageWithInsert.grid(row=4, column=1, sticky=stickyFill, pady=5, padx=5)
	pageWithInsert.insert(0, "1")

	tk.Button(inserterWindow, text="Update!", command=lambda: inserterWindow.quit()).grid(row=5,column=0, columnspan=3, padx=5, pady=10, sticky=stickyFill)

	inserterWindow.mainloop()

	filename = updateFile.get()
	filename = filename[:-4] + '-updated.pdf'
	updateFile = checkExist(updateFile.get())
	pageToInsert = int(pageToInsert.get())
	fileWithInsert = checkExist(fileWithInsert.get())
	pageWithInsert = int(pageWithInsert.get())


	if pageToInsert == 0 or pageWithInsert == 0:
		popup("invalid page number, must be greater than 0")

	originalPDF = PyPDF2.PdfFileReader(updateFile)
	PDFwithInsert = PyPDF2.PdfFileReader(fileWithInsert)

	updatedPDF = PyPDF2.PdfFileWriter()
	updatedPDF.cloneDocumentFromReader(originalPDF)
	try:
		updatedPDF.insertPage(PDFwithInsert.getPage(pageWithInsert-1),pageToInsert-1)
	except IndexError:
		popup("Please check if page number is within range")
	
	outputFile = open(filename,'wb')

	pdfOut = PyPDF2.PdfFileWriter()

	for i in range(updatedPDF.getNumPages()):
		pdfOut.addPage(updatedPDF.getPage(i))

	pdfOut.write(outputFile)
	outputFile.close()

	inserterWindow.destroy()
	finished(filename, "Page insert", inserterWindow)


	# print()
	# print("Inserts the first page of a PDF into another PDF at a desired page")
	# oName = input("Enter name of file to be updated: ")
	# original = checkExist(oName)
	# insert = checkExist(input("Enter name of the PDF to be inserted: "))
	# pageNum = int(input("Enter the page number to insert into: "))
	# if pageNum == 0:
	# 	print("invalid page number, must be greater than 0")
	# 	return
	# pageIndex = pageNum - 1

	# oPDF = PyPDF2.PdfFileReader(original)
	# iPDF = PyPDF2.PdfFileReader(insert)

	# oBook = PyPDF2.PdfFileWriter()
	# oBook.cloneDocumentFromReader(oPDF)
	# oBook.insertPage(iPDF.getPage(0),pageIndex)

	# outputFile = open(oName+'-inserted.pdf','wb')
	# pdfOut = PyPDF2.PdfFileWriter()

	# for i in range(oBook.getNumPages()):
	# 		pdfOut.addPage(oBook.getPage(i))

	# pdfOut.write(outputFile)

	# print()
	# print("Updated file with insertion saved as", oName+"-inserted.pdf")
	# print()

def deletePage():
	deleterWindow = tk.Tk()
	deleterWindow.title("PDF page deleter")

	tk.Label(deleterWindow, text="Deletes a single page inside an existing PDF").grid(row=0,column=0, columnspan=3, padx=10, pady=3, sticky=stickyFill)

	tk.Label(deleterWindow, text="Select PDF file to edit:").grid(row=1, column=0, padx=10, pady=3)
	updateFile = tk.Entry(deleterWindow)
	updateFile.grid(row=1, column=1, sticky=stickyFill, pady=5, padx=5)
	tk.Button(deleterWindow, text="Browse...", command=lambda entry=updateFile, window=deleterWindow: filePicker(entry, window)).grid(row=1, column=2, pady=5, padx=5, sticky=stickyFill)

	tk.Label(deleterWindow, text="Page to delete:").grid(row=2,column=0, padx=10, pady=3)
	pageToDelete = tk.Entry(deleterWindow)
	pageToDelete.grid(row=2, column=1, sticky=stickyFill, pady=5, padx=5)

	tk.Button(deleterWindow, text="Delete!", command=lambda: deleterWindow.quit()).grid(row=3,column=0, columnspan=3, padx=5, pady=10, sticky=stickyFill)

	deleterWindow.mainloop()

	filename = updateFile.get()
	filename = filename[:-4] + '-updated.pdf'
	updateFile = checkExist(updateFile.get())
	pageToDelete = int(pageToDelete.get())


	if pageToDelete == 0:
		popup("invalid page number, must be greater than 0")

	originalPDF = PyPDF2.PdfFileReader(updateFile)
	
	updatedPDF = PyPDF2.PdfFileWriter()
	updatedPDF.cloneDocumentFromReader(originalPDF)
	try:
		updatedPDF.getPage(pageToDelete-1)
	except IndexError:
		popup("Please check if page number is within range")
	
	outputFile = open(filename,'wb')

	pdfOut = PyPDF2.PdfFileWriter()

	for i in range(updatedPDF.getNumPages()):
		if i != pageToDelete-1:
			pdfOut.addPage(updatedPDF.getPage(i))

	pdfOut.write(outputFile)
	outputFile.close()

	deleterWindow.destroy()
	finished(filename, "Page delete", deleterWindow)

def checkExist(fileName):
	try:
		openedFile = open(fileName,'rb')
		return openedFile
	except FileNotFoundError:
		popup('"'+fileName+'"'+" not found. Please restart program")


selector = tk.Tk()
selector.configure(padx=10, pady=10)
selector.title("PDF Editor")
# selector.minsize(width=400, height=50)

stickyFill = tk.N+tk.E+tk.W+tk.S

#body of GUI
tk.Label(selector, text="Functions:").grid(row=0,column=1, pady=5, padx=5)
tk.Button(selector, text="Merge PDFs", command=merge).grid(row=1,column=1, sticky=stickyFill, pady=3, padx=5)
tk.Button(selector, text="Update a single page", command=pageUpdate).grid(row=2,column=1, sticky=stickyFill, pady=3, padx=5)
tk.Button(selector, text="Insert a page into an existing PDF", command=insertPage, padx=20).grid(row=3,column=1, sticky=stickyFill, pady=3, padx=5)
tk.Button(selector, text="Delete a single page", command=deletePage).grid(row=4,column=1, sticky=stickyFill, pady=3, padx=5)


selector.protocol("WM_DELETE_WINDOW", quit)
selector.mainloop()


# print()
# print("Enter 1 to merge two PDF files")
# print("Enter 2 to update a single page of a PDF")
# print("Enter 3 to insert a single page into an existing PDF")
# choice = input("Enter function number: ")
# if choice == '1':
# 	merge()
# if choice == '2':
# 	pageUpdate()
# if choice == '3':
# 	insertPage()

# time.sleep(300)