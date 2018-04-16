import PyPDF2
import os, time

print("PDF Merger and Editor by simon wong")
print()

def merge():
	print()
	print("Creates a new PDF file with the second file added to the end of the first file")
	firstFile = checkExist(input("Enter the first file name (without the .pdf): "))
	secondFile = checkExist(input("Enter the second file name (without the .pdf): "))
	fname = input("Enter the desired name of the merged file (must be different from other files): ")

	mergedBook = PyPDF2.PdfFileMerger()
	mergedBook.append(firstFile)
	mergedBook.append(secondFile)

	fullbook = open(fname+'.pdf','wb')
	print()
	print("Merged PDF saved as", fname+'.pdf')
	mergedBook.write(fullbook)
	mergedBook.close()

def pageUpdate():
	print()
	print("Replaces a page from one PDF with an updated page")
	oName = input("Enter name of file to be updated: ")
	original = checkExist(oName)
	newPage = checkExist(input("Enter name of the PDF of the updated page: "))
	pageNum = int(input("Page number to be updated: "))
	if pageNum == 0:
		print("invalid page number, must be greater than 0")
		return
	pageIndex = pageNum - 1

	oPDF = PyPDF2.PdfFileReader(original)
	uPDF = PyPDF2.PdfFileReader(newPage)

	oBook = PyPDF2.PdfFileWriter()
	oBook.cloneDocumentFromReader(oPDF)
	oBook.insertPage(uPDF.getPage(0),pageIndex)

	outputFile = open(oName+'-updated.pdf','wb')
	pdfOut = PyPDF2.PdfFileWriter()

	for i in range(oBook.getNumPages()):
		if i != pageNum:
			pdfOut.addPage(oBook.getPage(i))

	pdfOut.write(outputFile)

	print()
	print("Updated file saved as", oName+"-updated.pdf")
	print()

def insertPage():
	print()
	print("Inserts the first page of a PDF into another PDF at a desired page")
	oName = input("Enter name of file to be updated: ")
	original = checkExist(oName)
	insert = checkExist(input("Enter name of the PDF to be inserted: "))
	pageNum = int(input("Enter the page number to insert into: "))
	if pageNum == 0:
		print("invalid page number, must be greater than 0")
		return
	pageIndex = pageNum - 1

	oPDF = PyPDF2.PdfFileReader(original)
	iPDF = PyPDF2.PdfFileReader(insert)

	oBook = PyPDF2.PdfFileWriter()
	oBook.cloneDocumentFromReader(oPDF)
	oBook.insertPage(iPDF.getPage(0),pageIndex)

	outputFile = open(oName+'-inserted.pdf','wb')
	pdfOut = PyPDF2.PdfFileWriter()

	for i in range(oBook.getNumPages()):
			pdfOut.addPage(oBook.getPage(i))

	pdfOut.write(outputFile)

	print()
	print("Updated file with insertion saved as", oName+"-inserted.pdf")
	print()


def checkExist(fileName):
	try:
		openedFile = open(fileName+'.pdf','rb')
		return openedFile
	except FileNotFoundError:
		print(fileName+".pdf not found in the folder. Please restart program")
		print()
		time.sleep(60)
		quit()
		

print()
print("Enter 1 to merge two PDF files")
print("Enter 2 to update a single page of a PDF")
print("Enter 3 to insert a single page into an existing PDF")
choice = input("Enter function number: ")
if choice == '1':
	merge()
if choice == '2':
	pageUpdate()
if choice == '3':
	insertPage()

time.sleep(300)