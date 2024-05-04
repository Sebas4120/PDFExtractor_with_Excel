#Name: Sebastian Aguirre - 200558953 
#Date: 02/04/2024
#Description: Project 1

import pyinputplus as pyip
import PyPDF2,openpyxl,os,re,time

#To start the program the user have to type START
startProgram=pyip.inputStr(prompt="Please enter \"START\" to get your new PDF document:")
#Use of WHILE NOT waiting for the user answer
while not(startProgram.upper()=="START"):
    startProgram=pyip.inputStr(prompt="Please enter \"START\" to get your new PDF document:")
   
#Check that the excel data is correct.
def validationRange():
    #First let's extract the data from excel: document name, page range and individual pages to extract.
    excelGui=openpyxl.load_workbook("C:\\Users\\sebas\\OneDrive\\Documents\\GEORGIAN\\SEMESTER 2\\PYTHON\\PROJECT\\GUI_Project.xlsx")
    sheet=excelGui.active
    #pdfName=sheet['D3']
    sinceNumber=sheet['D5']
    untilNumber=sheet['D6']
    indiviPages=sheet['D8']
    #String to list conversion
    listIndiPages=(str(indiviPages.value).split(","))
    #Check if the user entered the page from where to start.
    if (sinceNumber.value!=None):
        #check if the user's entry is a number
        if not(str(sinceNumber.value).isdecimal()):
            print("Initial range is not a number")
            return False
    #In case the user does not enter the initial page
    else:
        print("Please enter the page from where you want to get the new pdf file.")
        return False
#Check if the user entered the final page of the range.
    if (untilNumber.value!=None):
        #check if the user's entry is a number
        if not(str(untilNumber.value).isdecimal()):
            print("The final range is not a number")
            return False
    #In case the user does not enter the final page
    else:
        print("Please enter the page to where you want to get the new pdf file.")
        return False
#Check if the individual pages you want are correct.
    if (listIndiPages!= ['None']):
        for i in range(len(listIndiPages)):
            if not(listIndiPages[i].isdecimal()):
                print("Not number one of the pages, please check")
                return False
    return True

#Once the validation has been carried out, the new pdf with the user's request is obtained.
def getNewPdf():
    #First let's extract the data from excel: document name, page range and individual pages to extract.
    excelGui=openpyxl.load_workbook("C:\\Users\\sebas\\OneDrive\\Documents\\GEORGIAN\\SEMESTER 2\\PYTHON\\PROJECT\\GUI_Project.xlsx")
    sheet=excelGui.active
    pdfName=sheet['D3']
    newPdfName=sheet['D12']
    sinceNumber=sheet['D5']
    untilNumber=sheet['D6']
    indiviPages=sheet['D8']
    listIndiPages=(str(indiviPages.value).split(","))
    #Open the source PDF file from which we want to obtain the pages and create a reader object to add the extracted pages. 
    pdfFile=open(pdfName.value,'rb')
    pdfReader=PyPDF2.PdfReader(pdfFile)
    pdfWriter=PyPDF2.PdfWriter()   
    for i in range(sinceNumber.value-1,untilNumber.value):
        pageObj=pdfReader.pages[i]
        pdfWriter.add_page(pageObj)
    #We check if the user wants to add individual pages to their new pdf document.
    if (listIndiPages!= ['None']):
        #Individual pages added to the new PDF
        for i in listIndiPages:
            pageObj=pdfReader.pages[int(i)-1]
            pdfWriter.add_page(pageObj)
    #The new pdf is created with the extracted pages.
    pdfFinalFile=open(newPdfName.value+'.pdf','wb')
    pdfWriter.write(pdfFinalFile)
    pdfFinalFile.close()
    pdfFile.close()
    print("new PDF created!")


#This function will search an specific word that the user wants to search within the PDF, in case the user want to.        
def searchWord():
    #Empty list for storing the pages where the word is found
    pagesWithWord=[]
    #Getting the values from the excel
    excelGui=openpyxl.load_workbook("C:\\Users\\sebas\\OneDrive\\Documents\\GEORGIAN\\SEMESTER 2\\PYTHON\\PROJECT\\GUI_Project.xlsx")
    sheet=excelGui.active
    pdfName=sheet['D3']
    wantedWord=sheet['D10']
    print("Seeking word:",wantedWord.value)
    #Check if the user wants to search for a word
    if (wantedWord.value!=None):
        #Using Regex
        regexObj=re.compile(wantedWord.value)
        #Starting the search in the pdf
        pdfFile=open(pdfName.value,'rb')
        pdfReader=PyPDF2.PdfReader(pdfFile)
        #Use of For-range loop to go through pages looking for the word and get the pages where it is found
        for i in range(len(pdfReader.pages)):
            page=pdfReader.pages[i]
            pageText=page.extract_text()
            if re.search(regexObj,pageText):
                pagesWithWord.append(i+1)
        pdfFile.close()
        return pagesWithWord
    else:
        #If the user does not put a word to search
        print("No word to search")
        
#This function will create a text file to create a history of user activity.
def historyFile():
    #Extract the data from excel: document name, page range and individual pages to extract,etc.
    excelGui=openpyxl.load_workbook("C:\\Users\\sebas\\OneDrive\\Documents\\GEORGIAN\\SEMESTER 2\\PYTHON\\PROJECT\\GUI_Project.xlsx")
    sheet=excelGui.active
    pdfName=sheet['D3']
    sinceNumber=sheet['D5']
    untilNumber=sheet['D6']
    indiviPages=sheet['D8']
    newPdfName=sheet['D12']
    listIndiPages=(str(indiviPages.value).split(","))
    historiNameFile=sheet['D13']
    wantedWord=sheet['D10']
    #Using TIME MODULE to set the time in the history file
    now=time.time()
    #In case the history document already exists, the new data is added.
    if (os.path.exists(historiNameFile.value + '.txt')==True):
        textFile=open(historiNameFile.value+'.txt','a')
    #If it does not exist, it is created 
    else:            
        textFile=open(historiNameFile.value+'.txt','w')
    #The data entered by the user in the excel file is added to the user's history.   
    textFile.write(time.ctime(now)+ "\n")
    textFile.write("PDF origin: " + pdfName.value[:(len(pdfName.value)-4)] + "\n")
    textFile.write("New PDF created: " + newPdfName.value + "\n")
    textFile.write("From the page: " + str(sinceNumber.value) + "\n")
    textFile.write("To the page: " + str(untilNumber.value) + "\n")
    textFile.write("List of individual pages: " + str(listIndiPages)+"\n")
    if (wantedWord.value!=None):
        textFile.write("Wanted word: " + str(wantedWord.value)+"\n")
    textFile.write("Page number where the word was found : " + str(searchWord())+"\n")
    textFile.write("------------------" + "\n")
    textFile.close()
    print("Text history file created")
        

try:
    #Validate if the data entered by the user is correct.
    if(validationRange()==True):
        #Execution of functions
        getNewPdf()
        print("Page number where the word was found:",searchWord())
        historyFile()

#Use of EXCEPTION HANDLING in case the the PDF document is not found  
except FileNotFoundError:
    print("File not found") 



