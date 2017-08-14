import csv
import numpy as np
import pandas as pd
from docx import Document

df = pd.read_csv('test.csv', sep = ';', header = 1)
req = [1, 4, 6, 6, 6, 5, 5]

def readCSVFile(file_name):
    with open(file_name, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=';', quotechar='|')
        reader.__next__()
        for row in reader:
            createSingleRecipe(row)

def createSingleRecipe(row):
    name = row[0]
    isHelpNeeded = [0]*7
    for i in range(1,8):
        if(int(row[i]) < req[i - 1]):
            isHelpNeeded[i - 1] = 1
    files = selectFiles(isHelpNeeded)
    return (mergeFiles(files[0],files[1], name))

def selectFiles(booleanVector):
    fileNames = []
    headerNames = []
    #might need to create some header specific to the student
    for i in range(0,7):
        if(booleanVector[i] == 1):
            headerNames.append('header' + str(i) + '.docx')
            fileNames.append('recipe' + str(i) + 'docx')
        else:
            headerNames.append('headerDone' + str(i) + 'docx')
    return (fileNames, headerNames)

def mergeWordFiles(recipeFiles, header, name):
    #can create the header by merging all the header files
    #or one can do it by manually selecting the paragraphs
    merged_document = createStandardHeader(name)
    for file in enumerate(header):
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    for i, file in enumerate(recipeFiles):
        sub_doc = Document(file)
        if(i < len(files) - 1):
            sub_doc.add_page_break()
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save('merged.docx')

def createStandardHeader(name):
    document.add_heading('Ditt matterecept ' + name + '!', 0)
    document.add_paragraph('följande document innehåller individuellt framtagna'
            ' övningsuppgifter som skall hjälpa dig lyckas med matematiken.')


#pasteRecipe(1,1)
#mergeWordFiles(['test1.docx','test2.docx'],'test1.docx')
#os.listdir(<path>)
readCSVFile('test.csv')
