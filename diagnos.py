import csv
import numpy as np
import pandas as pd
from docx import Document

START_DATA = 1
df = pd.read_csv('test.csv', sep = ';', header = 1)
req = [1, 4, 6, 6, 6, 5, 5]

def readCSVFile(file_name):
    with open(file_name, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=';', quotechar='|')
        reader.__next__()
        for row in reader:
            if(row[0] !=''):
                createSingleRecipe(row)
            else:
                break

def createSingleRecipe(row):
    name = row[0]
    isHelpNeeded = [0]*7
    for i in range(0,7):
        if(int(row[START_DATA + i]) < req[i]):
            isHelpNeeded[i - 1] = 1
    files = selectFiles(isHelpNeeded)
    return (mergeWordFiles(files[0],files[1], name))

def selectFiles(booleanVector):
    recipeFiles = []
    headerFiles = []
    #might need to create some header specific to the student
    for i in range(0,7):
        if(booleanVector[i] == 1):
            headerFiles.append('header' + str(i+1) + '.docx')
            recipeFiles.append('recipe' + str(i+1) + '.docx')
        else:
            headerFiles.append('headerDone' + str(i+1) + '.docx')
    return (recipeFiles, headerFiles)

def mergeWordFiles(recipeFiles, header, name):
    #can create the header by merging all the header files
    #or one can do it by manually selecting the paragraphs
    merged_document = createStandardHeader(name)
    del(merged_document.element.body[-1])
    for element in merged_document.element.body:
        print(element)
    print(recipeFiles)
    print(header)
    for fileName in header:
        sub_doc = Document(fileName)
        for i, element in enumerate(sub_doc.element.body):
            if(i < len(sub_doc.element.body)):
                merged_document.element.body.append(element)
    merged_document.add_page_break()
    for i, fileName in enumerate(recipeFiles):
        sub_doc = Document(fileName)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    
    for element in merged_document.sections:
        print(element)
    merged_document.save(name + 'merged.docx')

def createStandardHeader(name):
    document = Document()
    document.add_heading('Ditt matterecept ' + name + '!', 0)
    document.add_paragraph('följande document innehåller individuellt framtagna'
            ' övningsuppgifter som skall hjälpa dig lyckas med matematiken.\n')
    return(document)


#pasteRecipe(1,1)
#mergeWordFiles(['test1.docx','test2.docx'],'test1.docx')
#os.listdir(<path>)
readCSVFile('test.csv')
