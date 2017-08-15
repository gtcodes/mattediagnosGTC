import csv
from os import listdir, makedirs, path
from os.path import isfile, join
from docx import Document

START_DATA = 2 #What column the data starts at, 0 indexed so START_DATA = 2 would mean that the first data column is column C
RESULTS_FOLDER = 'resultat'
RECIPE_FOLDER = 'recept'
req = [1, 4, 6, 6, 6, 5, 5]

def readCSVFile(file_name):
    with open(file_name, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=';', quotechar='|')
        reader = list(reader)
        className = reader[0][0]

        print("class name", reader[0][0])
        if not path.exists(join(RECIPE_FOLDER, className)):
            makedirs(join(RECIPE_FOLDER, className))
        
        for row in reader[2:]:
            if(row[0] !=''):
                print("creating file for", row[0])
                createSingleRecipe(className, row)
            else:
                break

def createSingleRecipe(className, row):
    name = row[0]                                       #First cell is the name of the test taker
    isHelpNeeded = [0]*7    
    for i in range(0,7):
        if(int(row[START_DATA + i]) < req[i]):          #If the score on this part is less than required 
            isHelpNeeded[i] = 1                         #Help needed is set to one for this part
    files = selectFiles(isHelpNeeded)
    mergeWordFiles(className, files[0],files[1], name)

#Depending on what parts are needed for the test taker, different files are selected
def selectFiles(booleanVector):
    recipeFiles = []
    headerFiles = []
    for i in range(0,7):
        if(booleanVector[i] == 1):
            headerFiles.append('Header' + str(i+1) + '.docx')
            recipeFiles.append('Recipe' + str(i+1) + '.docx')
        else:
            headerFiles.append('HeaderDone' + str(i+1) + '.docx')
    return (recipeFiles, headerFiles)

def mergeWordFiles(className, recipeFiles, header, name):
    #can create the header by merging all the header files
    #or one can do it by manually selecting the paragraphs
    merged_document = createStandardHeader(name)
    del(merged_document.element.body[-1])
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
    
    merged_document.save(join(RECIPE_FOLDER, className.strip(), name.strip() + '.docx'))

#A standard header is created that is needed for each document, the rest of the headers are later appended to this.
def createStandardHeader(name):
    document = Document()
    document.add_heading('Ditt matterecept' + name + '!', 0)
    document.add_paragraph('Speciellt framtaget för att du ska kunna arbeta med rätt saker som ger'
            ' just dig så bra förutsättningar som möjligt att lyckas med matematik!\n')
    return(document)

for f in listdir(RESULTS_FOLDER):
    readCSVFile(join(RESULTS_FOLDER, f))