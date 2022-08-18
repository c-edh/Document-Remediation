#By Corey Edh

import re
import docx
import tkinter as tk
from tkinter.filedialog import askopenfilename
import helper

#Window
window = tk.Tk()
window.title("Document Assistant")


class userGUI():
    def __init__(self):
        self.path = None
        self.doc = None
    
    def open_file(self, outPutFileName):
        if outPutFileName == "":
            return False
        filepath = askopenfilename(
            filetypes=[("Documents", "*.docx")]
        )
        if(filepath):
            
            self.doc = helper.Word_Operations(filepath)
            print(filepath)
            lastSlash = 0
            for char in range(len(filepath)):
                if filepath[char] == "/":
                    lastSlash = char

            #Gets the directory of the file
            filedirect = filepath[0:lastSlash+1]
            print(filedirect)

            self.path = filedirect
            self.doc.path = self.path
            
            #Tells Helper to save the file with the following name and path
            self.doc.save = self.path + outPutFileName +".docx"
            return True
        
        else:
            return False

        
    
    #Adjust the page headers
    def adjustPageHeadings(self,start):
        if self.doc:
            if start:
                self.doc.pagesToHeading(start)
            else:
                print("Error, Start is missing")
                return False
        else:
            print("No doc was open")

    #Finds images
    def imageFinder(self):
        self.doc.findImagesThatNeedAltText()

    def getSection(self):
        self.doc.getSections()

    def open_SectionFile(self):

        filepath = askopenfilename(
            filetypes=[("Documents", "*.docx")]
        )
        
        if(filepath):

            self.Sections = docx.Document(filepath)
            sectionDoc = self.Sections
            sytle = ""
            styleOfSectionText = []

            for paragraph in sectionDoc.paragraphs:
                if paragraph.text.lower() == "h1":
                    style = "H1"
                elif paragraph.text.lower() == "h2":
                    style = "H2"
                elif paragraph.text.lower() == "h3":
                    style = "H3"
                elif paragraph.text.lower() == "h4":
                    style = "H4"
                elif paragraph.text.lower() == "h5":
                    style = "H5"
                elif paragraph.text.lower() == "h6":
                    style = "H6"
                else:
                    #Removes whitespaces, and 
                    paragraphText = re.sub(r"[\n\t\s]*", "", paragraph.text.lower())
                    styleOfSectionText.append([paragraphText,style])

            print(styleOfSectionText)

            self.doc.getSections(styleOfSectionText)
                
          
            return True
        else:
            return False
        


#Event Handler?
action = userGUI()

#GUI Layout----------------------------------------------------


window.columnconfigure([0,1,2], minsize = 75)
window.rowconfigure([0,
                     1,
                     2,
                     3], minsize = 20)


###Column 0--------------------------------------------

saveFrame = tk.Frame(window)
saveFrame.grid(column=0,
               row=1)



fileStatus = tk.StringVar()
fileStatus.set("No File has been selected")

fileIsOpenedLabel = tk.Label(textvariable=fileStatus)
fileIsOpenedLabel.grid(row=0,column=0)

saveToLabel = tk.Label(master = saveFrame, text="Save as:")
saveToLabel.pack()

fileName = tk.StringVar()
getFileName = tk.Entry(master=saveFrame,textvariable=fileName)
getFileName.pack()

#Opens file, and changes text to let user know the file has been selected
def guiFileOpen():
    outPutFile = fileName.get()
    print(outPutFile)
    if outPutFile:
        openFile = action.open_file(outPutFile)
        if openFile:
            fileStatus.set("File has been selected")
        else:
            fileStatus.set("No File has been selected")
    else:
        fileStatus.set("Name the Output File")

openFileButton = tk.Button(text="Open File", command=guiFileOpen)
openFileButton.grid(row=0,
                    column=2)

#Row 1------------------------------------------------------

pageFrame = tk.Frame(window)
pageFrame.grid(column=0,
               row=2)

#Vars to read userInput
startPage = tk.StringVar()

def getUserPageInput():
    start = startPage.get()
    fileStatus.set('This may take a minute \n if it freezes do not exit')
    pageInput = action.adjustPageHeadings(start)
    fileStatus.set('Pages have been \n converted to H6')
    if pageInput == False:
        fileStatus.set('Forgot to add pages')






#textfields to get User input
PageLabel = tk.Label(master=pageFrame, text="Start Page:", width=10)

startPageEntry = tk.Entry(master=pageFrame, textvariable=startPage)

#Adding buttons/labels to GUI
PageLabel.pack()
startPageEntry.pack()


#Column 1-----------------------------------------------------------

#Button to convert
pageHeaderButton = tk.Button(master=window, text = "Page Headers", 
command=getUserPageInput)
pageHeaderButton.grid(column=1, row=0)


imagesAltTextNeedButton = tk.Button(master= window, text = "Image Alt Text", command=action.imageFinder)
imagesAltTextNeedButton.grid(column=1, row=1)



def remediatePagesAndImages():
    start = startPage.get()
    action.adjustPageHeadings(start)
    action.imageFinder()

remediatePageAndImages = tk.Button(master = window, text = "Images and Pages", command = remediatePagesAndImages)
remediatePageAndImages.grid(column = 1, row=2)


openSection = tk.Button(master = window, text = "Section Document", command = action.open_SectionFile)
openSection.grid(column=1,row=3)

#Column2-----------------------------------------------------------------------------





window.mainloop()
