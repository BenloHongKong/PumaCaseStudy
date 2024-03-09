import os,shutil
import win32com.client as client
from pathlib import Path
import os,PyPDF2,openpyxl


class DownloadEmail:
    def __init__(self, outlookFolderName, parentFolder, email, moveEmailToFolder, writeToExcel, moveFolderTo):
        self.outlookFolderName = outlookFolderName
        self.parentFolder = parentFolder
        self.email = email
        self.moveEmailToFolder = moveEmailToFolder
        self.writeToExcel = writeToExcel
        self.moveFolderTo = moveFolderTo
        self.originalItem=self.checkOriginalFolderItem()

    def replaceEscapeCharacter(self, name):
        '''Replace all special characters including space'''
        escapeList = ["|", ":", "<", ">", "\"", "\\", "/", "*", "?", " "]
        for escape in escapeList:
            name = name.replace(escape, "_")
        return name

    def saveMessage(self, message, outputFolder):
        try:
            subject = message.Subject
            if not subject:
                subject = 'NoSubject'
            subject = self.replaceEscapeCharacter(subject)
            outputFile = os.path.join(outputFolder, f"{str(subject)}.msg")
            message.SaveAs(outputFile)
        except Exception as e:
            print(f"Error saving email: {e}")

    def downloadEmails(self, messages):
        '''Download email itself and download all attachments'''

        checkFolderInTargetFolder=os.listdir(self.moveFolderTo)
        for message in messages:
            while True:
                originalFolder=self.replaceEscapeCharacter(str(message.Subject))
                if checkFolderInTargetFolder.count(originalFolder) or os.listdir(self.parentFolder).count(originalFolder):
                    counter=1
                    eachEmailFolder=originalFolder
                    while checkFolderInTargetFolder.count(eachEmailFolder) or os.listdir(self.parentFolder).count(eachEmailFolder):
                        eachEmailFolder=f'{originalFolder}_{counter}'
                        counter+=1
                    break
                else:
                    eachEmailFolder = os.path.join(self.parentFolder, self.replaceEscapeCharacter(str(message.Subject)))
                    break
            attachments = message.Attachments
            finalPath=os.path.join(parentFolder,eachEmailFolder)
            os.makedirs(finalPath)

            if message.Class == 43:  # 43 = olMail class
                self.saveMessage(message, finalPath)
            for attachment in attachments:
                attachment.SaveAsFile(os.path.join(finalPath, self.replaceEscapeCharacter(str(attachment))))

    def moveEmails(self, messages):
        '''Move emails to a specified folder'''
        messageList=list(messages)
        outlook = client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
        resultFolder = outlook.Folders.Item(self.email).Folders.Item(self.moveEmailToFolder)
        for message in messageList:
            message.Move(resultFolder)
            print(f"Email moved: {message.Subject}")
        print("All emails moved successfully.")

    def moveFolders(self):
        '''Move folders to a specified destination'''
        for folder in os.listdir(self.parentFolder):
            if str(folder).startswith("__") or self.originalItem.count(folder):
                continue
            else:
                counter=1
                originalFolder=os.path.join(self.parentFolder,folder)
                targetFolder=os.path.join(self.moveFolderTo,folder)
                while os.path.exists(targetFolder):
                    targetFolder=os.path.join(self.moveFolderTo,folder)
                    targetFolder = f"{targetFolder}_{counter}"
                    counter += 1
                shutil.move(originalFolder,targetFolder)
    
    def checkOriginalFolderItem(self):
        originalItem=[]
        for item in os.listdir(self.parentFolder):
            originalItem.append(item)
        return originalItem


    def main(self):
        outlook = client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
        requiredMessage = outlook.Folders(self.email).Folders(self.outlookFolderName).Items

        # Download email and attachment
        self.downloadEmails(requiredMessage)
        # Move emails to specified folder
        self.moveEmails(requiredMessage)
        print("move all Email")
        # Move folders to specified destination
        #self.moveFolders()



class RecordInExcel:
    def __init__(self, parentFolder, excelPath):
        self.parentFolder = parentFolder
        self.excelPath = excelPath
        self.alreadyInsideFolder=[]
        self.checkAlreadyInsideFolder()

    def checkAlreadyInsideFolder(self):
        for item in os.listdir(self.parentFolder):
            self.alreadyInsideFolder.append(item)

    def getSubfolders(self):
        subfolderList = []
        for item in os.listdir(self.parentFolder):
            checkFolder = os.path.join(self.parentFolder, item)
            if os.path.isdir(checkFolder):
                subfolderList.append(item)
        return subfolderList

    def pdfToTxt(self, pdfPath, txtPath):
        try:
            with open(pdfPath, 'rb') as pdfFile:
                pdfReader = PyPDF2.PdfReader(pdfFile)
                with open(txtPath, 'w', encoding='utf-8') as txtFile:
                    txtFile.write(pdfReader.pages[0].extract_text())
        except Exception as e:
            print(f"Error: {e}")

    def getInformationToEnd(self, startWith, txtPath):
        with open(txtPath, 'r', encoding='utf-8') as txtFile:
            for line in txtFile:
                if line.startswith(startWith):
                    information = line.split(startWith)[1].strip()
                    break
            else:
                information = ""
        return information

    def getNextLine(self, lastLine, txtPath):
        with open(txtPath, 'r', encoding='utf-8') as txtFile:
            lines = txtFile.readlines()
            for i, line in enumerate(lines):
                if str(line).startswith(lastLine):
                    information = lines[i + 1].strip()  # Drop \n
                    break
            else:
                information = ""
        return information

    def processGetData(self):
        subfolderList = self.getSubfolders()
        allStoreList = [["Supplier Name", "Type of Certificate", "Expiry Date", "Certificate File Name", "Folder Name"]]
        for subfolder in subfolderList:
            subfolderPath = os.path.join(self.parentFolder, subfolder)
            if self.alreadyInsideFolder.count(subfolder):
                continue
            for file in os.listdir(subfolderPath):
                if file.lower().endswith(".pdf"):
                    onePdfPath = os.path.join(self.parentFolder, subfolder, file)
                    oneTxtPath = os.path.join(self.parentFolder, subfolder, file[0:file.rfind(".")] + ".txt")
                    self.pdfToTxt(onePdfPath, oneTxtPath)
                    fileStoreList = [
                        self.getNextLine("Control Union Certifications declares that", oneTxtPath),
                        self.getNextLine("has been inspected and assessed in accordance with the", oneTxtPath),
                        self.getNextLine("This certificate is valid until:", oneTxtPath),
                        str(file),
                        str(subfolder)
                    ]
                    if fileStoreList.count("") == len(allStoreList[0])-2:
                        pass  # want to exclude those pdf that are not certificate
                    else:
                        allStoreList.append(fileStoreList)
        allStoreList.pop(0)
        return allStoreList

    def deleteEmptyRow(self, workbook, sheet):
        maxRow = sheet.max_row
        maxCol = sheet.max_column
        rowsDelete = []
        for i in range(maxRow + 1):
            if i == 0:
                continue
            rowElement = []
            for j in range(maxCol + 1):
                if j == 0:
                    continue
                rowElement.append(sheet.cell(row=i, column=j).value)
            if rowElement.count(None) == maxCol:
                rowsDelete.append(i)
        for i, row in enumerate(rowsDelete):
            sheet.delete_rows(row - i)  # trick when first row deleted, another row will -1
        workbook.save(self.excelPath)

    def appendToExcel(self, workbook, sheet, data):
        try:
            # Find the last row with data
            maxRow = sheet.max_row
            # Append new data below the last row
            for rowIdx, rowData in enumerate(data, start=1):
                for colIdx, cellData in enumerate(rowData, start=1):
                    sheet.cell(row=maxRow + rowIdx, column=colIdx, value=cellData)

            workbook.save(self.excelPath)
            print(f"Data appended to '{self.excelPath}' successfully.")
        except Exception as e:
            print(f"Error: {e}")

    def main(self):
        allStoreList = self.processGetData()
        workbook = openpyxl.load_workbook(self.excelPath)
        sheet = workbook["Sheet"]
        self.appendToExcel(workbook, sheet, allStoreList)
        self.deleteEmptyRow(workbook, sheet)





if __name__ == "__main__":
    #preparation: create testAuto and testAuto_Done Folder in outlook first
    outlookFolderName = 'testAuto'
    parentFolder = os.getcwd()
    email = "1155158175@link.cuhk.edu.hk"
    moveEmailToFolder = "testAuto_Done"
    writeToExcel = Path("~/OneDrive/Documents/autoUpdate.xlsx").expanduser()
    moveFolderTo = Path("~/OneDrive/Documents").expanduser()


    downloader = DownloadEmail(outlookFolderName, parentFolder, email, moveEmailToFolder, writeToExcel, moveFolderTo)
    updateExcel=RecordInExcel(parentFolder,writeToExcel)
    
    downloader.main()
    updateExcel.main()
    
    downloader.moveFolders()
