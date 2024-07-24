from lowModel.utils import *
import openpyxl as xl
import win32com.client
from math import ceil
from shutil import copyfile

from openpyxl.styles import Font
from openpyxl.styles.colors import Color
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

from pypdf import PdfMerger
from pypdf import PdfReader
from pdf2image import convert_from_path
import pytesseract
import glob

class Excel:
    def __init__(self,pathToXlsx:str):
        self.wb = xl.load_workbook(pathToXlsx)#OpenPyXl WorkBook
        self.path = pathToXlsx
        self.fileName = self.getFileNameByFilePath(pathToXlsx)
        self.folderPath = self.getFolderPathByFilePath(pathToXlsx)
        self.savesBackupFolder = self.getBackupFolderByPath(pathToXlsx)

    def getFileNameByFilePath(self,path:str):#Returns the text between the last slash and the .xlsx (FileName)
        return path[path.rfind('\\')+1:path.rfind('.')]
        
    def getFolderPathByFilePath(self,path:str):#Returns all text before the last slash (FolderPath)
        return path[:path.rfind('\\')+1]
    
    def getBackupFolderByPath(self,path:str):#Returns the path to a new folder where backups of the modified Excel will be stored
        return self.getFolderPathByFilePath(path) + '0oldVersions\\'

    def getSheets(self) -> list:#Returns the name of WorkSheets
        return self.wb.sheetnames
    
    def getRow(self,sheet:str,row:int):#Returns a list of Cells in Row (Begins with 1)
        return list(self.wb[sheet].rows)[row-1]
    
    def getColumn(self,sheet:str,column:int):#Returns a list of Cells in Column (Begins with 1)
        return list(self.wb[sheet].columns)[column-1]
    
    def getRowOfValue(self,sheet:str,column,value,occurrenceIndex=1):#Return the Row of Value in Column (Index -1 returns all occurrences)
        column = self.convertColumn(column)
        occurrences = []
        for row,cell in enumerate(self.getColumn(sheet,column),1):
            if cell.value == value:
                occurrences.append(row)
                if len(occurrences) == occurrenceIndex:
                    return row
        if occurrenceIndex == -1:
            return occurrences
        return None
    
    def getColumnOfValue(self,sheet:str,row:int,value,occurrenceIndex=1):#Return the Column of Value in Row (Index -1 returns all occurrences)
        occurrences = []
        for column,cell in enumerate(self.getRow(sheet,row),1):
            if cell.value == value:
                occurrences.append(column)
                if len(occurrences) == occurrenceIndex:
                    return column
        if occurrenceIndex == -1:
            return occurrences
        return None
    
    def getCellValue(self,sheet:str,column,row:int,allowFormula=True):#Returns Cell Value (Row begins with 1 and Column can be int or str)
        column = self.convertColumn(column)
        if allowFormula:
            return self.wb[sheet].cell(row,column).value
        else:
            readWb = xl.load_workbook(self.path,read_only=True,data_only=True)
            result = readWb[sheet].cell(row,column).value
            readWb.close()
            return result
        
    def setCellValue(self,sheet:str,column,row:int,value):#Change Cell Value (needs "save()")
        column = self.convertColumn(column)
        self.wb[sheet].cell(row,column).value = value

    def getCellFontStyle(self,sheet:str,column,row:int) -> Font:#Get Font Style from Cell
        column = self.convertColumn(column)
        return self.wb[sheet].cell(row,column).font

    def setCellFontStyle(self,sheet:str,column,row:int,fontStyle:Font):#Set a cell Font Style
        column = self.convertColumn(column)#Cell.Font doesn't allow fontStyle but allow a new Font equivalent
        cell = self.wb[sheet].cell(row,column).font = Font(fontStyle.name,fontStyle.sz,fontStyle.b,fontStyle.i,fontStyle.charset,fontStyle.u,fontStyle.strike,fontStyle.color,fontStyle.scheme,fontStyle.family,fontStyle.size)

    def setSize(self,sheet:str,column=None,row=None,size=-1):#Set a row/column height/width
        if row:
            self.wb[sheet].row_dimensions[row].height = size
        if column:
            column = self.convertColumn(column,toStr=True)
            self.wb[sheet].column_dimensions[column].width = size

    def setHide(self,sheet:str,column=None,row=None,hide=True):#Hide or Unhide a row/column
        if row:
            self.wb[sheet].row_dimensions[row].hidden = hide
        if column:
            column = self.convertColumn(column,toStr=True)
            self.wb[sheet].column_dimensions[column].hidden = hide

    def save(self,path=None,backup=True,pagesPdf=[-1]):#Save file changes in format of path (if path==None replace original file)
        if path == None:
            path = self.path
        if backup:#If backup==True create a copy of file in 'filePath/0oldVersions/'
            self.backup()

        if path[-4:] == '.pdf':#If path ends with .pdf save file as PDF
            self.savePdf(path,pagesPdf)
        elif path[-5:] == '.xlsx':#If path ends with .xlsx save file as Excel
            self.wb.save(path)#OpenPyXl function to save Excel

    def backup(self):#Create a copy of Excel in Backup folder
        if not os.path.exists(self.savesBackupFolder):#If backup folder doesn't exist, create it
            os.mkdir(self.savesBackupFolder)

        backupFilePath = self.savesBackupFolder + self.fileName + '.xlsx'

        version = 0
        while os.path.exists(backupFilePath):#If always exists a file with this name in this folder change the file name
            version += 1
            backupFilePath = self.savesBackupFolder + self.fileName + f' ({version})' + '.xlsx'

        copyfile(self.path,backupFilePath)#Copy file

    def savePdf(self,path:str,pagesPdf:list):#Save file as .pdf
        if path == None:
            path = self.path.replace('.xlsx','.pdf')
        
        #The code below was copied from StackOverflow so i don't know how it works
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(self.path)
        wb.WorkSheets(self.convertPages(pagesPdf)).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0,path)
        os.system("taskkill /f /im excel.exe")#Kill Excel after save

    def convertPages(self,pages:list):#Replace page name by page index
        if pages == [-1]:#If pages==[-1] (default value), returns a list [1...n] where n is the index of last page in file
            return range(1,len(self.getSheets())+1)
        
        for n,page in enumerate(pages):
            if type(page) == str and page in self.getSheets():#If page is the name of some Sheet, change value to sheet index
                pages[n] = self.getSheets().index(page)
        return pages

    def convertColumn(self,column,toStr=False):#Converts column to Str or Int
        if type(column) == str and toStr == False:
            return ord(column.upper())+1 - ord('A')
        if type(column) == int and toStr == True:
            return str(column-1+ord('A'))
        return column
    
    def close(self):#Close Excel file REMEMBER IT!!!!
        self.wb.close()
    
class Pdf:
    current_path = os.getcwd()
    pytesseract.pytesseract.tesseract_cmd = rf'{current_path}\lowModel\exterior\Tesseract-OCR\tesseract.exe'
    popplerBinPath = rf'{current_path}\lowModel\exterior\poppler-24.02.0\Library\bin'

    def merge(pdfs:list,outputPath):
        merger = PdfMerger()
        for pdf in pdfs:
            merger.append(pdf)
        merger.write(outputPath)
        merger.close()

    def readPdf(path) -> str:
        reader = PdfReader(path)
        text = ''
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text

    def readScannedPdf(path) -> str:
        pdfs = glob.glob(path)
        for pdf_path in pdfs:
            pages = convert_from_path(pdf_path, 500, poppler_path=Pdf.popplerBinPath)
            text = ''
            for imgBlob in pages:
                text += pytesseract.image_to_string(imgBlob,lang='por')
            return text
    
    def saveXML2003(pdfPath,xmlPath):
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.Documents.Open(pdfPath)
        word.ActiveDocument.SaveAs(xmlPath,11)
        word.Quit()

