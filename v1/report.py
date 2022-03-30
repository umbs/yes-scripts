import win32com.client
import PIL
from PIL import ImageGrab, Image
import os
import sys
import docx
from docx.shared import Inches

name_of_sheet = ""
inputFolderPath = os.getcwd()
inputExcelName = "summary.xlsx"
inputExcelFilePath = os.path.join(inputFolderPath, inputExcelName)
reportList = ["Hubbard", "Dorsa", "Lyndale", "Ben Painter", "Horace Cureton", "Linda Vista", "Renaissance"]
    
def get_sheet2():
    global name_of_sheet
    global inputFolderPath
    global inputExcelName
    global inputExcelFilePath
    global reportList
        
    # Open the excel application using win32com
    o = win32com.client.Dispatch("Excel.Application")
    # Disable alerts and visibility to the user
    o.Visible = 0
    o.DisplayAlerts = False
    # Open workbook
    wx = o.Workbooks.Open(inputExcelFilePath)
    
    for i in range(len(reportList)):
        rname = reportList[i]
        create_img(o, rname)
        create_report()
    wx.Close()
    o.Quit()


def create_img(o,sname):
    global name_of_sheet
    
    # Extract first sheet
    sheet = o.Sheets(sname)
    name_of_sheet = str(sheet.Name)
    for n, shape in enumerate(sheet.Shapes):
        # Save shape to clipboard, then save what is in the clipboard to the file
        shape.Copy()
        image = ImageGrab.grabclipboard()
        length_x, width_y = image.size
        size = int(length_x), int(width_y)
        image_resize = image.resize(size, Image.ANTIALIAS)
        # Saves the image into the existing png file (overwriting) TODO ***** Have try except?
        outputPNGImage = str(sheet.Name) + str(n) + '.png'
        image_resize.save(outputPNGImage, 'PNG', quality=100, dpi=(300, 300))
        pass
    pass


def create_report():
    global name_of_sheet
    
    doc = docx.Document()
    i = 0
    while i < 10:
        doc.add_picture(name_of_sheet + str(i) + '.png', width=Inches(4), height=Inches(4))
        if i < 9:
            doc.add_page_break()
        i += 1
    doc.save(name_of_sheet + ' Report.docx')


def delete_img():
    global inputFolderPath
    
    files_in_directory = os.listdir(inputFolderPath)
    filtered_files = [file for file in files_in_directory if file.endswith(".png")]
    for file in filtered_files:
        path_to_file = os.path.join(inputFolderPath, file)
        #print(str(path_to_file))
        os.remove(path_to_file)


def main():
    get_sheet2()
    delete_img()
    


if __name__ == "__main__":
    main()
