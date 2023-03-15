import openpyxl
from openpyxl.drawing.image import Image
import PySimpleGUI as sg

def GifToExcel(gif_path):
    #Define the Excel workbook and sheet.
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    #Locate the GIF.
    img = Image(gif_path)

    #Add the GIF to the worksheet.
    ws.add_image(img)

    #Save the workbook.
    wb.save('gif_info.xlsx')

    print("Done")

  
# Add some color
# to the window
sg.theme('SandyBeach')     
  
# Very basic window.
# Return values using
# automatic-numbered keys
layout = [
    [sg.Text('Please select the gif to convert')],
    [sg.In(key="Gif"), sg.FileBrowse(file_types=(("Text Files", "*.gif"),))],
    # [sg.Text('Name', size =(15, 1)), sg.InputText()],
    [sg.Submit(), sg.Cancel()]
]
  
window = sg.Window('Gif data to Excel', layout)
event, values = window.read()
window.close()
  
# The input data looks like a simple list 
# when automatic numbered
print(event, values["Gif"])   
if event=="Submit":
    GifToExcel(values["Gif"])
