import win32com
from win32com.client import Dispatch

if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')

    objPres = ppt.Presentations.Open("F:\ppt\\5-15-讲义-UI动效设计.pptx", WithWindow=False)

    listSilde = objPres.Slides

    for slide in listSilde:
        slideRange = slide.NotesPage
        text = slideRange.Shapes.Placeholders(2).TextFrame.TextRange.Text
        print(str(text).replace("\r", " "))
