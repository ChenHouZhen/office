import win32com
from win32com.client import Dispatch

# 文本不完整
if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')

    objPres = ppt.Presentations.Open("F:\ppt\\pptword2.pptx", WithWindow=False)

    listSlides = objPres.Slides

    for slide in listSlides:
        listShape = slide.Shapes
        listShape = sorted(listShape, key=lambda x: (x.Top))
        for shape in listShape:
            # print(shape.Left)
            # print(shape.Top)
            if shape.HasTextFrame:
                s = shape.TextFrame2.TextRange.Text
                # 要去掉 \r
                print(s.replace("\r", " "))

    ppt.Quit()