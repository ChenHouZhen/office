import win32com
from win32com.client import Dispatch

# 文本不完整
if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')

    ppt.Visible = 1
    pptSel = ppt.Presentations.Open("F:\ppt\\pptword2.pptx", WithWindow=False)

    # 顺序没有稳定
    slide_count = pptSel.Slides.Count
    for i in range(1, slide_count + 1):
        shape_count = pptSel.Slides(i).Shapes.Count
        # print("组件数量："+str(shape_count))
        for j in range(1, shape_count + 1):
            shape = pptSel.Slides(i).Shapes(j)
            if pptSel.Slides(i).Shapes(j).HasTextFrame:
                s = pptSel.Slides(i).Shapes(j).TextFrame2.TextRange.Text

                # 要取掉 \r
                print(s.replace("\r", " "))

    ppt.Quit()