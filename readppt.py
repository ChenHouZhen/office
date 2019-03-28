import win32com
from win32com.client import Dispatch
import math


if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # MsoTriState.msoTrue        1
    objPres = ppt.Presentations.Open("F:\ppt\\123456789.pptx", WithWindow=True)

    # Presentations.Slides 返回PPT的所有幻灯片集合
    listSlides = objPres.Slides

    for slide in listSlides:

        sequence = slide.TimeLine.MainSequence

        len_effect = sequence.Count

        for j in range(1, len_effect + 1):
            effect = sequence.Item(j)
            effect_shape = effect.Shape
            print("=================== effect_shape.AnimationSettings.AnimationOrder :{}  ======================".format(effect_shape.AnimationSettings.AnimationOrder))
            print("=================== effect_shape.Name:{} ======================".format(effect_shape.Name))
            print("=================== effect_shape.Id:{} ======================".format(effect_shape.Id))
            print("=================== effect.Timing.Duration:{} ======================".format(effect.Timing.Duration))
            print("=================== effect.Shape.AnimationSettings.AdvanceTime:{} ======================".format(effect.Shape.AnimationSettings.AdvanceTime))
            print("=================== effect.Shape.AnimationSettings.AdvanceMode:{} ======================".format(effect.Shape.AnimationSettings.AdvanceMode))
            print("=================== effect.Timing.TriggerType:{}======================".format(effect.Timing.TriggerType))
            print("=================== effect.Shape.AnimationSettings.PauseAnimation:{} ======================".format(effect_shape.AnimationSettings.PlaySettings.PauseAnimation))
            print("=================== effect.Shape.AnimationSettings.StopAfterSlides:{} ======================".format(effect_shape.AnimationSettings.PlaySettings.StopAfterSlides))
            print("===================  ======================")
            print()
            print()

ppt.Quit()


