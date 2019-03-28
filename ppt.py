import win32com
from win32com.client import Dispatch


def add_voice(sequence,voice_path, order):
    print("========================== 插入音频,音频文件名为：{}".format(voice_path))

    # 插入媒体文件
    # 第一个参数：要添加的文件的名称。
    # 第二个参数： 指示是否链接到文件。
    # 第三个参数：指示是否随文档一起保存媒体。
    # 第四个参数：从幻灯片左边缘到媒体对象左边缘的距离（以磅为单位）。
    # 第五个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
    # 第六个参数： 媒体对象的宽度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
    # 第七个参数： 媒体对象的高度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
    shape = slide.Shapes.AddMediaObject2(voice_path, 0, -1, 1120, 680, 0, 0)

    # PlaySettings对象 关于指定的媒体剪辑在幻灯片放映中的播放方式的信息。
    # 指定的影片或声音后是否在激活后自动播放。 -1 为 true

    shape.AnimationSettings.PlaySettings.PlayOnEntry = True
    # shape.AnimationSettings.PlaySettings.PauseAnimation = True
    # shape.AnimationSettings.PlaySettings.StopAfterSlides = 0

    # shape.AnimationSettings.PlaySettings.LoopUntilStopped = False

    # shape.AnimationSettings.AdvanceMode = 2
    # shape.AnimationSettings.AdvanceTime = 0.0

    # 幻灯片放映期间指定媒体剪辑在不播放时是否隐藏。幻灯片放映过程中指定媒体隐藏。
    # shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = False

    # ------------------设置音频图标----------------------
    shape.MediaFormat.SetDisplayPictureFromFile('F:/ppt/voice15.png')
    shape.Width = 30
    shape.Height = 30
    # --------------------------------------------------

    # effect = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel);

    # shape.AnimationSettings.AnimationOrder = order
    print("========================== 插入音频成功,音频总时长为：{}".format(str(shape.MediaFormat.Length)))

    # 返回一个**Effect** 对象, 该对象代表一个新的动画效果添加到动画效果序列中。
    # 第一個參數：向其添加动画效果的形状。
    # 第二個參數：要应用的动画效果。
    # 第三個參數：为图表、 图示或文本，将对其应用的动画效果级别。
    # 第四個參數：触发动画效果的动作
    # 第五個參數：效果在动画效果集合中放置的位置。 默认值为 -1（添加到末尾）。
    # effect = sequence.AddEffect(shape, 1, 0, 0, -1)
    # effect.Timing.Duration = 50.00

    return shape


if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # MsoTriState.msoTrue        1
    objPres = ppt.Presentations.Open("F:\ppt\\11.pptx", WithWindow=True)

    # Presentations.Slides 返回PPT的所有幻灯片集合
    listSlides = objPres.Slides

    print("========================== 幻灯片总长度：{}".format(len(listSlides)))
    print("========================== 开始循环幻灯片")

    slide_num = 1
    for slide in listSlides:

        # 运动序列
        # Sequence 对象 是 Effect 对象集合
        # Effect 对象代表动画序列的计时属性。
        sequence = slide.TimeLine.MainSequence

        list_voice = []
        len_effect = sequence.Count

        # for j in range(1, len_effect + 1):
        #     new_sequence.append(sequence.Item(j))

        print("========================== 运动序列总长度：{}".format(len_effect))

        for effect_num in range(1, len_effect+1):
            effect = sequence.Item(effect_num)
            effect_shape = effect.Shape
            animationSettings = effect_shape.AnimationSettings
            print("============== 进入序列 帧：{} 动画：{}===================".format(slide_num, effect_num))
            print("============== 当前操作的 shape id:{} ===================".format(effect_shape.Id))
            print("============== 当前操作的 shape Name:{} ===================".format(effect_shape.Name))
            print("============== 当前操作的 shape AnimationOrder:{} ===================".format(str(animationSettings.AnimationOrder)))
            print()
            # 指示指定形状的动画是仅在被单击时切换还是在经过指定时间后自动切换
            # 1 ：单击时播放
            # 2 ：在指定的一段时间后自动。
            animationSettings.AdvanceMode = 2
            animationSettings.AdvanceTime = 50.0

            # 插入本序列音频文件
            # two_shape = add_voice(sequence, "F:\\ppt\\audio\\幻灯片3.JPG.wav", slide_num + 1)
            # list_voice.append(two_shape)
            effect_shape.AnimationSettings.SoundEffect.ImportFromFile("F:\\ppt\\audio\\幻灯片3.JPG.wav")


        # 插入本序列音频文件
        first_shape = add_voice(sequence, "F:\\ppt\\audio\\幻灯片3.JPG.wav", 1)
        first_shape.AnimationSettings.AnimationOrder = 1
        # --------- 以下设置 可以指定幻灯片的播放效果，如指定时长自动播放 --------------------
        # 一下
        # 設置幻灯片的切换效果
        # 设置幻灯片在经过指定时间后是否自动切换 -1 表示 True
        slide.SlideShowTransition.AdvanceOnTime = True
        # 设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换
        # slide.SlideShowTransition.AdvanceTime = 0.0
        # -----------------------------------------------------------------------------

        slide_num += 1

        # i = 1
        # for v in list_voice:
        #     v.AnimationSettings.AnimationOrder = i
        #     i += 2

    print()
    print()
    print("=================== 打印结果验证  ======================")
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
            print("===================  ======================")
            print("===================  ======================")
            print("===================  ======================")
            print()

    objPres.CreateVideo("F:\ppt\\12345678.mp4", True, 5, 320, 24, 60)