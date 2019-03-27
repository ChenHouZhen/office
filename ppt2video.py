import win32com
from win32com.client import Dispatch
import math


def add_voice(voice_path):
    print("==============  插入音频，文件名：{}==============".format(voice_path))

    # 插入媒体文件
    # 第一个参数：要添加的文件的名称。
    # 第二个参数： 指示是否链接到文件。
    # 第三个参数：指示是否随文档一起保存媒体。
    # 第四个参数：从幻灯片左边缘到媒体对象左边缘的距离（以磅为单位）。
    # 第五个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
    # 第六个参数： 媒体对象的宽度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
    # 第七个参数： 媒体对象的高度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
    shape = slide.Shapes.AddMediaObject2(voice_path, 0, -1, 1120, 680, 0, 0)
    # 插入图片
    # shape = slide.Shapes.AddPicture2("F:\\aa.jpg",0,-1,12,12)

    # --------------设置 音频为自动播放 和 音频图标是否隐藏，如果不隐藏，设置音频图标-------------------
    # 获取音频总时长
    # 1912 表示 1秒912分
    audio_time = shape.MediaFormat.Length
    print("音頻原始时长：" + str(audio_time))
    audio_time = math.ceil(audio_time / 1000)
    print("转换后的时长：" + str(audio_time))

    # PlaySettings对象 关于指定的媒体剪辑在幻灯片放映中的播放方式的信息。
    # 指定的影片或声音后是否在激活后自动播放。 -1 为 true
    shape.AnimationSettings.PlaySettings.PlayOnEntry = True

    # 幻灯片放映期间指定媒体剪辑在不播放时是否隐藏。幻灯片放映过程中指定媒体隐藏。
    shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = False

    # -----------------下面这两句不一定需要 下面两句 对应： 动画->开始 ----------------
    shape.AnimationSettings.AdvanceMode = 2
    shape.AnimationSettings.AdvanceTime = 0
    # ---------------------------------------------------

    # 在指定媒体剪辑播放结束前是否暂停幻灯片放映
    # shape.AnimationSettings.PlaySettings.PauseAnimation =True
    # ------------------设置音频图标----------------------
    shape.MediaFormat.SetDisplayPictureFromFile('F:/ppt/voice15.png')
    shape.Width = 30
    shape.Height = 30
    # --------------------------------------------------

    # print(shape.width)
    # 获取或设置媒体的剪裁区域终点的时间
    # print("媒体裁剪最终时间："+str(shape.MediaFormat.EndPoint))
    # print("媒体裁剪开始时间："+str(shape.MediaFormat.StartPoint))

    print("是否嵌入：" + str(shape.MediaFormat.IsEmbedded))
    print("是否链接媒体文件：" + str(shape.MediaFormat.IsLinked))
    print("媒体总长度：" + str(shape.MediaFormat.Length))
    print("是否静音：" + str(shape.MediaFormat.Muted))

    return shape


if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # MsoTriState.msoTrue        1
    objPres = ppt.Presentations.Open("F:\ppt\\12345678.pptx", WithWindow=True)

    # Presentations.Slides 返回PPT的所有幻灯片集合
    listSlides = objPres.Slides

    for slide in listSlides:

        i = 1
        # 设置音频的音量大小
        # shape.MediaFormat.Volume = 0

        # 运动序列
        # Sequence 对象 是 Effect 对象集合
        # Effect 对象代表动画序列的计时属性。
        sequence = slide.TimeLine.MainSequence

        len_effect = sequence.Count


        # 插入本序列音频文件
        video_shape = add_voice("F:\\ppt\\audio\\幻灯片{}.JPG.wav".format(i))
        video_shape.AnimationSettings.PlaySettings.PlayOnEntry = True
        video_shape.AnimationSettings.PlaySettings.LoopUntilStopped = False

        print("排序前，动画循序：AnimationOrder :"+ str(video_shape.AnimationSettings.AnimationOrder))
        video_shape.AnimationSettings.AnimationOrder = 0
        print("排序后，动画循序：AnimationOrder :" + str(video_shape.AnimationSettings.AnimationOrder))


        video_shape.AnimationSettings.AdvanceMode = 2
        video_shape.AnimationSettings.AdvanceTime = 0
        video_shape_id = video_shape.Id

        print("插入的音频id "+str(video_shape_id))
        print("============== 帧：{} 动画总长度:：{}==============".format(i, len_effect))

        slide_time = 2

        for j in range(1, len_effect+1):
            animation_order = 1
            print("============== 进入动画序列 ===================")
            print("============== 进入序列 帧：{} 动画：{}===================".format(i, j))

            effect = sequence.Item(j)

            effect_shape = effect.Shape
            print("动画{},AnimationOrder".format(j)+str(effect_shape.AnimationSettings.AnimationOrder))
            effect_shape.AnimationSettings.AnimationOrder = 1
            animation_order += 1
            # 插入本序列音频文件
            video_shape = add_voice("F:\\ppt\\audio\\幻灯片{}.JPG.wav".format(i))
            video_shape.AnimationSettings.AnimationOrder = 2
            animation_order += 1
            print("动画声音{},AnimationOrder".format(j)+str(video_shape.AnimationSettings.AnimationOrder))

            try:
                effect_shape.AnimationSettings
            except AttributeError as e:
                print("============== 异常 ============== ")
                continue

            print("============== 设置动画 ============== ")
            animationSettings = effect_shape.AnimationSettings

            # 指示指定形状的动画是仅在被单击时切换还是在经过指定时间后自动切换
            # 1 ：单击时播放
            # 2 ：在指定的一段时间后自动。
            animationSettings.AdvanceMode = 2

            # 如果是音频，不延迟
            # if effect.Shape == shape:
            #     print("============= 音频 跳过 ================")
            #     continue
            # print("动画名称：" + effect.DisplayName)
            animationSettings.AdvanceTime = 3
            # effect.Shape.AnimationSettings.TextLevelEffect = 16
            # effect.Shape.AnimationSettings.Animate = True
            # print("时间" + str(time.strftime("%Y-%m-%d %H:%M:%S" ,time.localtime())))
            # 指定动画的持续时间
            # duration = 5
            # effect.Timing.Duration = 2
            i += 1

        # --------- 以下设置 可以指定幻灯片的播放效果，如指定时长自动播放 --------------------
        # 一下
        # 設置幻灯片的切换效果
        # 设置幻灯片在经过指定时间后是否自动切换 -1 表示 True
        slide.SlideShowTransition.AdvanceOnTime = True
        # 设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换
        slide.SlideShowTransition.AdvanceTime = 10
        print("第 %d 帧动画延迟：%d" % (i, slide_time))
        # 测试效果：可以实现指定幻灯片按指定时长播放
        # -----------------------------------------------------------------------------



    # 幻灯片播放模式
    # objPres.SlideShowSettings.RangeType = 2
    # objPres.SlideShowSettings.StartingSlide = 2
    # objPres.SlideShowSettings.EndingSlide = 4
    # objPres.SlideShowSettings.AdvanceMode =  2
    #

    # 幻灯片运动序列
    # objSequence = objSlide.TimeLine.MainSequence

    #
    # for each in objSequence:
    #     each.Shape.AnimationSettings.Animate = 1
    #     each.Shape.AnimationSettings.SoundEffect.Type = 2
    #     each.Shape.SlideShowSettings.SoundEffect.ImportFromFile("‪F:\\ppt\\baidu.mp3")
    #     each.MainSequence(1).Timing.Duration = 10
    #     # PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious  3
    #     each.Timing.TriggerType = 3


    # count = objSequence.Count

    #  PowerPoint.PpSaveAsFileType.ppSaveAsMP4  39
    #  MsoTriState.msoTriStateMixed -2
    # objPres.SaveAs("F:\ppt\\output03.mp4", 39)



    # 第二个参数 是否使用计时和旁白。
    # 第三个参数 幻灯片的持续时间（秒）。
    # 第四个参数 幻灯片的分辨率。
    # 第五个参数 每秒的帧数。
    # 第六个参数 幻灯片的质量级别
    # objPres.CreateVideo("F:\ppt\\123456.mp4", True, 5, 320, 24, 60)
