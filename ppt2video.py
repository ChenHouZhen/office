import win32com
from win32com.client import Dispatch
import time
import math

if __name__ == '__main__':
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # MsoTriState.msoTrue        1
    objPres = ppt.Presentations.Open("F:\ppt\\123456.pptx", WithWindow=True)

    # Presentations.Slides 返回PPT的所有幻灯片集合
    listSlides = objPres.Slides
    i = 1
    for slide in listSlides:
        # 引入音频文件
        # slide.SlideShowTransition.SoundEffect.ImportFromFile("‪F:\\ppt\\baidu.mp3")
        # 指定是否循环播放为指定幻灯片切换所设置的声音，直到下一个声音开始。 False ：循环播放直到停止  True :跨幻灯片时播放

        # 插入媒体文件
        # 第一个参数：要添加的文件的名称。
        # 第二个参数： 指示是否链接到文件。
        # 第三个参数：指示是否随文档一起保存媒体。
        # 第四个参数：从幻灯片左边缘到媒体对象左边缘的距离（以磅为单位）。
        # 第五个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
        # 第六个参数： 媒体对象的宽度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
        # 第七个参数： 媒体对象的高度（以磅为单位）。 默认值为 -1。 指定 0 隐藏图标
        shape = slide.Shapes.AddMediaObject2("F:\\ppt\\audio\\幻灯片{}.JPG.wav".format(i), 0, -1, 1120, 680, 0, 0)
        # 插入图片
        # shape = slide.Shapes.AddPicture2("F:\\aa.jpg",0,-1,12,12)

        # --------------设置 音频为自动播放 和 音频图标是否隐藏，如果不隐藏，设置音频图标-------------------
        # 获取音频总时长
        # 1912 表示 1秒912分
        audio_time = shape.MediaFormat.Length
        print("音頻原始时长："+str(audio_time))
        audio_time = math.ceil(audio_time / 1000)
        print("转换后的时长："+str(audio_time))

        # PlaySettings对象 关于指定的媒体剪辑在幻灯片放映中的播放方式的信息。
        # 指定的影片或声音后是否在激活后自动播放。 -1 为 true
        shape.AnimationSettings.PlaySettings.PlayOnEntry = True

        # 幻灯片放映期间指定媒体剪辑在不播放时是否隐藏。幻灯片放映过程中指定媒体隐藏。
        shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = False

        # 设置音频图标
        shape.MediaFormat.SetDisplayPictureFromFile('F:/ppt/voice15.png')
        shape.Width = 30
        shape.Height = 30
        # --------------------------------------------------

        # print(shape.width)
        # 获取或设置媒体的剪裁区域终点的时间
        # print("媒体裁剪最终时间："+str(shape.MediaFormat.EndPoint))
        # print("媒体裁剪开始时间："+str(shape.MediaFormat.StartPoint))

        print("是否嵌入："+str(shape.MediaFormat.IsEmbedded))
        print("是否链接媒体文件："+str(shape.MediaFormat.IsLinked))
        print("媒体总长度："+str(shape.MediaFormat.Length))
        print("是否静音："+str(shape.MediaFormat.Muted))

        i += 1
        # 设置音频的音量大小
        # shape.MediaFormat.Volume = 0

        # --------- 以下设置 可以指定幻灯片的播放效果，如指定时长自动播放 --------------------
        # 一下
        # 設置幻灯片的切换效果
        # 设置幻灯片在经过指定时间后是否自动切换 -1 表示 True
        slide.SlideShowTransition.AdvanceOnTime = True
        # 设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换
        slide.SlideShowTransition.AdvanceTime = audio_time

        # 测试效果：可以实现指定幻灯片按指定时长播放
        # -----------------------------------------------------------------------------

        # 运动序列
        # Sequence 对象 是 Effect 对象集合
        # Effect 对象代表动画序列的计时属性。
        sequence = slide.TimeLine.MainSequence

        for effect in sequence:
            print("进入动画序列。。。。。。。")
            # 指示指定形状的动画是仅在被单击时切换还是在经过指定时间后自动切换
            # 1 ：单击时播放
            # 2 ：在指定的一段时间后自动。
            effect.Shape.AnimationSettings.AdvanceMode = 2
            print("动画名称：" + effect.DisplayName)
            # effect.Shape.AnimationSettings.AdvanceTime = 5
            # effect.Shape.AnimationSettings.TextLevelEffect = 16
            # effect.Shape.AnimationSettings.Animate = True
            # print("时间" + str(time.strftime("%Y-%m-%d %H:%M:%S" ,time.localtime())))
            # 指定动画的持续时间
            # duration = 5
            # effect.Timing.Duration = 2







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
