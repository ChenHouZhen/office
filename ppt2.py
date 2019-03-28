import sys

def make_video(self):
    """当前幻灯片总数"""
    count = self.objPres.Slides.Count
    i = 0
    """循环每个幻灯片 插入对应音频文件"""
    while i < count:
        # 插入媒体文件
        # 第一个参数：要添加的文件的名称。
        # 第二个参数： 指示是否链接到文件。
        # 第三个参数：指示是否随文档一起保存媒体。
        # 第四个参数：从幻灯片左边缘到媒体对象左边缘的距离（以磅为单位）。
        # 第五个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
        # 第六个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
        # 第七个参数： 从幻灯片上边缘到媒体对象上边缘的距离（以磅为单位）。
        """当前幻灯片的动画数量总数"""
        animateCount = self.objPres.Slides[i].TimeLine.MainSequence.Count
        if (animateCount == 0):
            self.objPres.Slides[i].Shapes.AddMediaObject2(sys.path[0] + "/statics/audio/audio%d.wav" % (i + 1), 0, -1,
                                                          0, 0, 0, 0)
        else:
            j = 0
            while j < self.objPres.Slides[i].Shapes.Count:
                """设为1秒后自动播放动画"""
                self.objPres.Slides[i].Shapes[j].AnimationSettings.AdvanceTime = 1
                self.objPres.Slides[i].Shapes[j].AnimationSettings.SoundEffect.ImportFromFile(
                    sys.path[0] + "/statics/audio/audio%d.%d.wav" % (i + 1, j + 1))
                j += 1
        i += 1
    # 第二个参数 是否使用计时和旁白。
    # 第三个参数 幻灯片的持续时间（秒）。
    # 第四个参数 幻灯片的分辨率。
    # 第五个参数 每秒的帧数。
    # 第六个参数 幻灯片的质量级别
    self.objPres.CreateVideo(self.videoPath, True, 5, 1080, 24, 60)



if __name__ == '__main__':
    make_video()