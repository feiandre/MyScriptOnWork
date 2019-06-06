Attribute VB_Name = "模块1"

Sub ImageEdite()
       '此宏作用: 修改PPT单页图片的大小及位置 1厘米=28.4像素
For i = 1 To ActivePresentation.Slides.Count

ActivePresentation.Slides(i).Select

ActiveWindow.Selection.SlideRange.Shapes("Picture 1").Select
    With ActiveWindow.Selection.ShapeRange
        
        '此行设定图片取消锁定纵横比
        .LockAspectRatio = msoFalse
        
        '1厘米=28.4像素
        .Height = 378
        .Width = 619
        .Left = 54
        .Top = 128
    End With
    Next
End Sub
