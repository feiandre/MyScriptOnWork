Attribute VB_Name = "模块2"
Private Sub Workbook_Open()
    
    Application.Width = 400
    Application.Height = 400
    Application.Left = 100
    Application.Top = 100
    
End Sub


//启动时设置窗口大小  好用
Private Sub Workbook_Open()
With ActiveWindow
.WindowState = xlNormal
.Width = 663
.Height = 453
.Left = 45
.Top = 25
.EnableResize = False
End With
End Sub

Sub SetGameWindow()
    Dim UsedW As Single, UsedH As Single
    Dim ViewRange As Range
    Set ViewRange = Range("A1:I9")
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
        .WindowState = xlNormal
        UsedW = .Width - .UsableWidth
        UsedH = .Height - .UsableHeight
        .Width = ViewRange.Width + UsedW
        .Height = ViewRange.Height + UsedH
        .ScrollRow = 1
        .ScrollColumn = 1
        .ActiveSheet.ScrollArea = ViewRange.Address
        .EnableResize = False
    End With
End Sub

//设置启动窗口大小  可用 但代码较繁琐
Private Sub Workbook_Open()
  
  Call SetGameWindow

End Sub
Private Sub SetGameWindow()
    Dim UsedW As Single, UsedH As Single
    Dim ViewRange As Range
    Set ViewRange = Range("A1:E8")
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
        .WindowState = xlNormal
        UsedW = .Width - .UsableWidth
        UsedH = .Height - .UsableHeight
        .Width = ViewRange.Width + UsedW
        .Height = ViewRange.Height + UsedH
        .ScrollRow = 1
        .ScrollColumn = 1
        .ActiveSheet.ScrollArea = ViewRange.Address
        .EnableResize = False
    End With
End Sub

