Attribute VB_Name = "NewMacros"
Sub 宏4()
Attribute 宏4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏4"
'
' 宏4 宏
'
'
    Selection.TypeParagraph
    Selection.TypeText Text:="到三顿饭阿斯顿发"
    Selection.Font.Shrink
    Selection.Font.Name = "Arial Unicode MS"
    ActiveDocument.Save
End Sub


End Sub

Dim wdSty$, strTxt$
    wdSty = "标题 1"
    With Selection
        .HomeKey Unit:=wdStory, Extend:=wdMove '光标移到文档首
        .Find.ClearFormatting
        .Find.Style = ActiveDocument.Styles(wdSty) '设置查找文本的样式为wdSty(“标题1”)
    End With
'循环查找文档里所有为“标题1”样式的段落，
    Do While Selection.Find.Execute(findtext:="*^13", MatchWildcards:=True, Format:=True)
        strTxt = Selection.Text '获取符合样式的文本
     '.......在这里录入处理代码

        Selection.Move Unit:=wdWord, Count:=1
        If Selection.MoveRight <> 1 Then '文档尾退出
            Exit Do
        Else
            Selection.MoveLeft
        End If
    Loop
    
    
Sub Test()
Dim i As Single
For i = 1 To ActiveDocument.BuiltInDocumentProperties(wdPropertyLines).Value

With Selection
    .GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=i
    .HomeKey Unit:=wdLine
    .EndKey Unit:=wdLine, Extend:=wdExtend
End With

If Selection.Style = "正文" Then

End If

If Selection.Style = "标题 1" Then
    
    Selection.Font.Shrink
    Selection.Font.Name = "黑体"
    Selection.Font.Size = 20
    Selection.Font.Bold = True '加粗
   End If
Next

ActiveDocument.Save

End Sub


Sub 西文字体TimesNewRoman()
Attribute 西文字体TimesNewRoman.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏7"

'   Selection.Font.Shrink               ' 段落收缩
'   Selection.Font.Size = 12            ' 字号大小
'   Selection.Font.Name = "宋体"    ' 字体型号
    
    Selection.Font.NameAscii = "Times New Roman"      ' 西文字体设置
    Selection.Font.NameOther = "Times New Roman"
'   Selection.Font.Bold = True          '加粗
    
    ActiveDocument.Save
    
End Sub
