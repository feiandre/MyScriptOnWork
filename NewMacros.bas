Attribute VB_Name = "NewMacros"
Sub ��4()
Attribute ��4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��4"
'
' ��4 ��
'
'
    Selection.TypeParagraph
    Selection.TypeText Text:="�����ٷ���˹�ٷ�"
    Selection.Font.Shrink
    Selection.Font.Name = "Arial Unicode MS"
    ActiveDocument.Save
End Sub


End Sub

Dim wdSty$, strTxt$
    wdSty = "���� 1"
    With Selection
        .HomeKey Unit:=wdStory, Extend:=wdMove '����Ƶ��ĵ���
        .Find.ClearFormatting
        .Find.Style = ActiveDocument.Styles(wdSty) '���ò����ı�����ʽΪwdSty(������1��)
    End With
'ѭ�������ĵ�������Ϊ������1����ʽ�Ķ��䣬
    Do While Selection.Find.Execute(findtext:="*^13", MatchWildcards:=True, Format:=True)
        strTxt = Selection.Text '��ȡ������ʽ���ı�
     '.......������¼�봦�����

        Selection.Move Unit:=wdWord, Count:=1
        If Selection.MoveRight <> 1 Then '�ĵ�β�˳�
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

If Selection.Style = "����" Then

End If

If Selection.Style = "���� 1" Then
    
    Selection.Font.Shrink
    Selection.Font.Name = "����"
    Selection.Font.Size = 20
    Selection.Font.Bold = True '�Ӵ�
   End If
Next

ActiveDocument.Save

End Sub


Sub ��������TimesNewRoman()
Attribute ��������TimesNewRoman.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��7"

'   Selection.Font.Shrink               ' ��������
'   Selection.Font.Size = 12            ' �ֺŴ�С
'   Selection.Font.Name = "����"    ' �����ͺ�
    
    Selection.Font.NameAscii = "Times New Roman"      ' ������������
    Selection.Font.NameOther = "Times New Roman"
'   Selection.Font.Bold = True          '�Ӵ�
    
    ActiveDocument.Save
    
End Sub
