Attribute VB_Name = "ģ��1"
Function getR(S)
 
 Dim i As Integer
 Dim m As String
 m = "-"
 
 i = InStr(S, m)
 getR = Right(S, Len(S) - i)
 
End Function



Sub ָ������������ճ������ֵ()

'�ú꽫�����������ֲ�Ʒ�ߴ�ɽ�ʾ���Ĺ������еĹ�ʽת��Ϊ��ֵ
   
   Dim i As Integer
  
  For i = 1 To Sheets.Count

'�жϹ�����

   Sheets(i).Select
   If (InStr(Sheets(i).Name, "�ֲ�Ʒ�ߴ�ɽ�ʾ")) Then

'  ��������ճ��

    Sheets(i).Select
    Range("A1:CQ50").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
    End If
   
   Next i
   
   
   ActiveWorkbook.Save
   Sheets(1).Select
   Range("A1").Select
   
   
End Sub


