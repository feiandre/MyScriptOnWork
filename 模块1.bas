Attribute VB_Name = "模块1"

Sub Macro1()
   
   Dim i As Integer
  
  For i = 1 To Sheets.Count

'判断工作表

   Sheets(i).Select
   If (InStr(Sheets(i).Name, "分产品线达成揭示")) Then

'  工作表复制粘贴

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

