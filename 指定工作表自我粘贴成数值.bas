Attribute VB_Name = "模块1"
Function getR(S)
 
 Dim i As Integer
 Dim m As String
 m = "-"
 
 i = InStr(S, m)
 getR = Right(S, Len(S) - i)
 
End Function



Sub 指定工作表自我粘贴成数值()

'该宏将表名包含“分产品线达成揭示”的工作表中的公式转换为数值
   
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


