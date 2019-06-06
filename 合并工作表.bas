Attribute VB_Name = "威创工作室VB模块"
Sub 合并工作表()

   '宏说明:合并同一工作簿的多个工作表数据到同一工作表, 格式不同也可合并.
   Set NewSheet = Sheets.Add(Type:=xlWorksheet) '生成一个新表
   Sheets(NewSheet.Index).Move Before:=Sheets(1) '将此新表移动到最前面
   For i = 2 To Worksheets.Count
   
   Dim x As Integer
       x = 1        'x=2时 合并后数据空一行; x=1时无空行.
   Sheets(i).UsedRange.Copy NewSheet.Cells([a65536].End(xlUp).Row + x, 1) '将其他表的已使用区域复制到新表中
   Next i
   MsgBox "合并完成"
End Sub

