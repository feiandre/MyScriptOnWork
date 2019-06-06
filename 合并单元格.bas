Attribute VB_Name = "模块1"
Sub 合并单元格()
'VBA脚本自动向下合并相同内容的一列单元格，其实位置为A1单元格。所合并单元格中不能有空。


Application.DisplayAlerts = 0       '屏蔽弹出窗口

Dim sr As Integer
Dim sc As Integer
Dim er As Integer
Dim ec As Integer
Dim i As Integer

    sr = 1           '初始化第一个单元格行坐标    第1行
    sc = 1           '初始化第一个单元格列坐标    第1列
    er = sr + 1
    ec = sc
        
    Do While (Not (Cells(er, ec).Value = ""))
      If Cells(sr, sc) = Cells(er, ec) Then
      Range(Cells(sr, sc), Cells(er, ec)).MergeCells = True
      Range(Cells(sr, sc), Cells(er, ec)).HorizontalAlignment = xlCenter
      Range(Cells(sr, sc), Cells(er, ec)).VerticalAlignment = xlCenter
      Else: sr = er     '初始化起点单元格
      End If
         
      er = er + 1
      
    Loop

Application.DisplayAlerts = 1      '恢复弹出窗口

End Sub
