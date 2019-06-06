Attribute VB_Name = "模块3"
Sub 月份纵列()
Attribute 月份纵列.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
'

'
    Dim c As Integer
    Dim r As Integer
    Dim ci As Integer
    Dim i As Integer
    
    c = 4
    r = 4
    ci = 1
  For i = 1 To 11
   
    Range(Cells(4, c), Cells(43, c)).Select
    Selection.Copy
    
   Cells(r + 40 * ci, 3).Select
    ActiveSheet.Paste
   c = c + 1
   ci = ci + 1
   
   Next i
   
   
   c = 17
    r = 4
    ci = 1
  For i = 1 To 11
   
    Range(Cells(4, c), Cells(43, c)).Select
    Selection.Copy
    
   Cells(r + 40 * ci, 16).Select
    ActiveSheet.Paste
   c = c + 1
   ci = ci + 1
   
   Next i
   
   
   
   
End Sub

Sub 复制粘贴()

    Sheets("云服务签单合并").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 2).Select
    ActiveSheet.Paste
    
    Sheets("云服务签单合并").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 3).Select
    ActiveSheet.Paste
        
    Sheets("云服务回款合并").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 4).Select
    ActiveSheet.Paste
    
    Sheets("云服务回款合并").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 5).Select
    ActiveSheet.Paste
    
    Sheets("电子政务签单").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 7).Select
    ActiveSheet.Paste
    
    Sheets("电子政务签单").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 8).Select
    ActiveSheet.Paste
        
   Sheets("电子政务回款").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 9).Select
    ActiveSheet.Paste
    
    Sheets("电子政务回款").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 10).Select
    ActiveSheet.Paste
    
    Sheets("大数据签单").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 12).Select
    ActiveSheet.Paste
    
    Sheets("大数据签单").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 13).Select
    ActiveSheet.Paste
        
   Sheets("大数据回款").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 14).Select
    ActiveSheet.Paste
    
    Sheets("大数据回款").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 15).Select
    ActiveSheet.Paste
    
End Sub
