Attribute VB_Name = "ģ��3"
Sub �·�����()
Attribute �·�����.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��1 ��
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

Sub ����ճ��()

    Sheets("�Ʒ���ǩ���ϲ�").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 2).Select
    ActiveSheet.Paste
    
    Sheets("�Ʒ���ǩ���ϲ�").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 3).Select
    ActiveSheet.Paste
        
    Sheets("�Ʒ���ؿ�ϲ�").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 4).Select
    ActiveSheet.Paste
    
    Sheets("�Ʒ���ؿ�ϲ�").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 5).Select
    ActiveSheet.Paste
    
    Sheets("��������ǩ��").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 7).Select
    ActiveSheet.Paste
    
    Sheets("��������ǩ��").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 8).Select
    ActiveSheet.Paste
        
   Sheets("��������ؿ�").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 9).Select
    ActiveSheet.Paste
    
    Sheets("��������ؿ�").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 10).Select
    ActiveSheet.Paste
    
    Sheets("������ǩ��").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 12).Select
    ActiveSheet.Paste
    
    Sheets("������ǩ��").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 13).Select
    ActiveSheet.Paste
        
   Sheets("�����ݻؿ�").Select
    Range("C4:C483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 14).Select
    ActiveSheet.Paste
    
    Sheets("�����ݻؿ�").Select
    Range("P4:P483").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Cells(2, 15).Select
    ActiveSheet.Paste
    
End Sub
