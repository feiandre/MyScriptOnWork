Attribute VB_Name = "����������VBģ��"
Sub �ϲ�������()

   '��˵��:�ϲ�ͬһ�������Ķ�����������ݵ�ͬһ������, ��ʽ��ͬҲ�ɺϲ�.
   Set NewSheet = Sheets.Add(Type:=xlWorksheet) '����һ���±�
   Sheets(NewSheet.Index).Move Before:=Sheets(1) '�����±��ƶ�����ǰ��
   For i = 2 To Worksheets.Count
   
   Dim x As Integer
       x = 1        'x=2ʱ �ϲ������ݿ�һ��; x=1ʱ�޿���.
   Sheets(i).UsedRange.Copy NewSheet.Cells([a65536].End(xlUp).Row + x, 1) '�����������ʹ�������Ƶ��±���
   Next i
   MsgBox "�ϲ����"
End Sub

