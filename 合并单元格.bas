Attribute VB_Name = "ģ��1"
Sub �ϲ���Ԫ��()
'VBA�ű��Զ����ºϲ���ͬ���ݵ�һ�е�Ԫ����ʵλ��ΪA1��Ԫ�����ϲ���Ԫ���в����пա�


Application.DisplayAlerts = 0       '���ε�������

Dim sr As Integer
Dim sc As Integer
Dim er As Integer
Dim ec As Integer
Dim i As Integer

    sr = 1           '��ʼ����һ����Ԫ��������    ��1��
    sc = 1           '��ʼ����һ����Ԫ��������    ��1��
    er = sr + 1
    ec = sc
        
    Do While (Not (Cells(er, ec).Value = ""))
      If Cells(sr, sc) = Cells(er, ec) Then
      Range(Cells(sr, sc), Cells(er, ec)).MergeCells = True
      Range(Cells(sr, sc), Cells(er, ec)).HorizontalAlignment = xlCenter
      Range(Cells(sr, sc), Cells(er, ec)).VerticalAlignment = xlCenter
      Else: sr = er     '��ʼ����㵥Ԫ��
      End If
         
      er = er + 1
      
    Loop

Application.DisplayAlerts = 1      '�ָ���������

End Sub
