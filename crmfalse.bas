Attribute VB_Name = "ģ��1"
Function crmfalse(Rng As Range) As String
          
          Dim ������ɫ As Integer
          Dim ������ɫ As Integer
          
          Application.Volatile
          ������ɫ = Rng.Font.ColorIndex
          ������ɫ = Rng.Interior.ColorIndex
          
          If (������ɫ = 3) Or (������ɫ = 6) Then
          
               crmfalse = "����"
          Else
               crmfalse = "��ȷ"
               
          End If
          
        
End Function
        


