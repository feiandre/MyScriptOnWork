Attribute VB_Name = "模块1"
Function crmfalse(Rng As Range) As String
          
          Dim 字体颜色 As Integer
          Dim 背景颜色 As Integer
          
          Application.Volatile
          字体颜色 = Rng.Font.ColorIndex
          背景颜色 = Rng.Interior.ColorIndex
          
          If (字体颜色 = 3) Or (背景颜色 = 6) Then
          
               crmfalse = "错误"
          Else
               crmfalse = "正确"
               
          End If
          
        
End Function
        


