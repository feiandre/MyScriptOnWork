Sub 替换顿号为分号()

'测试成功
'选中邮件正文中的需要替换顿号的文字,替换为分号,然后置入剪切板,以备复制粘贴

    Dim Inspector As Outlook.Inspector
    Dim wdDoc As Word.Document
    Dim Selection As Word.Selection
    Dim tt As String
    Dim olReply As MailItem
    Dim olItem As Outlook.MailItem
    
    'Set olItem = Application.ActiveExplorer
    Set Inspector = Application.ActiveInspector()
    Set wdDoc = Inspector.WordEditor
    Set Selection = wdDoc.Application.Selection

   'Application.DisplayAlerts = 0
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    tt = Trim(tt)
    tt = Replace(Selection.Text, "、", ";")
     
   
   '重名人员替换为邮箱，避免每次手动选择，再有重名人员可继续添加；以及替换掉一些多余的文字
    tt = Replace(tt, "张强", "zhangqiang@inspur.com;")
    tt = Replace(tt, "张楠", "zhang_nan@inspur.com;")
    tt = Replace(tt, "王艳", "wangyanrj@inspur.com;")
    tt = Replace(tt, "王珊珊", "wangshanshanbj@inspur.com;")
    tt = Replace(tt, "张晓燕", "zhangxy1@inspur.com;")
    tt = Replace(tt, "郑琰", "zheng.yan@inspur.com;")
    tt = Replace(tt, "需知晓：", " ;")
    tt = Replace(tt, "需知晓:", "; ")
    tt = Replace(tt, "王明飞", " wangmf@inspur.com;")
    tt = Replace(tt, "张倩", " zhangqiancsg@inspur.com;")
    tt = Replace(tt, "赵飞", " zhaofei01@inspur.com;")
    tt = Replace(tt, "王冲", " wangchongsy@inspur.com;")
    tt = Replace(tt, "朱勇", " zhuyong@inspur.com;")
    tt = Replace(tt, "王方", " wangfang@inspur.com;")
    tt = Replace(tt, "张峰", " zhangfengoa@inspur.com;")
    tt = Replace(tt, "丁杨", " dingyang@inspur.com;")
    tt = Replace(tt, "王璐", " wanglu1@inspur.com;")
    tt = Replace(tt, "王君", " wangjunyfw@inspur.com;")
    'tt = Replace(tt, "贺蕾", " helei02@inspur.com;")
    
    tt = Replace(tt, "/", ";")
    tt = Replace(tt, "\", ";")
   ' tt = Replace(tt, vbCrLf, " ;")
    tt = Replace(tt, Chr(10), ";")
    tt = Replace(tt, Chr(13), ";")
    tt = Replace(tt, Chr(7), ";")
 
        
    'Selection.Find.Execute Replace:=wdReplaceAll
    
    Dim myData As DataObject
    Set myData = New DataObject
    With myData
        .SetText tt                    'SetText方法设置对象的文本内容
        .PutInClipboard                            '把对象置入System Clipboard
    End With
    
    
   ' Application.DisplayAlerts = 1
    'MsgBox CStr(tt)
    Inspector.CurrentItem.CC = Inspector.CurrentItem.CC + ";" + tt
    Inspector.CurrentItem.Recipients.ResolveAll   '检查姓名
    
End Sub

Sub ReplyMSG()
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply
    Dim ebody As String

    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.ReplyAll
            olReply.Subject = Replace(olReply.Subject, "答复:", "【评审提醒】")
           ' olReply.HTMLBody = "领导，    您好。此次评审还需要您回复评审意见，敬请根据评审编号及时在原始评审邮件流程中全部答复并回复评审意见。祝好 " & vbCrLf & olReply.HTMLBody
           'olReply.HTMLBody = "<p><span style='font-size:13pt;  font-family:微软雅黑,sans-serif'>郭经理，</span></p>" & vbCrLf & olReply.HTMLBody
           
           
            ebody = "<p style='padding=1;margin=0'><span style='font-size:13pt;  font-family:微软雅黑,sans-serif'>" & _
            "领导，" & _
            "<br/>&nbsp&nbsp&nbsp&nbsp&nbsp " & _
            "您好。此次评审还需要您回复评审意见，敬请根据评审编号及时在原始评审邮件流程中全部答复并回复评审意见。" & _
            "<br/>&nbsp&nbsp&nbsp&nbsp&nbsp " & _
            "提示！务必不要在此邮件基础上回复，此邮件仅为评审提醒。请知悉。" & _
            "<br/>&nbsp&nbsp&nbsp&nbsp&nbsp " & _
            "祝好。" & _
            "</span></p>"

            
            
            olReply.CC = ""
            olReply.HTMLBody = ebody & vbCrLf & olReply.HTMLBody
            olReply.FlagRequest = "请您评审"
        
            ' We set the due date for the reminder two days from today
            'olReply.FlagDueBy = Now + (2 / 60 / 24)
            
            'DateAdd("d", 2, Date)
            
            
            olReply.ReminderSet = True
            'Set a custom reminder time
            olReply.ReminderTime = Now + (3 / 60 / 24)
            olReply.Display

            'olReply.Send
        Exit For
        
    Next
    
End Sub
Private Sub Outlook_Open()
        Application.OnKey "{F2}", "评审意见2Excel"
End Sub
    
Sub 评审意见toExcel()


    Dim olItem As Outlook.MailItem

    For Each olItem In Application.ActiveExplorer.Selection
    'For Each olItem In Application.CurrentFolder.Selection
    'Code here

    Dim objMail As Outlook.MailItem
    
    Dim strExcelFile As String
    Dim objExcelApp As Excel.Application
    Dim objExcelWorkBook As Excel.Workbook
    Dim objExcelWorkSheet As Excel.Worksheet
    Dim nNextEmptyRow As Integer
    Dim strColumnB As String
    Dim strColumnC As String
    Dim strColumnD As String
    Dim strColumnE As String
    
    Set objMail = olItem
    
    'If Item.Class = olMail Then
    '   Set objMail = Item
    'End If
 
    'Specify the Excel file which you want to auto export the email list
    'You can change it as per your case
    strExcelFile = "D:\Inspur Files\邮件记录.xlsx"
 
    'Get Access to the Excel file
    On Error Resume Next
    Set objExcelApp = GetObject(, "Excel.Application")
    If Error <> 0 Then
       Set objExcelApp = CreateObject("Excel.Application")
    End If
    Set objExcelWorkBook = objExcelApp.Workbooks.Open(strExcelFile)
    Set objExcelWorkSheet = objExcelWorkBook.Sheets("Sheet2")
 
    'Get the next empty row in the Excel worksheet
    nNextEmptyRow = objExcelWorkSheet.Range("A" & objExcelWorkSheet.Rows.Count).End(xlUp).Row + 1
 
    'Specify the corresponding values in the different columns
    strColumnB = objMail.SenderName
    'strColumnC = objMail.SenderEmailAddress
    strColumnD = objMail.Subject
    strColumnE = objMail.ReceivedTime
 
    'Add the vaules into the columns
    objExcelWorkSheet.Range("A" & nNextEmptyRow) = nNextEmptyRow - 1
    objExcelWorkSheet.Range("B" & nNextEmptyRow) = strColumnB
    objExcelWorkSheet.Range("C" & nNextEmptyRow) = strColumnC
    objExcelWorkSheet.Range("D" & nNextEmptyRow) = strColumnD
    objExcelWorkSheet.Range("E" & nNextEmptyRow) = strColumnE
 
    'Fit the columns from A to E
    'objExcelWorkSheet.Columns("A:E").EntireColumn.AutoFit
 
    'Save the changes and close the Excel file
    objExcelWorkBook.Close SaveChanges:=True



    'Code here
    
    
    
        Exit For
    Next
    
End Sub
Sub GetSelectedItems()
 
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim MsgTxt As String
 Dim x As Integer
  
 
 MsgTxt = "You have selected items from: "
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 
 For x = 1 To myOlSel.Count
 MsgTxt = MsgTxt & myOlSel.Item(x).SenderName & ";" & myOlSel.Item(x).Subject
 
 Next x
 
 MsgBox MsgTxt
 
End Sub
Sub 评审意见2Excel()

    '这个脚本成功！！！
    Dim olItem As Outlook.MailItem

For Each olItem In Application.ActiveExplorer.Selection
    'For Each olItem In Application.CurrentFolder.Selection
    'Code here
    
    Dim xMailItem As Outlook.MailItem
    Dim xExcelFile As String
    Dim xExcelApp As Excel.Application
    Dim xWb As Excel.Workbook
    Dim xWs As Excel.Worksheet
    Dim xNextEmptyRow As Integer
    
    'On Error Resume Next
    'If Item.Class <> olMail Then Exit Sub
    
    Set xMailItem = olItem
    
    xExcelFile = "D:\Program\Nutstore\Data\坚果云同步\FY2018生态合作部协议评审.xlsm"
    If IsWorkBookOpen(xExcelFile) = True Then
        Set xExcelApp = GetObject(, "Excel.Application")
        Set xWb = GetObject(xExcelFile)
        If Not xWb Is Nothing Then xWb.Close True
    Else
        Set xExcelApp = New Excel.Application
    End If
    Set xWb = xExcelApp.Workbooks.Open(xExcelFile)
    Set xWs = xWb.Sheets("评审人答复")
    xNextEmptyRow = xWs.Range("A" & xWs.Rows.Count).End(xlUp).Row + 1
    
    On Error Resume Next
    
    '正则表达式匹配发件人
    s = xMailItem.SenderName
    Set re = CreateObject("vbscript.regexp")
    
    re.Pattern = "\(([^)]+)"
    Set ms = re.Execute(s)
    mailsender = ms(0).submatches(0)
    

    '正则表达式匹配ICC或ICA
    ss = xMailItem.Subject
    Set re = CreateObject("vbscript.regexp")
    
    
    re.Pattern = "(ICC\d{8})"
    Set ms = re.Execute(ss)
    icc = ms(0).submatches(0)

    re.Pattern = "(ICA\d{8})"
    Set ms = re.Execute(ss)
    ica = ms(0).submatches(0)
    
    On Error GoTo 0
    
    If icc <> "" Then  '判断icc是为空
         ic = icc
    Else: ic = ica
    End If
    
    
    'MsgBox xMailItem.Recipients(2).Name
     
    With xWs
        .Cells(xNextEmptyRow, 1) = xNextEmptyRow - 1
        .Cells(xNextEmptyRow, 2) = mailsender
        .Cells(xNextEmptyRow, 3) = ic
        .Cells(xNextEmptyRow, 4) = xMailItem.Subject
        .Cells(xNextEmptyRow, 5) = xMailItem.ReceivedTime
        '.Cells(xNextEmptyRow, 6) = xMailItem.Body
    End With
    'xWs.Columns("A:E").AutoFit
    xWb.Save
    'xWb.Close
    
       Exit For
   Next
   
    'Dim WshShell As Object
    'Set WshShell = CreateObject("Wscript.Shell")
    'WshShell.Popup "成功！", 1, "提示！"
  
  
End Sub
Function IsWorkBookOpen(FileName As String)
    Dim xFreeFile As Long, xErrNo As Long
    On Error Resume Next
    xFreeFile = FreeFile()
    Open FileName For Input Lock Read As #xFreeFile
    Close xFreeFile
    xErrNo = Err
    On Error GoTo 0
    Select Case xErrNo
        Case 0: IsWorkBookOpen = False
        Case 70: IsWorkBookOpen = True
        Case Else: Error xErrNo
    End Select
End Function
Sub 收件人2excel()

    '这个脚本成功
    Dim olItem As Outlook.MailItem

For Each olItem In Application.ActiveExplorer.Selection
    'For Each olItem In Application.CurrentFolder.Selection
    'Code here
    
    Dim xMailItem As Outlook.MailItem
    Dim xExcelFile As String
    Dim xExcelApp As Excel.Application
    Dim xWb As Excel.Workbook
    Dim xWs As Excel.Worksheet
    Dim xNextEmptyRow As Integer
    
    'On Error Resume Next
    'If Item.Class <> olMail Then Exit Sub
    
    Set xMailItem = olItem
    
    xExcelFile = "D:\Program\Nutstore\Data\坚果云同步\FY2018生态合作部协议评审.xlsm"
    If IsWorkBookOpen(xExcelFile) = True Then
        Set xExcelApp = GetObject(, "Excel.Application")
        Set xWb = GetObject(xExcelFile)
        If Not xWb Is Nothing Then xWb.Close True
    Else
        Set xExcelApp = New Excel.Application
    End If
    Set xWb = xExcelApp.Workbooks.Open(xExcelFile)
    Set xWs = xWb.Sheets("评审记录")
    xNextEmptyRow = xWs.Range("A" & xWs.Rows.Count).End(xlUp).Row + 1
    
    On Error Resume Next
    
    '正则表达式匹配发件人
    s = xMailItem.To
    Set re = CreateObject("vbscript.regexp")
    
    re.Pattern = "\(([^)]*)"
    re.Global = True
    Set mss = re.Execute(s)
    

    '正则表达式匹配ICC或ICA
    ss = xMailItem.Subject
    Set re = CreateObject("vbscript.regexp")
    
    
    re.Pattern = "(ICC\d{8})"
    Set ms = re.Execute(ss)
    icc = ms(0).submatches(0)

    re.Pattern = "(ICA\d{8})"
    Set ms = re.Execute(ss)
    ica = ms(0).submatches(0)
    
    On Error GoTo 0
    
    If icc <> "" Then  '判断icc是为空
         ic = icc
    Else: ic = ica
    End If
    
    i = 1
    For Each Item In mss
            'MsgBox mss(1).submatches(0)
            
            'MsgBox Item.submatches(0)
            xWs.Cells(xNextEmptyRow, i + 5) = Item.submatches(0)
            i = i + 1
   Next
     
    With xWs
        .Cells(xNextEmptyRow, 1) = xNextEmptyRow - 1
        .Cells(xNextEmptyRow, 2) = xMailItem.SenderName
        .Cells(xNextEmptyRow, 3) = ic
        .Cells(xNextEmptyRow, 4) = xMailItem.Subject
        .Cells(xNextEmptyRow, 5) = xMailItem.ReceivedTime
        '.Cells(xNextEmptyRow, 6) = xMailItem.Body
    End With
    'xWs.Columns("A:E").AutoFit
    xWb.Save
    'xWb.Close
    
       Exit For
   Next
   
    'Dim WshShell As Object
    'Set WshShell = CreateObject("Wscript.Shell")
    'WshShell.Popup "成功！", 1, "提示！"
  
  
End Sub
