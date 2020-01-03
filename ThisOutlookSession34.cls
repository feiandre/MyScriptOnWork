VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents GMailItems As Outlook.Items
Attribute GMailItems.VB_VarHelpID = -1
Private Sub Application_Startup()
    Set GMailItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub GMailItems_ItemAdd(ByVal Item As Object)
    Dim xMailItem As Outlook.MailItem
    Dim xExcelFile As String
    Dim xExcelApp As Excel.Application
    Dim xWb As Excel.Workbook
    Dim xWs As Excel.Worksheet
    Dim xNextEmptyRow As Integer
    On Error Resume Next
    If Item.Class <> olMail Then Exit Sub
    Set xMailItem = Item
    xExcelFile = "D:\Inspur Files\ÓÊ¼þ¼ÇÂ¼.xlsx"
    If IsWorkBookOpen(xExcelFile) = True Then
        Set xExcelApp = GetObject(, "Excel.Application")
        Set xWb = GetObject(xExcelFile)
        If Not xWb Is Nothing Then xWb.Close True
    Else
        Set xExcelApp = New Excel.Application
    End If
    Set xWb = xExcelApp.Workbooks.Open(xExcelFile)
    Set xWs = xWb.Sheets(1)
    xNextEmptyRow = xWs.Range("A" & xWs.Rows.Count).End(xlUp).Row + 1
    With xWs
        .Cells(xNextEmptyRow, 1) = xNextEmptyRow - 1
        .Cells(xNextEmptyRow, 2) = xMailItem.SenderName
        '.Cells(xNextEmptyRow, 3) = xMailItem.SenderEmailAddress
        .Cells(xNextEmptyRow, 4) = xMailItem.Subject
        .Cells(xNextEmptyRow, 5) = xMailItem.ReceivedTime
        '.Cells(xNextEmptyRow, 6) = xMailItem.Body
    End With
    'xWs.Columns("A:E").AutoFit
    xWb.Save
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


