VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents objMails As Outlook.Items
Attribute objMails.VB_VarHelpID = -1

Private Sub Application_Startup()
    Set objMails = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub objMails_ItemAdd(ByVal Item As Object)
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

    If Item.Class = olMail Then
       Set objMail = Item
    End If
 
    'Specify the Excel file which you want to auto export the email list
    'You can change it as per your case
    strExcelFile = "D:\Inspur Files\ÓÊ¼þ¼ÇÂ¼.xlsx"
 
    'Get Access to the Excel file
    On Error Resume Next
    Set objExcelApp = GetObject(, "Excel.Application")
    If Error <> 0 Then
       Set objExcelApp = CreateObject("Excel.Application")
    End If
    Set objExcelWorkBook = objExcelApp.Workbooks.Open(strExcelFile)
    Set objExcelWorkSheet = objExcelWorkBook.Sheets("Sheet1")
 
    'Get the next empty row in the Excel worksheet
    nNextEmptyRow = objExcelWorkSheet.Range("A" & objExcelWorkSheet.Rows.Count).End(xlUp).Row + 1
 
    'Specify the corresponding values in the different columns
    strColumnB = objMail.SenderName
    strColumnC = objMail.SenderEmailAddress
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
End Sub



    


