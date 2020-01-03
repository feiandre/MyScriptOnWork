Option Explicit
Public Sub 正文插入文本()
    Dim Inspector As Outlook.Inspector
    Dim wdDoc As Word.Document
    Dim Selection As Word.Selection


    Set Inspector = Application.ActiveInspector()
    Set wdDoc = Inspector.WordEditor
    Set Selection = wdDoc.Application.Selection
        'Selection.InsertBefore Format(Now, "DD/MM/YYYY")
         Selection.InsertBefore "情况说明如下。已通过一级评审，请您评审。"
              
    Set Inspector = Nothing
    Set wdDoc = Nothing
    Set Selection = Nothing
End Sub

Public Sub 格式化本文()
    Dim Inspector As Outlook.Inspector
    Dim wdDoc As Word.Document
    Dim Selection As Word.Selection


    Set Inspector = Application.ActiveInspector()
    Set wdDoc = Inspector.WordEditor
    Set Selection = wdDoc.Application.Selection
        
       '选中的文本格式化 微软雅黑 小四  单倍行距
       Selection.Font.Name = "微软雅黑"
       Selection.Font.Size = 12
       'Selection.Font.Color = RGB(0, 0, 0)
       Selection.Range.Font.Bold = False
       Selection.Font.Color = wdColorBlack
    
    With Selection.ParagraphFormat
                '.RightIndent = InchesToPoints(1)
                '.RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaselineAlignment = wdBaselineAlignAuto
    End With
             
    Set Inspector = Nothing
    Set wdDoc = Nothing
    Set Selection = Nothing
End Sub

Sub regdemo()

s = "Olivia Shang (尚丹平)"
Set re = CreateObject("vbscript.regexp")
re.Pattern = "\(([^)]+)"
Set ms = re.Execute(s)
MsgBox ms(0).submatches(0)

're.Pattern = "\((.*?)\)"
'Set ms = re.Execute(s)
'MsgBox ms(0).submatches(0)
End Sub


Sub regdemo2()
s = "合同协议评审--ICA19102501&ICC19102501宜宾政务云（宜宾新力拓）推动客户业务上云合作协议及外包合同评审"

Set re = CreateObject("vbscript.regexp")
re.Pattern = "(ICC\d{8})"
Set ms = re.Execute(s)
MsgBox ms(0).submatches(0)

re.Pattern = "(ICA\d{8})"
Set ms = re.Execute(s)
MsgBox ms(0).submatches(0)


End Sub

