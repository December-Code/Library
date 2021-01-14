Attribute VB_Name = "Module1"
Sub transfer_Click()
Dim i, rowNum, ColNum As Long
Dim context As String

Dim Name, Th_CName, Th_EName, Adv_Name, Tp, Year, Abstr, Schname, Subject As Range
Dim Lan, KeyC, KeyE, Link As Range
    
ColNum = Worksheets("records").Cells(1, Columns.Count).End(xlToLeft).Column
    
Set Name = Worksheets("records").Range("A1:BA1").Find(What:="作者姓名(中文)", LookAt:=xlWhole)
Set Th_CName = Worksheets("records").Range("A1:BA1").Find(What:="論文名稱(中文)", LookAt:=xlWhole)
Set Th_EName = Worksheets("records").Range("A1:BA1").Find(What:="論文名稱(外文)", LookAt:=xlWhole)
Set Adv_Name = Worksheets("records").Range("A1:BA1").Find(What:="指導教授姓名(中文)", LookAt:=xlWhole)
Set Tp = Worksheets("records").Range("A1:BA1").Find(What:="學位類別", LookAt:=xlWhole)
Set Year = Worksheets("records").Range("A1:BA1").Find(What:="畢業學年度", LookAt:=xlWhole)
Set Abstr = Worksheets("records").Range("A1:BA1").Find(What:="摘要(中文)", LookAt:=xlWhole)
Set Lan = Worksheets("records").Range("A1:BA1").Find(What:="語文別", LookAt:=xlWhole)
Set KeyC = Worksheets("records").Range("A1:BA1").Find(What:="中文關鍵詞", LookAt:=xlWhole)
Set KeyE = Worksheets("records").Range("A1:BA1").Find(What:="外文關鍵詞", LookAt:=xlWhole)
Set Link = Worksheets("records").Range("A1:BA1").Find(What:="引用網址", LookAt:=xlWhole)
Set Schname = Worksheets("records").Range("A1:BA1").Find(What:="校院名稱", LookAt:=xlWhole)
Set Subject = Worksheets("records").Range("A1:BA1").Find(What:="系所名稱", LookAt:=xlWhole)
    
rowNum = Worksheets("records").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To rowNum
    Worksheets("transfer").Cells(i, "A") = Worksheets("records").Cells(i, Th_CName.Column) '論文名稱
    Worksheets("transfer").Cells(i, "B") = Worksheets("records").Cells(i, "F") '論文名稱(外文)
    Worksheets("transfer").Cells(i, "C") = Worksheets("records").Cells(i, Name.Column) '研究生
    Worksheets("transfer").Cells(i, "D") = 1911 + Worksheets("records").Cells(i, "S") 'Date
    Worksheets("transfer").Cells(i, "E") = Worksheets("transfer").Cells(1, "R") 'type
    Worksheets("transfer").Cells(i, "F") = Worksheets("records").Cells(i, "Q") '論文摘要
    'Worksheets("transfer").Cells(i, "G") = Worksheets("records").Cells(i, "U") '中文關鍵字
    Worksheets("transfer").Cells(i, "G") = Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "U"), Chr(10), ";") '中文關鍵字
    Worksheets("transfer").Cells(i, "H") = Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "V"), Chr(10), ";") '外文關鍵字
    
    If InStr(Worksheets("records").Cells(i, "P"), "中文") = 1 Then
    Worksheets("transfer").Cells(i, "I") = "zh-TW"
    ElseIf InStr(Worksheets("records").Cells(i, "P"), "英文") = 1 Then
    Worksheets("transfer").Cells(i, "I") = "en"
    End If
    
    'Worksheets("transfer").Cells(i, "I") = Worksheets("records").Cells(i, "P") 'language
    Worksheets("transfer").Cells(i, "J") = Worksheets("records").Cells(i, "AQ") '引用link
    Worksheets("transfer").Cells(i, "K") = Worksheets("records").Cells(i, "H") '校院名稱
    Worksheets("transfer").Cells(i, "L") = "學位:" + Worksheets("records").Cells(i, "K") '學位
    Worksheets("transfer").Cells(i, "M") = "指導教授:" + Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "L"), Chr(10), "、") '教授

Next i

MsgBox "轉換完畢!"

End Sub



Sub test()
    
    Dim i, rowNum, ColNum As Long
    Dim context As String
    Dim Name, Th_CName, Th_EName, Adv_Name, Tp, Year, Abstr, Schname, Subject As Range
    Dim Lan, KeyC, KeyE, Link As Range
    
    ColNum = Worksheets("records").Cells(1, Columns.Count).End(xlToLeft).Column
    MsgBox ColNum
    
    Set Name = Worksheets("records").Range("A1:BA1").Find(What:="作者姓名(中文)", LookAt:=xlWhole)
    Set Th_CName = Worksheets("records").Range("A1:BA1").Find(What:="論文名稱(中文)", LookAt:=xlWhole)
    Set Th_EName = Worksheets("records").Range("A1:BA1").Find(What:="論文名稱(外文)", LookAt:=xlWhole)
    Set Adv_Name = Worksheets("records").Range("A1:BA1").Find(What:="指導教授姓名(中文)", LookAt:=xlWhole)
    Set Tp = Worksheets("records").Range("A1:BA1").Find(What:="學位類別", LookAt:=xlWhole)
    Set Year = Worksheets("records").Range("A1:BA1").Find(What:="畢業學年度", LookAt:=xlWhole)
    Set Abstr = Worksheets("records").Range("A1:BA1").Find(What:="摘要(中文)", LookAt:=xlWhole)
    Set Lan = Worksheets("records").Range("A1:BA1").Find(What:="語文別", LookAt:=xlWhole)
    Set KeyC = Worksheets("records").Range("A1:BA1").Find(What:="中文關鍵詞", LookAt:=xlWhole)
    Set KeyE = Worksheets("records").Range("A1:BA1").Find(What:="外文關鍵詞", LookAt:=xlWhole)
    Set Link = Worksheets("records").Range("A1:BA1").Find(What:="引用網址", LookAt:=xlWhole)
    Set Schname = Worksheets("records").Range("A1:BA1").Find(What:="校院名稱", LookAt:=xlWhole)
    Set Subject = Worksheets("records").Range("A1:BA1").Find(What:="系所名稱", LookAt:=xlWhole)
    
    MsgBox "作者姓名" + Str(Name.Column)
    MsgBox "論文名稱(中文)" + Str(Th_CName.Column)
    MsgBox "論文名稱(外文)" + Str(Th_EName.Column)
    MsgBox "指導教授姓名(中文)" + Str(Adv_Name.Column)
    MsgBox "學位類別" + Str(Tp.Column)
    MsgBox "畢業學年度" + Str(Year.Column)
    MsgBox "摘要(中文)" + Str(Abstr.Column)
    MsgBox "語文別" + Str(Lan.Column)
    MsgBox "中文關鍵詞" + Str(KeyC.Column)
    MsgBox "外文關鍵詞" + Str(KeyE.Column)
    MsgBox "引用網址" + Str(Link.Column)
    MsgBox "校院名稱" + Str(Schname.Column)
    MsgBox "系所名稱" + Str(Subject.Column)
    
End Sub
