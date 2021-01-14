Attribute VB_Name = "Module1"
Sub transfer_Click()
Dim i, rowNum, ColNum As Long
Dim context As String

Dim Name, Th_CName, Th_EName, Adv_Name, Tp, Year, Abstr, Schname, Subject As Range
Dim Lan, KeyC, KeyE, Link As Range
    
ColNum = Worksheets("records").Cells(1, Columns.Count).End(xlToLeft).Column
    
Set Name = Worksheets("records").Range("A1:BA1").Find(What:="�@�̩m�W(����)", LookAt:=xlWhole)
Set Th_CName = Worksheets("records").Range("A1:BA1").Find(What:="�פ�W��(����)", LookAt:=xlWhole)
Set Th_EName = Worksheets("records").Range("A1:BA1").Find(What:="�פ�W��(�~��)", LookAt:=xlWhole)
Set Adv_Name = Worksheets("records").Range("A1:BA1").Find(What:="���ɱб©m�W(����)", LookAt:=xlWhole)
Set Tp = Worksheets("records").Range("A1:BA1").Find(What:="�Ǧ����O", LookAt:=xlWhole)
Set Year = Worksheets("records").Range("A1:BA1").Find(What:="���~�Ǧ~��", LookAt:=xlWhole)
Set Abstr = Worksheets("records").Range("A1:BA1").Find(What:="�K�n(����)", LookAt:=xlWhole)
Set Lan = Worksheets("records").Range("A1:BA1").Find(What:="�y��O", LookAt:=xlWhole)
Set KeyC = Worksheets("records").Range("A1:BA1").Find(What:="���������", LookAt:=xlWhole)
Set KeyE = Worksheets("records").Range("A1:BA1").Find(What:="�~�������", LookAt:=xlWhole)
Set Link = Worksheets("records").Range("A1:BA1").Find(What:="�ޥκ��}", LookAt:=xlWhole)
Set Schname = Worksheets("records").Range("A1:BA1").Find(What:="�հ|�W��", LookAt:=xlWhole)
Set Subject = Worksheets("records").Range("A1:BA1").Find(What:="�t�ҦW��", LookAt:=xlWhole)
    
rowNum = Worksheets("records").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To rowNum
    Worksheets("transfer").Cells(i, "A") = Worksheets("records").Cells(i, Th_CName.Column) '�פ�W��
    Worksheets("transfer").Cells(i, "B") = Worksheets("records").Cells(i, "F") '�פ�W��(�~��)
    Worksheets("transfer").Cells(i, "C") = Worksheets("records").Cells(i, Name.Column) '��s��
    Worksheets("transfer").Cells(i, "D") = 1911 + Worksheets("records").Cells(i, "S") 'Date
    Worksheets("transfer").Cells(i, "E") = Worksheets("transfer").Cells(1, "R") 'type
    Worksheets("transfer").Cells(i, "F") = Worksheets("records").Cells(i, "Q") '�פ�K�n
    'Worksheets("transfer").Cells(i, "G") = Worksheets("records").Cells(i, "U") '��������r
    Worksheets("transfer").Cells(i, "G") = Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "U"), Chr(10), ";") '��������r
    Worksheets("transfer").Cells(i, "H") = Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "V"), Chr(10), ";") '�~������r
    
    If InStr(Worksheets("records").Cells(i, "P"), "����") = 1 Then
    Worksheets("transfer").Cells(i, "I") = "zh-TW"
    ElseIf InStr(Worksheets("records").Cells(i, "P"), "�^��") = 1 Then
    Worksheets("transfer").Cells(i, "I") = "en"
    End If
    
    'Worksheets("transfer").Cells(i, "I") = Worksheets("records").Cells(i, "P") 'language
    Worksheets("transfer").Cells(i, "J") = Worksheets("records").Cells(i, "AQ") '�ޥ�link
    Worksheets("transfer").Cells(i, "K") = Worksheets("records").Cells(i, "H") '�հ|�W��
    Worksheets("transfer").Cells(i, "L") = "�Ǧ�:" + Worksheets("records").Cells(i, "K") '�Ǧ�
    Worksheets("transfer").Cells(i, "M") = "���ɱб�:" + Application.WorksheetFunction.Substitute(Worksheets("records").Cells(i, "L"), Chr(10), "�B") '�б�

Next i

MsgBox "�ഫ����!"

End Sub



Sub test()
    
    Dim i, rowNum, ColNum As Long
    Dim context As String
    Dim Name, Th_CName, Th_EName, Adv_Name, Tp, Year, Abstr, Schname, Subject As Range
    Dim Lan, KeyC, KeyE, Link As Range
    
    ColNum = Worksheets("records").Cells(1, Columns.Count).End(xlToLeft).Column
    MsgBox ColNum
    
    Set Name = Worksheets("records").Range("A1:BA1").Find(What:="�@�̩m�W(����)", LookAt:=xlWhole)
    Set Th_CName = Worksheets("records").Range("A1:BA1").Find(What:="�פ�W��(����)", LookAt:=xlWhole)
    Set Th_EName = Worksheets("records").Range("A1:BA1").Find(What:="�פ�W��(�~��)", LookAt:=xlWhole)
    Set Adv_Name = Worksheets("records").Range("A1:BA1").Find(What:="���ɱб©m�W(����)", LookAt:=xlWhole)
    Set Tp = Worksheets("records").Range("A1:BA1").Find(What:="�Ǧ����O", LookAt:=xlWhole)
    Set Year = Worksheets("records").Range("A1:BA1").Find(What:="���~�Ǧ~��", LookAt:=xlWhole)
    Set Abstr = Worksheets("records").Range("A1:BA1").Find(What:="�K�n(����)", LookAt:=xlWhole)
    Set Lan = Worksheets("records").Range("A1:BA1").Find(What:="�y��O", LookAt:=xlWhole)
    Set KeyC = Worksheets("records").Range("A1:BA1").Find(What:="���������", LookAt:=xlWhole)
    Set KeyE = Worksheets("records").Range("A1:BA1").Find(What:="�~�������", LookAt:=xlWhole)
    Set Link = Worksheets("records").Range("A1:BA1").Find(What:="�ޥκ��}", LookAt:=xlWhole)
    Set Schname = Worksheets("records").Range("A1:BA1").Find(What:="�հ|�W��", LookAt:=xlWhole)
    Set Subject = Worksheets("records").Range("A1:BA1").Find(What:="�t�ҦW��", LookAt:=xlWhole)
    
    MsgBox "�@�̩m�W" + Str(Name.Column)
    MsgBox "�פ�W��(����)" + Str(Th_CName.Column)
    MsgBox "�פ�W��(�~��)" + Str(Th_EName.Column)
    MsgBox "���ɱб©m�W(����)" + Str(Adv_Name.Column)
    MsgBox "�Ǧ����O" + Str(Tp.Column)
    MsgBox "���~�Ǧ~��" + Str(Year.Column)
    MsgBox "�K�n(����)" + Str(Abstr.Column)
    MsgBox "�y��O" + Str(Lan.Column)
    MsgBox "���������" + Str(KeyC.Column)
    MsgBox "�~�������" + Str(KeyE.Column)
    MsgBox "�ޥκ��}" + Str(Link.Column)
    MsgBox "�հ|�W��" + Str(Schname.Column)
    MsgBox "�t�ҦW��" + Str(Subject.Column)
    
End Sub
