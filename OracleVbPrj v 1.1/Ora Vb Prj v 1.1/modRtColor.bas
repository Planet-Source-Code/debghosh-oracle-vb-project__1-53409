Attribute VB_Name = "modRtColor"
'I get this code from Planet Source Code.
Option Explicit
Public Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Public Const EM_GETLINE As Long = &HC4
Public Const EM_GETLINECOUNT As Long = &HBA
Public Const EM_LINEINDEX As Long = &HBB
Dim l As Long
Dim c As Long
Dim startline, nowline
Type WORD_TYPE
    Text As String  ' word to be colored
    Color As Long   ' color to color the word :)
End Type

Dim Words() As WORD_TYPE
Public Sub RtLineCol(rt As RichTextBox, lbl As Label)
    l = SendMessage(rt.hwnd, EM_GETLINECOUNT, 0&, 0&)
    l = 1 + rt.GetLineFromChar(rt.SelStart)
    c = SendMessage(rt.hwnd, EM_LINEINDEX, ByVal l - 1, 0&)
    c = (rt.SelStart) - c
    lbl.Caption = "" & l & "," & c & " "
End Sub

Public Sub InitWords()
    ReDim Words(0 To 56)
    Words(0).Text = "Create"
    Words(0).Color = vbBlue
    Words(1).Text = "CREATE"
    Words(1).Color = vbBlue
    Words(2).Text = "create"
    Words(2).Color = vbBlue
    Words(3).Text = "Select"
    Words(3).Color = vbBlue
    Words(4).Text = "select"
    Words(4).Color = vbBlue
    Words(5).Text = "SELECT"
    Words(5).Color = vbBlue
    Words(6).Text = "TABLE"
    Words(6).Color = vbBlue
    Words(7).Text = "table"
    Words(7).Color = vbBlue
    Words(8).Text = "Table"
    Words(8).Color = vbBlue
    Words(9).Text = "VIEW"
    Words(9).Color = vbBlue
    Words(10).Text = "view"
    Words(10).Color = vbBlue
    Words(11).Text = "View"
    Words(11).Color = vbBlue
    Words(12).Text = "PROCEDURE"
    Words(12).Color = vbBlue
    Words(13).Text = "procedure"
    Words(13).Color = vbBlue
    Words(14).Text = "Procedure"
    Words(14).Color = vbBlue
    Words(15).Text = "SYNONYM"
    Words(15).Color = vbBlue
    Words(16).Text = "synonym"
    Words(16).Color = vbBlue
    Words(17).Text = "Synonym"
    Words(17).Color = vbBlue
    Words(18).Text = "Sequence"
    Words(18).Color = vbBlue
    Words(19).Text = "SEQUENCE"
    Words(19).Color = vbBlue
    Words(20).Text = "sequence"
    Words(20).Color = vbBlue
    Words(21).Text = "Index"
    Words(21).Color = vbBlue
    Words(22).Text = "INDEX"
    Words(22).Color = vbBlue
    Words(23).Text = "index"
    Words(23).Color = vbBlue
    Words(24).Text = "Package"
    Words(24).Color = vbBlue
    Words(25).Text = "PACKAGE"
    Words(25).Color = vbBlue
    Words(26).Text = "package"
    Words(26).Color = vbBlue
    Words(27).Text = "Type"
    Words(27).Color = vbBlue
    Words(28).Text = "TYPE"
    Words(28).Color = vbBlue
    Words(29).Text = "type"
    Words(29).Color = vbBlue
    Words(30).Text = "Package Body"
    Words(30).Color = vbBlue
    Words(31).Text = "PACKAGE BODY"
    Words(31).Color = vbBlue
    Words(32).Text = "package body"
    Words(32).Color = vbBlue
    Words(33).Text = "FROM"
    Words(33).Color = vbBlue
    Words(34).Text = "from"
    Words(34).Color = vbBlue
    Words(35).Text = "From"
    Words(35).Color = vbBlue
    Words(36).Text = "Where"
    Words(36).Color = vbBlue
    Words(37).Text = "WHERE"
    Words(37).Color = vbBlue
    Words(38).Text = "where"
    Words(38).Color = vbBlue
    Words(39).Text = "AS"
    Words(39).Color = vbBlue
    Words(40).Text = "As"
    Words(40).Color = vbBlue
    Words(41).Text = "as"
    Words(41).Color = vbBlue
    Words(42).Text = "If"
    Words(42).Color = vbBlue
    Words(43).Text = "IF"
    Words(43).Color = vbBlue
    Words(44).Text = "if"
    Words(44).Color = vbBlue
    Words(45).Text = "Order By"
    Words(45).Color = vbBlue
    Words(46).Text = "ORDER BY"
    Words(46).Color = vbBlue
    Words(47).Text = "order by"
    Words(47).Color = vbBlue
    Words(48).Text = "Group By"
    Words(48).Color = vbBlue
    Words(49).Text = "group by"
    Words(49).Color = vbBlue
    Words(50).Text = "GROUP BY"
    Words(50).Color = vbBlue
    Words(51).Text = "HAVING"
    Words(51).Color = vbBlue
    Words(52).Text = "having"
    Words(52).Color = vbBlue
    Words(53).Text = "Having"
    Words(53).Color = vbBlue
    Words(54).Text = "Between"
    Words(54).Color = vbBlue
    Words(55).Text = "BETWEEN"
    Words(55).Color = vbBlue
    Words(56).Text = "between"
    Words(56).Color = vbBlue
End Sub
Public Sub DoColor(rtB As RichTextBox)
Dim i As Long
Dim p1 As Long, p2 As Long
Dim Text As String
Dim sTmp As String
    'On Error Resume Next
    ' cache the text - speeds things up a bit
    Text = rtB.Text
    ' go through each item in the Words array
    For i = LBound(Words) To UBound(Words)
        ' find each instance of the word in the rtb
        p1 = InStr(1, Text, Words(i).Text)
        Do While p1 > 0
            ' color it to the appropriate color
            rtB.SelStart = p1 - 1
            rtB.SelLength = Len(Words(i).Text)
            rtB.SelColor = Words(i).Color
            ' go on to the next word
            p1 = InStr(p1 + 1, Text, Words(i).Text)
        Loop
    Next i
    ' go through and color all the comment lines
    p1 = 1
    Do While p1 <> 2 And p1 < Len(Text)
        ' find the next eol character
        p2 = InStr(p1 + 1, Text, vbCrLf)
        If p2 = 0 Then p2 = Len(Text)
        ' grab this line into a temp variable
        sTmp = Mid$(Text, p1, p2 - p1)
        ' if it's a comment line - color it
        If Left(Trim$(sTmp), 1) = "'" Then
            rtB.SelStart = p1
            rtB.SelLength = p2 - p1
            rtB.SelColor = vbRed
        End If
        ' move onto the next line
        p1 = p2 + 2
    Loop
End Sub

