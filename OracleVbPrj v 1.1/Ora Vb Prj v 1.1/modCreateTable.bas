Attribute VB_Name = "modCreateTable"
Option Explicit
Public LoadValidate As Boolean
Dim tRs As New ADODB.Recordset
Dim cRs As New ADODB.Recordset
Public Sub OnFormLoad()
    Dim i, j
    With frmCreateTable
        .txtTableName.Text = ""
        .txtTableName.TabIndex = 0
        .cmbColNo.TabIndex = 1
        .cmdLoad.TabIndex = 2
        .Fg.TabIndex = 3
        .rtCond.TabIndex = 4
        .dLblShowSQL.TabIndex = 5
        .dLblHelp.TabIndex = 6
        .Fg.ColWidth(0) = 300
        .Fg.RowHeight(0) = 300
        For j = 1 To 20
        .cmbColNo.AddItem j
        Next j
    .cmbColNo.ListIndex = 4
    End With
    LoadValidate = False
End Sub
Public Sub OnLoadClick()
Dim i, j
    With frmCreateTable
        If .txtTableName.Text = "" Then
            gDmsg.DebMsgBox "Please Input Table Name First Then Select No Of Columns", .Caption, DebmsgInformation
            .txtTableName.SetFocus
            Exit Sub
        End If
        If .cmbColNo < 1 Then
            gDmsg.DebMsgBox "Table Must Contain One Column", .Caption, DebmsgInformation
            .cmbColNo.SetFocus
            Exit Sub
        End If
        .Fg.Cols = 17
        .Fg.Rows = Val(.cmbColNo.Text) + 1
        .Fg.ColWidth(1) = 3000
        .Fg.TextMatrix(0, 1) = "Column Name  (1)"
        .Fg.ColWidth(2) = 2000
        .Fg.TextMatrix(0, 2) = "Datatype  (2)"
        .Fg.ColWidth(3) = 1000
        .Fg.TextMatrix(0, 3) = "Datalength  (3)"
        .Fg.ColWidth(4) = 1000
        .Fg.TextMatrix(0, 4) = "Scale  (4)"
        .Fg.ColWidth(5) = 3000
        .Fg.TextMatrix(0, 5) = "Constraint  (5)"
        For j = 6 To 16
            .Fg.ColWidth(j) = 2500
            .Fg.TextMatrix(0, j) = "Opt  ( " & j & ")"
        Next j
        For i = 1 To Val(.cmbColNo.Text)
            .Fg.TextMatrix(i, 0) = i
        Next i
    End With
    LoadValidate = True
End Sub
Sub MSHFlexGridEdit(MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
   ' Use the character that was typed.
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 32
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   ' Anything else means replace the current text.
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
   End Select
   ' Show Edt at the right place.
   Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8
   Edt.Visible = True
   ' And make it work.
   Edt.SetFocus
End Sub
Sub EditKeyCode(MSHFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
   ' Standard edit control processing.
   Select Case KeyCode
   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSHFlexGrid.SetFocus
   Case 13   ' ENTER return focus to MSHFlexGrid.
      MSHFlexGrid.SetFocus
   Case 38      ' Up.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.Row > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.Row = MSHFlexGrid.Row - 1
      End If
   Case 40      ' Down.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.Row < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.Row = MSHFlexGrid.Row + 1
      End If
   End Select
End Sub
Public Sub ColData(dt As Integer) ' Column DatatType And Size Of The Datatype
Dim i
With frmCreateTable
    Select Case dt
        Case 2 ' Here 2 Represents The Column No.
            .cmbCol.Clear
            With .cmbCol
                .AddItem "CHAR"
                .AddItem "DATE"
                .AddItem "FLOAT"
                .AddItem "LONG"
                .AddItem "LONG RAW"
                .AddItem "NUMBER"
                .AddItem "RAW"
                .AddItem "VARCHAR2"
            End With
        .OFrmCtl1.StatusCaption = "Select Datatype"
        
        Case 3
            .cmbCol.Clear
            With .cmbCol
                If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "CHAR" Then
                    For i = 1 To 2000
                        .AddItem i
                    Next i
                Else
                    If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "DATE" Or _
                             frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "FLOAT" Or _
                             frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "LONG" Or _
                             frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "RAW" Or _
                             frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "LONG RAW" Then
                                .Clear
                                .AddItem " "
                    Else
                        If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "NUMBER" Then
                            For i = 1 To 38
                                .AddItem i
                            Next i
                        Else
                            If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "VARCHAR2" Then
                                For i = 1 To 4000
                                    .AddItem i
                                Next i
                            End If
                        End If
                    End If
                End If
            End With
        .OFrmCtl1.StatusCaption = "Select Size"
        
        Case 4
            .cmbCol.Clear
                With .cmbCol
                    If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "NUMBER" Then
                        If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 3) <> "" Then
                            For i = -8 To 127
                                .AddItem i
                            Next i
                        Else
                            gDmsg.DebMsgBox "SELECT PRECISION (" & frmCreateTable.Fg.Row & ",3)", frmCreateTable.Caption, DebmsgInformation
                        End If
                    End If
                End With
       .OFrmCtl1.StatusCaption = "Select Precision"
       
       Case 5
        .cmbCol.Clear
        With .cmbCol
            If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 1) = "" Or frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "" Then
                gDmsg.DebMsgBox "ENTER COLUMN NAME", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            Else
            .AddItem "CHECK"
            .AddItem "CONSTRAINT"
            .AddItem "DEFAULT"
            .AddItem "NULL"
            .AddItem "NOT NULL"
            .AddItem "PRIMARY KEY"
            .AddItem "REFERENCES"
            .AddItem "UNIQUE"
            End If
        End With
      .OFrmCtl1.StatusCaption = "Select Constraint"
      
      Case 6 To .Fg.Cols - 1
        Call ColConstr
    End Select
End With
End Sub
Public Sub ColConstr()
With frmCreateTable
   If .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) = "CHECK" Then
    .cmbCol.Clear
    .OFrmCtl1.StatusCaption = "ENTER  " & .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) & " CONS. IDENTIFIER AT (" & .Fg.Row & "," & .Fg.Col & ")"
   Else
    If .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) = "CONSTRAINT" Then
        .cmbCol.Clear
        .OFrmCtl1.StatusCaption = "ENTER  " & .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) & " CONS. IDENTIFIER AT (" & .Fg.Row & "," & .Fg.Col & ")"
    Else
        If .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) = "DEFAULT" Then
            .cmbCol.Clear
            .OFrmCtl1.StatusCaption = "ENTER  " & .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) & " CONS. IDENTIFIER AT (" & .Fg.Row & "," & .Fg.Col & ")"
            With .cmbCol
               .AddItem "NULL"
               .AddItem "NOT NULL"
                    If frmCreateTable.Fg.TextMatrix(frmCreateTable.Fg.Row, 2) = "DATE" Then
                        .AddItem "SYSDATE"
                    End If
            End With
        Else
            If frmCreateTable.Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) = "REFERENCES" Then
                .cmbCol.Clear
                .OFrmCtl1.StatusCaption = "ENTER  " & .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) & " CONS. IDENTIFIER AT (" & .Fg.Row & "," & .Fg.Col & ")"
                '.cmbCol.AddItem "ENTER REFERENCIAL KEY"
                Call ReferencialKey
            Else
                If frmCreateTable.Fg.TextMatrix(.Fg.Row, .Fg.Col - 2) = "CONSTRAINT" Then
                    .cmbCol.Clear
                    With .cmbCol
                        .AddItem "NULL"
                        .AddItem "NOT NULL"
                        .AddItem "PRIMARY KEY"
                        .AddItem "REFERENCES"
                        .AddItem "UNIQUE"
                    End With
                    .OFrmCtl1.StatusCaption = "ENTER  " & .Fg.TextMatrix(.Fg.Row, .Fg.Col - 1) & " CONS. IDENTIFIER AT (" & .Fg.Row & "," & .Fg.Col & ")"
                Else
                .cmbCol.Clear
                End If
            End If
        End If
    End If
End If
End With
End Sub
Public Sub CreateTable()
Dim ch, df, nl, nln, pc, re, un, i, r, c As Integer
Dim cnt, pr As Integer
cnt = 0
pr = 0
With frmCreateTable

For r = 1 To .Fg.Rows - 1
    If .Fg.TextMatrix(r, 1) = "" Then
        gDmsg.DebMsgBox "ENTER COLUMN NAME (" & r & ",1)", frmCreateTable.Caption, DebmsgInformation
        Exit Sub
    Else
        If .Fg.TextMatrix(r, 2) = "" Then
            gDmsg.DebMsgBox "SELECT DATATYPE (" & r & ",2)", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
    End If
    If .Fg.TextMatrix(r, 2) = "CHAR" Or .Fg.TextMatrix(r, 2) = "VARCHAR2" Then
        If .Fg.TextMatrix(r, 3) = "" Then
           gDmsg.DebMsgBox "SELECT SIZE (" & r & ",3)", frmCreateTable.Caption, DebmsgInformation
           Exit Sub
        End If
    End If
    If .Fg.TextMatrix(r, 2) = "FLOAT" Or .Fg.TextMatrix(r, 2) = "NUMBER" Then
        If .Fg.TextMatrix(r, 4) <> "" Then
            If .Fg.TextMatrix(r, 3) = "" Then
                gDmsg.DebMsgBox "SELECT PRECISION (" & r & ",3)", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
        End If
    End If
    If .Fg.TextMatrix(r, 2) = "RAW" Then
        If .Fg.TextMatrix(r, 3) = "" Then
            gDmsg.DebMsgBox "ENTER RAW SIZE ( " & r & ",3)", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
    End If
    For c = 5 To .Fg.Cols - 1
        If .Fg.TextMatrix(r, c) = "CHECK" Then
        If (c <= .Fg.Cols - 2) Then
            If .Fg.TextMatrix(r, c + 1) = "" Then
                gDmsg.DebMsgBox "ENTER CHECK IDENTIFIER AT (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
        Else
            gDmsg.DebMsgBox "SELECT CHECK CONSTRAINT AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
        End If
        If .Fg.TextMatrix(r, c) = "CONSTRAINT" Then
        If (c <= .Fg.Cols - 3) Then
            If .Fg.TextMatrix(r, c + 1) = "" Then
                gDmsg.DebMsgBox "ENTER CONSTRAINT IDENTIFIER (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
            If .Fg.TextMatrix(r, c + 2) = "" Then
                gDmsg.DebMsgBox "ENTER CONSTRAINT IDENTIFIER (" & r & "," & c + 2 & ")", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
        Else
            gDmsg.DebMsgBox "SELECT CONSTRAINT AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
        End If
        If .Fg.TextMatrix(r, c) = "DEFAULT" Then
        If (c <= .Fg.Cols - 2) Then
            If .Fg.TextMatrix(r, c + 1) = "" Then
                gDmsg.DebMsgBox "ENTER DEFAULT TYPE (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
        Else
            gDmsg.DebMsgBox "SELECT DEFAULT IDENTIFIER AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
        End If
        If .Fg.TextMatrix(r, c) = "REFERENCES" Then
        If (c <= .Fg.Cols - 2) Then
            If .Fg.TextMatrix(r, c + 1) = "" Then
                gDmsg.DebMsgBox "ENTER REFERENCIAL KEY (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            End If
        Else
            gDmsg.DebMsgBox "SELECT REFERENCIAL KEY AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
            Exit Sub
        End If
        End If
    Next c
Next r
End With
End Sub
Public Sub ReferencialKey()
On Error GoTo rsError
    Screen.MousePointer = vbHourglass
    Set tRs = New ADODB.Recordset
    Set tRs = db.OpenSchema(adSchemaTables, Array(Empty, "" & uid & "", Empty, "Table"))
        Do Until tRs.EOF
            Set cRs = db.OpenSchema(adSchemaColumns, Array(Empty, "" & uid & "", "" & tRs!TABLE_NAME & ""))
            Do Until cRs.EOF
                frmCreateTable.cmbCol.AddItem "" & uid & "." & tRs!TABLE_NAME & "(" & cRs!COLUMN_NAME & ")"
                cRs.MoveNext
            Loop
            tRs.MoveNext
        Loop
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error : " & Err.Description & "", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Public Sub TextSQL()
    Dim i, c, x, a
    Dim s As String
    Dim r
    Dim tx As String
    Dim ch, df, nl, nln, pc, re, un  As Integer
    Dim cnt, pr As Integer
    cnt = 0
    pr = 0

    frmCreateTable.txtFgCell.Text = ""
    'Call CreateTable
    With frmCreateTable
        
        'Check For Same Column Name
        For i = 1 To .Fg.Rows - 1
            s = Trim$(UCase(.Fg.TextMatrix(i, 1)))
            x = i
            If x <= .Fg.Rows - 2 Then
                For c = (i + 1) To .Fg.Rows - 1
                    If s = Trim$(UCase(.Fg.TextMatrix(c, 1))) Then
                        gDmsg.DebMsgBox "Column Name Should Not Be Same : " & i & " No. Column And " & c & " No. Column And Column Name Is : " & s & "", "Error", DebmsgOkOnly
                        Exit Sub
                    End If
                Next c
            End If
        Next i
        
        For r = 1 To .Fg.Rows - 1
            If .Fg.TextMatrix(r, 1) = "" Then
                gDmsg.DebMsgBox "ENTER COLUMN NAME (" & r & ",1)", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
            Else
                If .Fg.TextMatrix(r, 2) = "" Then
                    gDmsg.DebMsgBox "SELECT DATATYPE (" & r & ",2)", frmCreateTable.Caption, DebmsgInformation
                    Exit Sub
                End If
            End If
            
            If .Fg.TextMatrix(r, 2) = "CHAR" Or .Fg.TextMatrix(r, 2) = "VARCHAR2" Then
                If .Fg.TextMatrix(r, 3) = "" Then
                    gDmsg.DebMsgBox "SELECT SIZE (" & r & ",3)", frmCreateTable.Caption, DebmsgInformation
                Exit Sub
                End If
            End If
            
            If .Fg.TextMatrix(r, 2) = "FLOAT" Or .Fg.TextMatrix(r, 2) = "NUMBER" Then
                If .Fg.TextMatrix(r, 4) <> "" Then
                    If .Fg.TextMatrix(r, 3) = "" Then
                        gDmsg.DebMsgBox "SELECT PRECISION (" & r & ",3)", frmCreateTable.Caption, DebmsgInformation
                        Exit Sub
                    End If
                End If
            End If
    
            If .Fg.TextMatrix(r, 2) = "RAW" Then
                If .Fg.TextMatrix(r, 3) = "" Then
                    gDmsg.DebMsgBox "ENTER RAW SIZE ( " & r & ",3)", frmCreateTable.Caption, DebmsgInformation
                    Exit Sub
                End If
            End If
    
            For c = 5 To .Fg.Cols - 1
                If .Fg.TextMatrix(r, c) = "CHECK" Then
                    If (c <= .Fg.Cols - 2) Then
                        If .Fg.TextMatrix(r, c + 1) = "" Then
                            gDmsg.DebMsgBox "ENTER CHECK IDENTIFIER AT (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                            Exit Sub
                        End If
                    Else
                        gDmsg.DebMsgBox "SELECT CHECK CONSTRAINT AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
                        Exit Sub
                    End If
                End If
                If .Fg.TextMatrix(r, c) = "CONSTRAINT" Then
                    If (c <= .Fg.Cols - 3) Then
                        If .Fg.TextMatrix(r, c + 1) = "" Then
                            gDmsg.DebMsgBox "ENTER CONSTRAINT NAME (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                            Exit Sub
                        End If
                        If .Fg.TextMatrix(r, c + 2) = "" Then
                            gDmsg.DebMsgBox "ENTER CONSTRAINT IDENTIFIER (" & r & "," & c + 2 & ")", frmCreateTable.Caption, DebmsgInformation
                            Exit Sub
                        End If
                    Else
                        gDmsg.DebMsgBox "SELECT CONSTRAINT AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
                        Exit Sub
                    End If
                End If
        
                If .Fg.TextMatrix(r, c) = "DEFAULT" Then
                    If (c <= .Fg.Cols - 2) Then
                        If .Fg.TextMatrix(r, c + 1) = "" Then
                            gDmsg.DebMsgBox "ENTER DEFAULT TYPE (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                            Exit Sub
                        End If
                    Else
                        gDmsg.DebMsgBox "SELECT DEFAULT CONSTRAINT AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
                        Exit Sub
                    End If
                End If
        
                If .Fg.TextMatrix(r, c) = "REFERENCES" Then
                    If (c <= .Fg.Cols - 2) Then
                        If .Fg.TextMatrix(r, c + 1) = "" Then
                            gDmsg.DebMsgBox "ENTER REFERENCIAL KEY (" & r & "," & c + 1 & ")", frmCreateTable.Caption, DebmsgInformation
                            Exit Sub
                        End If
                    Else
                        gDmsg.DebMsgBox "SELECT REFERENCIAL KEY AT RIGHT PLACE (" & r & "," & c & ")", frmCreateTable.Caption, DebmsgInformation
                        Exit Sub
                    End If
                End If
        Next c
    Next r
        
        'Check Datatype
        For r = 1 To .Fg.Rows - 1
        s = ""
        tx = ""
        Select Case Trim$(.Fg.TextMatrix(r, 2))
            
            Case "CHAR", "RAW", "VARCHAR2"
                tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & "( " & .Fg.TextMatrix(r, 3) & ") "
                    
            Case "FLOAT", "NUMBER"
                If .Fg.TextMatrix(r, 3) <> "" Then
                    If .Fg.TextMatrix(r, 4) <> "" Then
                        tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & "(" & .Fg.TextMatrix(r, 3) & "," & .Fg.TextMatrix(r, 4) & ")"
                            
                    Else
                        tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & "(" & .Fg.TextMatrix(r, 3) & ")"
                            
                    End If
                Else
                    tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & ""
                End If
                    
            Case "LONG", "LONG RAW", "DATE"
                tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & ""
                    
            Case Else
                If .Fg.TextMatrix(r, 3) <> "" Then
                    If .Fg.TextMatrix(r, 4) <> "" Then
                        tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & "(" & .Fg.TextMatrix(r, 3) & "," & .Fg.TextMatrix(r, 4) & ")"
                    Else
                        tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & "(" & .Fg.TextMatrix(r, 3) & ")"
                    End If
                Else
                    tx = " " & .Fg.TextMatrix(r, 1) & "  " & .Fg.TextMatrix(r, 2) & ""
                End If
                        
        End Select
    
            For a = 5 To .Fg.Cols - 1
                If .Fg.TextMatrix(r, a) <> "" Then
                    s = " " & s & " " & .Fg.TextMatrix(r, a) & " "
                End If
            Next a
        
            If r < .Fg.Rows - 1 Then
                .txtFgCell.SelText = " " & tx & "  " & s & "," & vbCrLf
                s = ""
                tx = ""
            Else
                .txtFgCell.SelText = " " & tx & "  " & s & " "
                s = ""
                tx = ""
            End If
            
        Next r
        
        .rtSQL.SelText = "CREATE TABLE " & uid & "." & Trim$(.txtTableName.Text) & " (" & vbCrLf _
                    & " " & .txtFgCell.Text & " " & vbCrLf _
                    & ") " & vbCrLf _
                    & " " & Trim$(.rtCond.Text) & " "
        'Dim rMs
        'rMs = gDmsg.DebMsgBox("Do You Want See SQL", "Conformation", DebmsgYesNo)
            'If rMs = DebmsgYes Then
                'frmCreateTblSQL.Show vbModal, frmCreateTable
            'Else
                'Call CreatTableSQL
            'End If
    End With
End Sub
Sub CreatTableSQL()
    On Error GoTo cRsError
    Set cRs = New ADODB.Recordset
    cRs.Open "" & frmCreateTable.rtSQL.Text & " ", db, adOpenDynamic, adLockBatchOptimistic
cRsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error Occured (Creating Table) : " & Err.Description & "", "Error"
    Else
        gDmsg.DebMsgBox "Table Created Successfully", "Successful", DebmsgInformation
    End If
End Sub

