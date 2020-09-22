VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTableDesc 
   Caption         =   "Table Description"
   ClientHeight    =   5565
   ClientLeft      =   1710
   ClientTop       =   2010
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8655
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   2370
      Top             =   5685
   End
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   5010
      TabIndex        =   12
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   4665
      TabIndex        =   11
      Top             =   270
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   4680
      TabIndex        =   10
      Top             =   135
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   3795
      TabIndex        =   9
      Top             =   90
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin VB.TextBox txtCombo 
      Height          =   300
      Left            =   4800
      TabIndex        =   8
      Top             =   1065
      Visible         =   0   'False
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   1275
      Left            =   2730
      TabIndex        =   7
      Top             =   4110
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2249
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmTableDesc.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2535
      Left            =   2745
      TabIndex        =   6
      Top             =   1515
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1845
      Top             =   4425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":007C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":0956
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":1630
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":1F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":2BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":34BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":4198
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTableDesc.frx":4A72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3855
      Left            =   150
      TabIndex        =   5
      Top             =   1515
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6800
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbSchema 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1050
      Width           =   2610
   End
   Begin DebaFrmCtl.DGrad cmdLoad 
      Height          =   345
      Left            =   2850
      TabIndex        =   3
      Top             =   630
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      OnMouseMoveGradient=   3
      Caption         =   "Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmTableDesc.frx":574C
      MousePointer    =   99
      ScaleHeight     =   23
      ScaleMode       =   0
      ScaleWidth      =   79
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdShowSchema 
      Height          =   360
      Left            =   1635
      TabIndex        =   1
      Top             =   255
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      OnMouseMoveGradient=   3
      Caption         =   "Show Schema"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmTableDesc.frx":58AE
      MousePointer    =   99
      ScaleHeight     =   24
      ScaleMode       =   0
      ScaleWidth      =   92
   End
   Begin VB.Label lblSchemaName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Schema Name And Click On Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   675
      Width           =   2820
   End
   Begin VB.Label lblShowSchema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click On Show Schema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   1605
   End
End
Attribute VB_Name = "frmTableDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh


Option Explicit
Dim rs As New ADODB.Recordset
Dim t As Node
Dim l As ListItem
Private Sub cmdLoad_Click()
    Screen.MousePointer = vbHourglass
On Error GoTo rsError
    tv.Nodes.Clear
    Set t = tv.Nodes.Add(, , "Table", "Table", 1, 2)
    t.Bold = True
    t.Expanded = True
    txtCombo.Text = ""
    txtCombo.Text = Trim$(cmbSchema.Text)
    txtCombo.Text = Trim$(txtCombo.Text)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & txtCombo.Text & "", Empty, "Table"))
    If rs.RecordCount <> 0 Then
        Do Until rs.EOF
            Set t = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
            rs.MoveNext
        Loop
        t.Expanded = True
    Else
        Set t = tv.Nodes.Add("Table", tvwChild, , "No Table", 3, 4)
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error: " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub cmdShowSchema_Click()
    lblShowSchema.Visible = False
    lblSchemaName.Visible = True
    cmbSchema.Visible = True
    cmdLoad.Visible = True
    cmdShowSchema.Visible = False
End Sub
Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    With lv.ColumnHeaders
        .Add , , "Sr.", 34
        .Add , , "COLUMN NAME", 167
        .Add , , "FLAGS", 100
        .Add , , "NULL", 120
        .Add , , "DATATYPE", 200
        .Add , , "SIZE", 80
        .Add , , "PRECISION", 80
        .Add , , "SCALE", 80
    End With
    lblSchemaName.Visible = False
    cmbSchema.Visible = False
    cmdLoad.Visible = False
    Timer1.Enabled = False
    Timer1.Interval = 1200
    Timer1.Enabled = True
    If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
End Sub
Private Sub Form_Resize()
    Me.ScaleMode = vbPixels
    On Error Resume Next
    If Me.Width < 9000 Then
        Me.Width = 9000
    End If
    If Me.Height < 6000 Then
        Me.Height = 6000
    End If
    lblShowSchema.Move 16, 45
    cmdShowSchema.Move lblShowSchema.Left + lblShowSchema.Width + 10, lblShowSchema.Top - 2
    lblSchemaName.Move 16, 45
    cmbSchema.Move lblSchemaName.Left + lblSchemaName.Width + 10, lblSchemaName.Top - 2
    cmdLoad.Move cmbSchema.Left + cmbSchema.Width + 10, lblSchemaName.Top - 2
    tv.Move 16, cmbSchema.Top + cmbSchema.Height + 10, 200, Me.ScaleHeight - 120
    lv.Move tv.Left + tv.Width + 8, tv.Top, Me.ScaleWidth - tv.Width - 46, (tv.Top + tv.Height) / 2 - 5
    rt.Move lv.Left, lv.Top + lv.Height + 5, lv.Width, (tv.Height - lv.Height) - 5
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then
        rs.Close
    End If
    frmMain.Show
End Sub
Private Sub Timer1_Timer()
    Screen.MousePointer = vbHourglass
    On Error GoTo rsError
    Set rs = db.OpenSchema(adSchemaSchemata, Array(Empty, Empty, Empty))
    Do Until rs.EOF
        cmbSchema.AddItem rs!SCHEMA_NAME
        rs.MoveNext
    Loop
    cmbSchema.ListIndex = 0
    If rs.State = 1 Then
        rs.Close
    End If
    Timer1.Enabled = False
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Failed To Retrieve Schema Name : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Screen.MousePointer = vbHourglass
On Error GoTo rsError
    If Node.Key = "TT" & Node.Text Then
    Dim c As Integer
    c = 1
    lv.ListItems.Clear
    Set rs = db.OpenSchema(adSchemaColumns, Array(Empty, "SCOTT", "" & Node.Text & ""))
    Do Until rs.EOF
        Set l = lv.ListItems.Add(, , " " & c & "")
        l.SubItems(1) = rs!COLUMN_NAME
        l.SubItems(2) = rs!COLUMN_FLAGS
        If rs!IS_NULLABLE = 0 Then
            l.SubItems(3) = "NOT NULL"
        Else
            If rs!IS_NULLABLE = -1 Then
                l.SubItems(3) = "NULL"
            Else
                l.SubItems(3) = "UNKNOWN"
            End If
        End If
        Call SetDataType
        c = c + 1
    rs.MoveNext
    Loop
    Call KeyUsed
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub SetDataType()
    Dim x
    Dim d As Integer
    d = rs!DATA_TYPE
    Select Case d
        Case 5 ' Float
            
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "FLOAT"
                l.SubItems(6) = rs!NUMERIC_PRECISION
                    If rs!NUMERIC_SCALE <> "" Then
                        l.SubItems(7) = rs!NUMERIC_SCALE
                    Else
                        l.SubItems(7) = "NULL"
                    End If
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
        
        Case 128 ' Raw Or Long Raw
            
            If l.SubItems(2) = 104 Then
                l.SubItems(4) = "RAW Or LONG RAW"
                l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
            
        Case 129 ' Char Or Varchar2 Or Long Raw
            
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "CHAR"
                l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                If l.SubItems(2) = 104 Or l.SubItems(2) = 8 Then
                    l.SubItems(4) = "VARCHAR2"
                    l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                    If l.SubItems(2) = 232 Then
                        l.SubItems(4) = "LONG"
                        l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
                    Else
                        l.SubItems(4) = "UNKNOWN"
                    End If
                End If
            End If
        
        Case 131
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "NUMBER"
                l.SubItems(6) = rs!NUMERIC_PRECISION
                    If rs!NUMERIC_SCALE <> "" Then
                        l.SubItems(7) = rs!NUMERIC_SCALE
                    Else
                        l.SubItems(7) = "NULL"
                    End If
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
        
        Case 135
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "DATE"
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
            
        Case Else
            l.SubItems(4) = "UNKNOWN"
    End Select
End Sub
Private Sub KeyUsed()
    Screen.MousePointer = vbHourglass
    Dim i
    rt.Text = ""
    rt.SelBold = True
    rt.SelColor = &H8000&
    rt.SelText = " Table Name : - " & tv.SelectedItem.Text & "   " & vbCrLf
    rt.SelText = "--------------------------------------------------" & vbCrLf
    
    rt.SelColor = vbBlack
    rt.SelText = "Tablespace: - "
    Set rs = New ADODB.Recordset
        rs.Open "Select TABLESPACE_NAME From ALL_ALL_TABLES WHERE OWNER='" & txtCombo.Text & "' AND TABLE_NAME='" & tv.SelectedItem.Text & "'", db, adOpenDynamic, adLockBatchOptimistic
        rt.SelColor = vbBlue
        rt.SelText = rs!TABLESPACE_NAME & vbCrLf
        rt.SelText = vbCrLf
        
        rt.SelColor = vbBlack
        rt.SelText = "Created Time: - "
    Set rs = New ADODB.Recordset
        rs.Open "Select CREATED From ALL_OBJECTS Where OWNER='" & txtCombo.Text & "' AND OBJECT_NAME='" & tv.SelectedItem.Text & "' AND  OBJECT_TYPE='TABLE'", db, adOpenDynamic, adLockBatchOptimistic
        rt.SelColor = vbBlue
        rt.SelText = rs!CREATED & vbCrLf
        rt.SelText = vbCrLf
        
        rt.SelColor = vbBlack
        rt.SelText = "Primary Key : - "
    Set rs = db.OpenSchema(adSchemaPrimaryKeys, Array(Empty, "" & txtCombo.Text & "", "" & tv.SelectedItem.Text & ""))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!COLUMN_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!COLUMN_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
        
        rt.SelColor = vbBlack
        rt.SelText = "Foreign Key Table Name : - "
    Set rs = db.OpenSchema(adSchemaForeignKeys, Array(Empty, "" & txtCombo.Text & "", "" & tv.SelectedItem.Text & "", Empty, "" & txtCombo.Text & "", Empty))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!FK_TABLE_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!FK_TABLE_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
        
        rt.SelColor = vbBlack
        rt.SelText = "Foreign Key  : - "
    Set rs = db.OpenSchema(adSchemaForeignKeys, Array(Empty, "" & txtCombo.Text & "", "" & tv.SelectedItem.Text & "", Empty, "" & txtCombo.Text & "", Empty))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!FK_COLUMN_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!FK_COLUMN_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
    
        rt.SelColor = vbBlack
        On Error Resume Next
        rt.SelText = "Index Name  : - "
    Set rs = db.OpenSchema(adSchemaIndexes, Array(Empty, "" & txtCombo.Text & "", Empty, Empty, "" & tv.SelectedItem.Text & ""))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!INDEX_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!INDEX_NAME
                rs.MoveNext
            Loop
        End If
    Screen.MousePointer = vbDefault
End Sub


