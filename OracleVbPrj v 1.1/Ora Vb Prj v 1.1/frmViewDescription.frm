VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewDescription 
   Caption         =   "View Description"
   ClientHeight    =   6465
   ClientLeft      =   1755
   ClientTop       =   1545
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8805
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   6945
      TabIndex        =   12
      Top             =   555
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   6840
      TabIndex        =   11
      Top             =   450
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   7020
      TabIndex        =   10
      Top             =   300
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   6765
      TabIndex        =   9
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8865
      Top             =   6525
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   1890
      Left            =   2445
      TabIndex        =   8
      Top             =   3435
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   3334
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmViewDescription.frx":0000
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
      Height          =   1785
      Left            =   2445
      TabIndex        =   7
      Top             =   1605
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   3149
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
      Left            =   945
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":007C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":0956
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":1630
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":1F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":2BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewDescription.frx":34BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3720
      Left            =   210
      TabIndex        =   6
      Top             =   1620
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   6562
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1140
      Width           =   2565
   End
   Begin VB.TextBox txtCombo 
      Height          =   315
      Left            =   4335
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   2175
   End
   Begin DebaFrmCtl.DGrad cmdShowSchema 
      Height          =   330
      Left            =   2445
      TabIndex        =   3
      Top             =   690
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      DefaultGradient =   1
      Caption         =   "Show"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmViewDescription.frx":4198
      MousePointer    =   99
      ScaleHeight     =   22
      ScaleMode       =   0
      ScaleWidth      =   94
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdLoad 
      Height          =   315
      Left            =   3030
      TabIndex        =   1
      Top             =   240
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      DefaultGradient =   1
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
      MouseIcon       =   "frmViewDescription.frx":42FA
      MousePointer    =   99
      ScaleHeight     =   21
      ScaleMode       =   0
      ScaleWidth      =   84
   End
   Begin VB.Label lblShowSchema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click On Show Schema Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   660
      Width           =   2145
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
      Left            =   135
      TabIndex        =   0
      Top             =   300
      Width           =   2820
   End
End
Attribute VB_Name = "frmViewDescription"
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
    txtCombo.Text = ""
    If cmbSchema.Text <> "" Then
    Set t = tv.Nodes.Add(, , "Table", "VIEW", 1, 2)
    t.Bold = True
    t.Expanded = True
    txtCombo.Text = Trim$(cmbSchema.Text)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & txtCombo.Text & "", Empty, "View"))
    If rs.RecordCount <> 0 Then
        Do Until rs.EOF
            Set t = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
            rs.MoveNext
        Loop
    Else
        Set t = tv.Nodes.Add("Table", tvwChild, , "No View", 3, 4)
    End If
    Else
        gDmsg.DebMsgBox "Please Select Schema Name", "Error", DebmsgExclamation
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
    lblSchemaName.Visible = True
    cmbSchema.Visible = True
    cmdLoad.Visible = True
    cmdShowSchema.Visible = False
    lblShowSchema.Visible = False
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
    cmdShowSchema.Visible = True
    lblShowSchema.Visible = True
    Timer1.Enabled = False
    Timer1.Interval = 1000
    Timer1.Enabled = True
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
End Sub
Private Sub Form_Resize()
    Me.ScaleMode = vbPixels
    On Error Resume Next
    If Me.Width < 9000 Then
        Me.Width = 9000
    End If
    If Me.Height < 7000 Then
        Me.Height = 7000
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
    If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
    Screen.MousePointer = vbDefault

rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Screen.MousePointer = vbHourglass
On Error GoTo rsError
    Dim c As Integer
    If Node.Key = "TT" & Node.Text Then
    c = 1
    lv.ListItems.Clear
    Set rs = db.OpenSchema(adSchemaColumns, Array(Empty, "" & txtCombo.Text & "", "" & Node.Text & ""))
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
    rt.Text = ""
    rt.SelText = vbCrLf
    rt.SelBold = False
    rt.SelColor = &H8000&
    rt.SelText = " View Name : - " & tv.SelectedItem.Text & "   " & vbCrLf
    rt.SelText = "--------------------------------------------------" & vbCrLf
    
        rt.SelColor = vbBlack
        rt.SelText = "Created Time: - "
    Set rs = New ADODB.Recordset
        rs.Open "Select CREATED From ALL_OBJECTS Where OWNER='" & txtCombo.Text & "' AND OBJECT_NAME='" & tv.SelectedItem.Text & "' AND  OBJECT_TYPE='VIEW'", db, adOpenDynamic, adLockBatchOptimistic
        rt.SelColor = vbBlue
        rt.SelText = rs!CREATED & vbCrLf
        rt.SelText = vbCrLf
        
        Call KeyUsed
        rt.SelText = vbCrLf
        rt.SelText = vbCrLf
        
        rt.SelColor = vbBlack
        rt.SelText = "                 TEXT" & vbCrLf
        rt.SelColor = vbBlack
        rt.SelText = "--------------------------------------------------" & vbCrLf
        rt.SelText = vbCrLf
        Set rs = New ADODB.Recordset
        rt.SelColor = vbBlue
        rs.Open "Select TEXT From ALL_VIEWS WHERE OWNER='" & Trim$(txtCombo.Text) & "' AND VIEW_NAME='" & Node.Text & "'", db, adOpenDynamic, adLockBatchOptimistic
            If rs!Text <> "" Then
                rt.SelText = rs!Text
            Else
                rt.SelText = "No Text"
            End If
        End If
        Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error " & Err.Description & " ", "Error", DebmsgExclamation
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
