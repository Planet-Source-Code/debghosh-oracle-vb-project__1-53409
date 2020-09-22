VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditRecordset 
   Caption         =   "Edit Recordset"
   ClientHeight    =   6750
   ClientLeft      =   1080
   ClientTop       =   1380
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9165
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   6150
      TabIndex        =   19
      Top             =   45
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   5850
      TabIndex        =   18
      Top             =   210
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   5850
      TabIndex        =   17
      Top             =   75
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   4995
      TabIndex        =   16
      Top             =   30
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin DebaFrmCtl.DLbl cmdHelp 
      Height          =   240
      Left            =   150
      TabIndex        =   15
      Top             =   6150
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   423
      BackColor       =   16711680
      Caption         =   "Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   16711935
      Picture         =   "frmEditRecordset.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":3082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":395C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":4636
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":4F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":5BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":64C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":719E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":7A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":8752
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":902C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":9D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditRecordset.frx":A5E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   1335
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   6165
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCombo 
      Height          =   315
      Left            =   975
      TabIndex        =   10
      Top             =   6300
      Visible         =   0   'False
      Width           =   2025
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   4005
      Left            =   2505
      TabIndex        =   9
      Top             =   1080
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   7064
      _Version        =   393216
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin DebaFrmCtl.DGrad cmdShowSchema 
      Height          =   330
      Left            =   2100
      TabIndex        =   8
      Top             =   5925
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      OnMouseMoveGradient=   3
      Caption         =   "Show Schema"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmEditRecordset.frx":B2BA
      MousePointer    =   99
      ScaleHeight     =   22
      ScaleMode       =   0
      ScaleWidth      =   111
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdLoad 
      Height          =   315
      Left            =   945
      TabIndex        =   7
      Top             =   5925
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
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
      MouseIcon       =   "frmEditRecordset.frx":B41C
      MousePointer    =   99
      ScaleHeight     =   21
      ScaleMode       =   0
      ScaleWidth      =   76
      OnMouseMoveForeColor=   12648447
   End
   Begin VB.ComboBox cmbSchema 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5580
      Width           =   2745
   End
   Begin DebaFrmCtl.DLbl lblUpdate 
      Height          =   195
      Left            =   5415
      TabIndex        =   5
      Top             =   6135
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   344
      Caption         =   "Update"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   49152
      PictureVisible  =   0   'False
   End
   Begin DebaFrmCtl.DLbl lblDRec 
      Height          =   195
      Left            =   5520
      TabIndex        =   4
      Top             =   5880
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   344
      Caption         =   "Update"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   49152
      PictureVisible  =   0   'False
   End
   Begin DebaFrmCtl.DLbl lblCancel 
      Height          =   195
      Left            =   6285
      TabIndex        =   3
      Top             =   5925
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   344
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   33023
      PictureVisible  =   0   'False
   End
   Begin DebaFrmCtl.DLbl lblDeleteRec 
      Height          =   195
      Left            =   3885
      TabIndex        =   2
      Top             =   6120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   344
      Caption         =   "Delete Record"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   255
      PictureVisible  =   0   'False
   End
   Begin DebaFrmCtl.DLbl lblAddnew 
      Height          =   195
      Left            =   3885
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   344
      Caption         =   "Add New Record"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      OnMouseMoveForeColor=   49152
      PictureVisible  =   0   'False
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4110
      TabIndex        =   20
      Top             =   5325
      Width           =   990
   End
   Begin VB.Label lblTblOrViewName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   5655
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblSchema 
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
      Left            =   1065
      TabIndex        =   12
      Top             =   5325
      Width           =   2820
   End
   Begin VB.Label lblShowSchema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Schema Name"
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
      Left            =   1005
      TabIndex        =   11
      Top             =   5070
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT RECORD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3915
      TabIndex        =   0
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   885
      Left            =   3795
      Shape           =   4  'Rounded Rectangle
      Top             =   5595
      Width           =   3210
   End
End
Attribute VB_Name = "frmEditRecordset"
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
Private Sub cmdHelp_Click()
    frmEditRecHelp.Show vbModal, Me
End Sub
Private Sub cmdLoad_Click()
    Screen.MousePointer = vbHourglass
    On Error GoTo rsError
    tv.Nodes.Clear
    Set t = tv.Nodes.Add(, , "MAIN", "" & cmbSchema.Text & "", 1, 2)
    t.Bold = True
    t.Expanded = True
    txtCombo.Text = ""
    txtCombo.Text = Trim$(cmbSchema.Text)
    txtCombo.Text = Trim$(txtCombo.Text)
    Set t = tv.Nodes.Add("MAIN", tvwChild, "Table", "Table", 3, 4)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "Table"))
    If rs.RecordCount <> 0 Then
        Do Until rs.EOF
            Set t = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME, 7, 8)
            rs.MoveNext
        Loop
        t.Expanded = True
    Else
        Set t = tv.Nodes.Add("Table", tvwChild, , "No Table", 7, 8)
    End If
    Set t = tv.Nodes.Add("MAIN", tvwChild, "View", "View", 5, 6)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "View"))
    If rs.RecordCount <> 0 Then
        Do Until rs.EOF
            Set t = tv.Nodes.Add("View", tvwChild, "VV" & rs!TABLE_NAME, rs!TABLE_NAME, 7, 8)
            rs.MoveNext
        Loop
        t.Expanded = True
    Else
        Set t = tv.Nodes.Add("View", tvwChild, , "No View", 7, 8)
    End If
    If rs.State = 1 Then
        rs.Close
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
    Screen.MousePointer = vbHourglass
    On Error GoTo rsError
    Set rs = db.OpenSchema(adSchemaSchemata)
    Do Until rs.EOF
        cmbSchema.AddItem rs!SCHEMA_NAME
        rs.MoveNext
    Loop
    If rs.State = 1 Then
        rs.Close
    End If
    cmbSchema.ListIndex = 0
    lblShowSchema.Visible = False
    cmdShowSchema.Visible = False
    lblSchema.Visible = True
    cmbSchema.Visible = True
    cmdLoad.Visible = True
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & "", "Error"
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    lblSchema.Visible = False
    cmbSchema.Visible = False
    cmdLoad.Visible = False
    cmdShowSchema.Visible = True
    Call DefLabel
    If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
    lblName.Visible = False
    Call FormToolTip
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
End Sub
Private Sub Form_Resize()
    Me.ScaleMode = vbPixels
    On Error Resume Next
    If Me.Width < 9300 Then
        Me.Width = 9300
    End If
    If Me.Height < 7200 Then
        Me.Height = 7200
    End If
    lblShowSchema.Move 16, 44
    cmdShowSchema.Move lblShowSchema.Left + lblShowSchema.Width + 10, 42
    lblSchema.Move 16, 44
    cmbSchema.Move lblSchema.Left + lblSchema.Width + 12, 42
    cmdLoad.Move cmbSchema.Left + cmbSchema.Width + 10, cmbSchema.Top
    lblName.Move cmdLoad.Left + cmdLoad.Width + 10, cmdLoad.Top + 4
    lblTblOrViewName.Move cmdLoad.Left + cmdLoad.Width + 10, cmdLoad.Top
    tv.Move 16, cmdShowSchema.Top + cmdShowSchema.Height + 10, 200, Me.ScaleHeight - 140
    dg.Move tv.Left + tv.Width + 10, tv.Top, Me.ScaleWidth - tv.Width - 40, tv.Height - 40
    Shape2.Move ((dg.Left + dg.Width)) / 2, dg.Top + dg.Height + 5
    lblEdit.Move ((dg.Left + dg.Width)) / 2 + 10, dg.Top + dg.Height + 14
    lblAddnew.Move lblEdit.Left, lblEdit.Top + lblEdit.Height + 15
    lblDeleteRec.Move Shape2.Left + Shape2.Width - lblDeleteRec.Width - 4, lblAddnew.Top
    lblUpdate.Move lblAddnew.Left, lblAddnew.Top
    lblCancel.Move lblDeleteRec.Left, lblDeleteRec.Top
    lblDRec.Move lblUpdate.Left, lblUpdate.Top
    cmdHelp.Move dg.Left + dg.Width - cmdHelp.Width - 10, lblEdit.Top
End Sub
Sub DefDatagrid()
    With dg
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .MarqueeStyle = dbgSolidCellBorder
        .TabAction = dbgGridNavigation
        .WrapCellPointer = True
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then
        rs.Close
    End If
    frmMain.Show
End Sub
Private Sub lblAddnew_Click()
On Error GoTo rsError
    rs.AddNew
    lblAddnew.Visible = False
    lblDeleteRec.Visible = False
    lblUpdate.Visible = True
    lblCancel.Visible = True
    lblDRec.Visible = False
    With dg
        .AllowAddNew = True
        .AllowDelete = True
        .AllowUpdate = True
        .SetFocus
    End With
rsError:
    If Err.Number <> 0 Then
        lblCancel.Visible = False
        lblUpdate.Visible = False
        lblAddnew.Visible = True
        lblDeleteRec.Visible = True
        lblDRec.Visible = False
        With dg
            .AllowAddNew = False
            .AllowDelete = False
            .AllowUpdate = False
            .SetFocus
        End With
        gDmsg.DebMsgBox "Error: " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub lblCancel_Click()
On Error GoTo rsError
    lblCancel.Visible = False
    lblUpdate.Visible = False
    lblAddnew.Visible = True
    lblDeleteRec.Visible = True
    lblDRec.Visible = False
    With dg
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .SetFocus
    End With
    rs.CancelUpdate
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub lblDeleteRec_Click()
On Error GoTo rsError
    lblAddnew.Visible = False
    lblDeleteRec.Visible = False
    lblUpdate.Visible = False
    lblDRec.Visible = True
    lblCancel.Visible = True
    With dg
        .AllowAddNew = False
        .AllowDelete = True
        .AllowUpdate = True
        .SetFocus
    End With
rsError:
    If Err.Number <> 0 Then
        lblCancel.Visible = False
        lblUpdate.Visible = False
        lblAddnew.Visible = True
        lblDeleteRec.Visible = True
        lblDRec.Visible = False
        With dg
            .AllowAddNew = False
            .AllowDelete = False
            .AllowUpdate = False
            .SetFocus
        End With
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Private Sub lblDRec_Click()
    On Error GoTo rsError
    lblAddnew.Visible = True
    lblDeleteRec.Visible = True
    lblDRec.Visible = False
    lblUpdate.Visible = False
    lblCancel.Visible = False
    With dg
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .MarqueeStyle = dbgSolidCellBorder
        .SetFocus
    End With
    Dim rM
    rM = gDmsg.DebMsgBox("Save All Changes", "Confirmation", DebmsgYesNo)
    If rM = DebmsgYes Then
        rs.Delete adAffectCurrent
        rs.UpdateBatch adAffectAllChapters
        db.CommitTrans
        db.BeginTrans
        Set rs = New ADODB.Recordset
            rs.Open "SELECT * FROM " & lblName.Caption & "", db, adOpenDynamic, adLockBatchOptimistic
        Set dg.DataSource = rs
    Else
        If rM = DebmsgNo Then
            rs.CancelUpdate
            db.RollbackTrans
            db.BeginTrans
        End If
    End If
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error : " & Err.Description & "", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Private Sub lblUpdate_Click()
    On Error GoTo rsError
    lblAddnew.Visible = True
    lblDeleteRec.Visible = True
    lblDRec.Visible = False
    lblUpdate.Visible = False
    lblCancel.Visible = False
    With dg
        .AllowAddNew = False
        .AllowDelete = False
        .AllowUpdate = False
        .SetFocus
    End With
    Dim rM
    rM = gDmsg.DebMsgBox("Save All Changes", "Confirmation", DebmsgYesNo)
    If rM = DebmsgYes Then
        rs.UpdateBatch adAffectAllChapters
        db.CommitTrans
        db.BeginTrans
        Set rs = New ADODB.Recordset
            rs.Open "SELECT * FROM " & lblName.Caption & "", db, adOpenDynamic, adLockBatchOptimistic
        Set dg.DataSource = rs
    Else
        If rM = DebmsgNo Then
            rs.CancelUpdate
            db.RollbackTrans
            db.BeginTrans
        End If
    End If
rsError:
    If Err.Number <> 0 Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM " & lblName.Caption & "", db, adOpenDynamic, adLockBatchOptimistic
        Set dg.DataSource = rs
        gDmsg.DebMsgBox "Error : " & Err.Description & "", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Screen.MousePointer = vbHourglass
On Error GoTo rsError
    If Node.Key = "TT" & Node.Text Or Node.Key = "VV" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
        Set dg.DataSource = rs
        Call DefDatagrid
        dg.SetFocus
        lblName.Visible = False
        lblName.Caption = Trim$(Node.Text)
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Sub DefLabel()
    Shape2.ZOrder 1
    lblEdit.Visible = True
    lblAddnew.Visible = True
    lblAddnew.BackColor = RGB(0, 0, 240)
    lblUpdate.Visible = False
    lblUpdate.BackColor = RGB(0, 0, 240)
    lblCancel.Visible = False
    lblCancel.BackColor = RGB(0, 0, 240)
    lblDeleteRec.Visible = True
    lblDeleteRec.BackColor = RGB(0, 0, 240)
    lblDRec.Visible = False
    lblDRec.BackColor = RGB(0, 0, 240)
    lblTblOrViewName.Visible = False
End Sub
Sub FormToolTip()
    modToolTip.CreateBalloon lblAddnew, lblAddnew.hwnd, "Add New Record", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon lblDeleteRec, lblDeleteRec.hwnd, "Delete Record", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon lblDRec, lblDRec.hwnd, "Update Recordset", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon lblCancel, lblCancel.hwnd, "Cancel Update", szBalloon, False, Me.Caption, etiError
    modToolTip.CreateBalloon lblUpdate, lblUpdate.hwnd, "Update Recordset", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdHelp, cmdHelp.hwnd, "Click Here For Help", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdLoad, cmdLoad.hwnd, "Load Table And View In Treeview", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdShowSchema, cmdShowSchema.hwnd, "Show Schema Name", szBalloon, False, Me.Caption, etiInfo
End Sub

