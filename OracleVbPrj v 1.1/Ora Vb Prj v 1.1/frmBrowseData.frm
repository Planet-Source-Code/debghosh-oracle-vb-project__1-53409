VERSION 5.00
Object = "*\A..\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrowseData 
   Caption         =   "Browse Data"
   ClientHeight    =   6855
   ClientLeft      =   1395
   ClientTop       =   1410
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8775
   Begin MSComDlg.CommonDialog cd 
      Left            =   8835
      Top             =   6915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   7380
      TabIndex        =   19
      Top             =   210
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   7695
      TabIndex        =   18
      Top             =   390
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   7740
      TabIndex        =   17
      Top             =   240
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   7410
      TabIndex        =   16
      Top             =   585
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   190
      Left            =   75
      TabIndex        =   15
      Top             =   240
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtSchemaName 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4200
      TabIndex        =   14
      Top             =   6855
      Visible         =   0   'False
      Width           =   2505
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   465
      Top             =   2955
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
            Picture         =   "frmBrowseData.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":15B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":2B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":3442
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":411C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":49F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":56D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":5FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":6C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowseData.frx":755E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3090
      Left            =   60
      TabIndex        =   6
      Top             =   1215
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5450
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
   Begin DebaFrmCtl.DGrad cmdSave 
      Height          =   300
      Left            =   1095
      TabIndex        =   5
      Top             =   4335
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
      OnMouseMoveGradient=   3
      Caption         =   "Save As"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBrowseData.frx":8238
      MousePointer    =   99
      ScaleHeight     =   20
      ScaleMode       =   0
      ScaleWidth      =   66
   End
   Begin DebaFrmCtl.DGrad cmdSearch 
      Height          =   300
      Left            =   45
      TabIndex        =   4
      Top             =   4335
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
      OnMouseMoveGradient=   3
      Caption         =   "Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBrowseData.frx":839A
      MousePointer    =   99
      ScaleHeight     =   20
      ScaleMode       =   0
      ScaleWidth      =   66
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7545
      Top             =   6915
   End
   Begin DebaFrmCtl.DGrad cmdLoad 
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   585
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      OnMouseMoveGradient=   3
      Caption         =   "Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBrowseData.frx":84FC
      MousePointer    =   99
      ScaleHeight     =   22
      ScaleMode       =   0
      ScaleWidth      =   82
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2475
   End
   Begin VB.TextBox tvSelectedItem 
      Height          =   285
      Left            =   15
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   2385
      ScaleHeight     =   4425
      ScaleWidth      =   4830
      TabIndex        =   7
      Top             =   1215
      Width           =   4860
      Begin DebaFrmCtl.DLbl dLblExprtToExl 
         Height          =   300
         Left            =   2700
         TabIndex        =   9
         Top             =   3960
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         BackColor       =   16777215
         Caption         =   "Export To Excel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnMouseMoveForeColor=   255
         Picture         =   "frmBrowseData.frx":865E
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Fg 
         Height          =   2340
         Left            =   240
         TabIndex        =   8
         Top             =   1515
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   4128
         _Version        =   393216
         ForeColor       =   10485760
         BackColorFixed  =   14408667
         ForeColorFixed  =   4194304
         BackColorBkg    =   15724527
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblTableName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   255
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   2325
      ScaleHeight     =   4800
      ScaleWidth      =   5085
      TabIndex        =   10
      Top             =   1170
      Width           =   5115
      Begin RichTextLib.RichTextBox rtSQL 
         Height          =   3180
         Left            =   300
         TabIndex        =   13
         Top             =   1635
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   5609
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmBrowseData.frx":891C
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
      Begin DebaFrmCtl.DLbl dLblSaveAs 
         Height          =   300
         Left            =   3615
         TabIndex        =   12
         Top             =   4500
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         BackColor       =   16777215
         Caption         =   "Save As"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnMouseMoveForeColor=   255
         Picture         =   "frmBrowseData.frx":8998
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgRt 
         Height          =   885
         Left            =   4500
         TabIndex        =   11
         Top             =   4785
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1561
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblProc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procedure"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   21
         Top             =   795
         Width           =   870
      End
   End
   Begin VB.Label lblSelectSchemaName 
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
      TabIndex        =   0
      Top             =   930
      Width           =   2820
   End
End
Attribute VB_Name = "frmBrowseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh


Option Explicit
Dim pbr As New clsProgreeBar
Dim pp As New clsPicturePaint
Dim rs As New ADODB.Recordset
Dim prs As New ADODB.Recordset
Dim t As Node
Dim n As Node
Dim pic1 As Boolean
Dim pic2 As Boolean
Private Sub cmdLoad_Click()
On Error GoTo rsError
    Dim c As Integer
    Screen.MousePointer = vbHourglass
    txtSchemaName.Text = Trim$(cmbSchema.Text)
    tv.Nodes.Clear
    Set n = tv.Nodes.Add(, , "ORACLE", "" & cmbSchema.Text & " ", 1, 2)
    n.Expanded = True

    'Load Table In Treeview
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Table", "Table", 5, 6)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "Table"))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        
    'Load View In Treeview
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "View", "View", 7, 8)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "View"))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("View", tvwChild, "VV" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        
    'Load Sequence
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Sequence", "Sequence", 9, 10)
    Set rs = New ADODB.Recordset
        rs.Open "Select * from ALL_SEQUENCES where SEQUENCE_OWNER='" & cmbSchema.Text & "'", db, adOpenDynamic, adLockBatchOptimistic
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Sequence", tvwChild, "SQ" & rs!SEQUENCE_NAME, rs!SEQUENCE_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
    
    'Load Synonyms
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Synonyms", "Synonyms", 11, 12)
    Set rs = New ADODB.Recordset
    rs.Open "Select * from ALL_SYNONYMS Where OWNER='" & cmbSchema.Text & "'", db, adOpenDynamic, adLockBatchOptimistic
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Synonyms", tvwChild, "SS" & rs!SYNONYM_NAME, rs!SYNONYM_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
    
    'Load Procedure
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Procedure", "Procedure", 5, 6)
    Set rs = db.OpenSchema(adSchemaProcedures, Array(Empty, "" & cmbSchema.Text & "", Empty, Empty))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Procedure", tvwChild, "PP" & rs!PROCEDURE_NAME, rs!PROCEDURE_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
    
    'Load Package
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Package", "Package", 7, 8)
    Set rs = New ADODB.Recordset
    rs.Open "Select * from ALL_SOURCE Where OWNER='" & cmbSchema.Text & "' AND Type='PACKAGE'", db, adOpenDynamic, adLockBatchOptimistic
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                On Error Resume Next
                Set n = tv.Nodes.Add("Package", tvwChild, "PK" & rs!Name, rs!Name, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
                On Error Resume Next
            Loop
        End If
        pb.Visible = False
        
    'Load Package Body
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "PackageBody", "Package Body", 9, 10)
    Set rs = New ADODB.Recordset
    rs.Open "Select * From ALL_SOURCE Where OWNER='" & cmbSchema.Text & "' AND Type='PACKAGE BODY'", db, adOpenDynamic, adLockBatchOptimistic
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                On Error Resume Next
                Set n = tv.Nodes.Add("PackageBody", tvwChild, "PKD" & rs!Name, rs!Name, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
                On Error Resume Next
            Loop
        End If
        pb.Visible = False
    
    'Load Type
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Type", "Type", 11, 12)
    Set rs = New ADODB.Recordset
        rs.Open "Select * From ALL_TYPES where OWNER='" & cmbSchema.Text & "' ", db, adOpenDynamic, adLockBatchOptimistic
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Type", tvwChild, "TP" & rs!TYPE_NAME, rs!TYPE_NAME, 3, 4)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        
    lblSelectSchemaName.Visible = True
    cmbSchema.Visible = True
    cmdLoad.Visible = True
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox Err.Description, "Error"
        Exit Sub
    End If
End Sub
Private Sub cmdSave_Click()
    On Error GoTo tvError
    If tv.Nodes.Count <> 0 Then
        Call modTVSearch.SaveData(cd, tv)
    Else
        gDmsg.DebMsgBox "No Data", "Error", DebmsgExclamation
    End If
tvError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description, "Error"
        Exit Sub
    End If
End Sub
Private Sub cmdSearch_Click()
    On Error GoTo tvError
    If tv.Nodes.Count <> 0 Then
        Call modTVSearch.FindNode(tv)
    Else
        gDmsg.DebMsgBox "No Data", "Error", DebmsgExclamation
    End If
tvError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description, "Error"
        Exit Sub
    End If
End Sub
Private Sub dLblExprtToExl_Click()
On Error GoTo fError:
    If tvSelectedItem.Text <> "" Then
        modExportToExcel.Exp2Exl Fg, tv.SelectedItem.Text
    Else
        gDmsg.DebMsgBox "Please Select Table Or View", "Error", DebmsgExclamation
    End If
fError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description, "Error"
        Exit Sub
    End If
End Sub
Private Sub dLblSaveAs_Click()
    If rtSQL.Text <> "" Then
        modRtSaveAs.SaveTextAs rtSQL, cd
    End If
End Sub
Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    Set pp = New clsPicturePaint
    pp.PictureIni Picture1, 0, 150, "Table Or View"
    pp.PictureIni Picture2, 151, 250, "Procedure"
    Set pbr = New clsProgreeBar
    pbr.DProgressBar pb, cc3D, DRed, Standard
    pb.Visible = False
    Timer1.Enabled = False
    Timer1.Interval = 1000
    Timer1.Enabled = True
    pic1 = True
    pic2 = False
    modRtColor.InitWords
    modRtColor.DoColor rtSQL
    rtSQL.SelStart = 0
    Fg.ColWidth(0) = 300
    If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
    Call FormToolTip
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Me.ScaleMode = vbPixels
    If Me.Width < 8900 Then
        Me.Width = 8900
    End If
    If Me.Height < 7300 Then
        Me.Height = 7300
    End If
    lblSelectSchemaName.Move 16, 52
    cmbSchema.Move lblSelectSchemaName.Left + lblSelectSchemaName.Width + 10, 50
    cmdLoad.Move cmbSchema.Left + cmbSchema.Width + 10, 50
    pb.Move 16, 36, Me.ScaleWidth - 36
    tv.Move 16, cmbSchema.Top + cmbSchema.Height + 10, 200, Me.ScaleHeight - 150
    cmdSearch.Move 16, tv.Top + tv.Height + 10, tv.Width / 2 - 5
    cmdSave.Move cmdSearch.Left + cmdSearch.Width + 10, cmdSearch.Top, cmdSearch.Width
    Picture1.Move tv.Left + tv.Width + 8, tv.Top, Me.ScaleWidth - tv.Width - 36, Me.ScaleHeight - 120
    Picture2.Move Picture1.Left, Picture1.Top, Picture1.ScaleWidth, Picture1.ScaleHeight
    lblTableName.Move 10, 42
    Fg.Move 10, 60, Picture1.ScaleWidth - 30, Picture1.ScaleHeight - 90
    dLblExprtToExl.Move Fg.Left + Fg.Width - dLblExprtToExl.Width, Fg.Top + Fg.Height + 4
    lblProc.Move 10, 42
    rtSQL.Move 10, 60, Fg.Width, Fg.Height
    dLblSaveAs.Move rtSQL.Left + rtSQL.Width - dLblSaveAs.Width, rtSQL.Top + rtSQL.Height + 4
    If pic1 = True Then
        Call Picture1_Click
    End If
    If pic2 = True Then
        Call Picture2_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then
        rs.Close
    End If
    If prs.State = 1 Then
        prs.Close
    End If
    Set pp = Nothing
    Set pbr = Nothing
    frmMain.Show
End Sub
Private Sub Picture1_Click()
    pic1 = True
    pic2 = False
    pp.PictureClick Picture1, 0, 150, "Table Or View"
    pp.PictureIni Picture2, 151, 250, "Procedure"
End Sub
Private Sub Picture2_Click()
    pic1 = False
    pic2 = True
    pp.PictureClick Picture2, 151, 250, "Procedure"
    pp.PictureIni Picture1, 0, 150, "Table Or View"
End Sub
Private Sub rtSQL_Change()
    Dim lCursor As Long
    lCursor = rtSQL.SelStart
    modRtColor.DoColor rtSQL
    rtSQL.SelStart = lCursor
    rtSQL.SelColor = vbBlack
End Sub
Private Sub Timer1_Timer()
    frmProcess.Show vbModal
    Timer1.Enabled = False
End Sub
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo rsError
    Screen.MousePointer = vbHourglass
    Dim i As Integer
    If Node.Expanded = True Then
        Node.Expanded = False
    Else
        Node.Expanded = True
    End If
    If Node.Key = "TT" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            lblTableName.Caption = "Table Name: " & Node.Text & " "
            Call FlexGridRow
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "VV" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            lblTableName.Caption = "View Name: " & Node.Text & " "
            Call FlexGridRow
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "SQ" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            lblTableName.Caption = "Sequence Name: " & Node.Text & " "
            Call FlexGridRow
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "SS" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            lblTableName.Caption = "Synonym Name : " & Node.Text & " "
            Call FlexGridRow
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "PP" & Node.Text Then
        rtSQL.Text = ""
        Set prs = db.OpenSchema(adSchemaProcedureParameters, Array(Empty, "" & txtSchemaName.Text & "", "" & Node.Text & "", Empty))
        If prs.RecordCount <> 0 Then
            rtSQL.SelText = "Parameter" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
            Do Until prs.EOF
                rtSQL.SelText = prs!PARAMETER_NAME & vbCrLf
                prs.MoveNext
            Loop
                rtSQL.SelText = vbCrLf
                rtSQL.SelText = vbCrLf
                rtSQL.SelText = "TEXT" & vbCrLf
                rtSQL.SelText = "-------------------------------------------" & vbCrLf
        Else
            rtSQL.SelText = "Parameter" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
            rtSQL.SelText = vbCrLf
            rtSQL.SelText = vbCrLf
            rtSQL.SelText = "TEXT" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
        End If
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & txtSchemaName.Text & "' And Type = 'PROCEDURE' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
                For i = 1 To fgRt.Rows - 1
                    rtSQL.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProc.Caption = "Procedure Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "PK" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & txtSchemaName.Text & "' And Type = 'PACKAGE' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rtSQL.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rtSQL.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProc.Caption = "Package Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "PKD" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & txtSchemaName.Text & "' And Type = 'PACKAGE BODY' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rtSQL.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rtSQL.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProc.Caption = "Package Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "TP" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_TYPE_VERSIONS where OWNER='" & txtSchemaName.Text & "' And TYPE_NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rtSQL.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rtSQL.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProc.Caption = "Package Name : " & Node.Text & " "
            Call Picture2_Click
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error"
        Exit Sub
    End If
End Sub
Sub FlexGridRow()
    Dim fR As Integer
    For fR = 1 To Fg.Rows - 1
        Fg.TextMatrix(fR, 0) = fR
    Next fR
End Sub
Sub FormToolTip()
    modToolTip.CreateBalloon cmdLoad, cmdLoad.hwnd, " Click Here To Load Schema Table,View,Procedure etc", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdSave, cmdSave.hwnd, "Save Treeview Data", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdSearch, cmdSearch.hwnd, "Search Treeview", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmbSchema, cmbSchema.hwnd, "Select Schema And Click On Load", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblExprtToExl, dLblExprtToExl.hwnd, "Export Data To Microsoft Excel", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblSaveAs, dLblSaveAs.hwnd, "Save Text", szBalloon, False, Me.Caption, etiInfo
End Sub

