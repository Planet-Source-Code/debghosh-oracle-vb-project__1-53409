VERSION 5.00
Object = "*\A..\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Main Window"
   ClientHeight    =   6240
   ClientLeft      =   900
   ClientTop       =   1470
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   Begin MSComDlg.CommonDialog cd 
      Left            =   8850
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DebaFrmCtl.DLbl lblMenu 
      Height          =   240
      Left            =   5040
      TabIndex        =   17
      Top             =   1035
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   423
      BackColor       =   16711680
      Caption         =   "MENU"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin VB.TextBox tvSelectedItem 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2190
      TabIndex        =   15
      Top             =   2670
      Visible         =   0   'False
      Width           =   1680
   End
   Begin DebaFrmCtl.DGrad cmdSave 
      Height          =   300
      Left            =   1215
      TabIndex        =   14
      Top             =   3810
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   529
      DefaultGradient =   3
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648447
      ScaleHeight     =   20
      ScaleMode       =   3
      ScaleWidth      =   64
      OnMouseMoveForeColor=   16777215
   End
   Begin DebaFrmCtl.DGrad cmdSearch 
      Height          =   300
      Left            =   225
      TabIndex        =   13
      Top             =   3810
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   529
      DefaultGradient =   3
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
      ForeColor       =   12648447
      ScaleHeight     =   20
      ScaleMode       =   3
      ScaleWidth      =   64
      OnMouseMoveForeColor=   16777215
   End
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   8850
      Top             =   6300
   End
   Begin VB.PictureBox Picture2 
      Height          =   1890
      Left            =   2220
      ScaleHeight     =   1830
      ScaleWidth      =   1425
      TabIndex        =   7
      Top             =   765
      Width           =   1485
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgRt 
         Height          =   810
         Left            =   810
         TabIndex        =   16
         Top             =   735
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1429
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin DebaFrmCtl.DLbl dLblSaveAs 
         Height          =   300
         Left            =   3795
         TabIndex        =   11
         Top             =   4455
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         BackColor       =   16777215
         Caption         =   "Save As"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnMouseMoveForeColor=   16744576
         Picture         =   "frmMain.frx":0000
      End
      Begin RichTextLib.RichTextBox rt 
         Height          =   3390
         Left            =   -150
         TabIndex        =   8
         Top             =   1830
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   5980
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":02BE
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
      Begin VB.Label lblProcedure 
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
         Left            =   210
         TabIndex        =   12
         Top             =   285
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   2250
      ScaleHeight     =   2250
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   30
      Width           =   1455
      Begin DebaFrmCtl.DLbl dLblExprtToExl 
         Height          =   300
         Left            =   4650
         TabIndex        =   10
         Top             =   4305
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
         OnMouseMoveForeColor=   32768
         Picture         =   "frmMain.frx":0339
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
         Height          =   2790
         Left            =   75
         TabIndex        =   6
         Top             =   1380
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   4921
         _Version        =   393216
         BackColorFixed  =   14408667
         ForeColorFixed  =   10485760
         BackColorBkg    =   15658734
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblTable 
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
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   750
      Top             =   1170
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
            Picture         =   "frmMain.frx":05F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2485
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":315F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A39
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4713
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":727B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B55
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3690
      Left            =   210
      TabIndex        =   4
      Top             =   15
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   6509
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
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   4830
      TabIndex        =   3
      Top             =   165
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   4845
      TabIndex        =   2
      Top             =   45
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   5370
      TabIndex        =   1
      Top             =   30
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   3885
      TabIndex        =   0
      Top             =   15
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCreate 
         Caption         =   "Create"
         Begin VB.Menu mnuCreatetable 
            Caption         =   "Create Table"
         End
         Begin VB.Menu mnuB1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCreateView 
            Caption         =   "Create View"
         End
         Begin VB.Menu mnuB2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCreateProcedure 
            Caption         =   "Create Procedure"
         End
         Begin VB.Menu mnuB3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuChangeDatabase 
            Caption         =   "Change Database"
         End
      End
      Begin VB.Menu mnuDescription 
         Caption         =   "Description"
         Begin VB.Menu mnuTableDescription 
            Caption         =   "Table Description"
         End
         Begin VB.Menu mnuB4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewDescription 
            Caption         =   "View Description"
         End
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browse"
         Begin VB.Menu mnuBrowseData 
            Caption         =   "Browse Data"
         End
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Begin VB.Menu mnuAboutProject 
            Caption         =   "About Project"
         End
         Begin VB.Menu mnuB5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExit 
            Caption         =   "Exit"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh

Option Explicit
Dim p As New clsPicturePaint 'Paint Picture Box.
Dim pic1 As Boolean
Dim pic2 As Boolean
Dim rs  As New ADODB.Recordset
Dim prs As New ADODB.Recordset
Dim n As Node

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

Private Sub Form_Load()
    Set p = New clsPicturePaint
    p.PictureIni Picture1, 0, 100, "Table Or View"
    p.PictureIni Picture2, 100, 200, "Procedure"
     If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
    Call Picture1_Click
    
    modRtColor.InitWords
    modRtColor.DoColor rt
    rt.SelStart = 0
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
    Call FormToolTp
End Sub
Private Sub Form_Resize()
On Error Resume Next
    If Me.Width < 9000 Then
        Me.Width = 9000
    End If
    If Me.Height < 6700 Then
        Me.Height = 6700
    End If
    tv.Move 16, 40, 190, Me.ScaleHeight - 120
    cmdSearch.Move tv.Left, tv.Top + tv.Height + 5, tv.Width / 2 - 2
    cmdSave.Move cmdSearch.Left + cmdSearch.Width + 3, cmdSearch.Top, cmdSearch.Width
    Picture1.Move tv.Left + tv.Width + 8, tv.Top, Me.ScaleWidth - tv.Width - 36, Me.ScaleHeight - 90
    lblTable.Move 10, 45
    Fg.Move 10, 60, Picture1.ScaleWidth - 25, Picture1.Height - dLblExprtToExl.Height - 70
    dLblExprtToExl.Move Fg.Left + Fg.Width - dLblExprtToExl.Width - 2, Fg.Top + Fg.Height + 4
    Picture2.Move Picture1.Left, Picture1.Top, Picture1.ScaleWidth, Picture1.ScaleHeight
    lblMenu.Move Me.ScaleWidth - lblMenu.Width - 20, 40
    If pic1 = True Then
        Call Picture1_Click
    End If
    If pic2 = True Then
        Call Picture2_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set p = Nothing
End Sub
Private Sub lblMenu_Click()
    PopupMenu mnuMain
End Sub

Private Sub mnuAboutProject_Click()
    'frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBrowseData_Click()
    Unload Me
    frmBrowseData.Show
End Sub

Private Sub mnuChangeDatabase_Click()
    Unload Me
    frmLogOn.Show
End Sub

Private Sub mnuCreateProcedure_Click()
    gDmsg.DebMsgBox "Available in my next submission on Planet Source Code", "ORACLE VB Project", DebmsgInformation
End Sub

Private Sub mnuCreatetable_Click()
    Unload Me
    frmCreateTable.Show
End Sub

Private Sub mnuCreateView_Click()
    gDmsg.DebMsgBox "Available in my next submission on Planet Source Code", "ORACLE VB Project", DebmsgInformation
End Sub

Private Sub mnuEditRecord_Click()
    Unload Me
    frmEditRecordset.Show
End Sub

Private Sub mnuExit_Click()
    'Set gDmsg = Nothing
    Unload Me
End Sub

Private Sub mnuTableDescription_Click()
    Unload Me
    frmTableDesc.Show
End Sub
Private Sub mnuViewDescription_Click()
    Unload Me
    frmViewDescription.Show
End Sub

Private Sub Picture1_Click()
    pic1 = True
    pic2 = False
    p.PictureClick Picture1, 0, 100, "Table Or View"
    p.PictureIni Picture2, 100, 200, "Procedure"
End Sub
Private Sub Picture2_Click()
    pic1 = False
    pic2 = True
    p.PictureIni Picture1, 0, 100, "Table Or View"
    p.PictureClick Picture2, 100, 200, "Procedure"
    lblProcedure.Move 10, 45
    rt.Move 10, 60, Fg.Width, Fg.Height
    dLblSaveAs.Move rt.Left + rt.Width - dLblSaveAs.Width - 2, dLblExprtToExl.Top
End Sub
Sub TreeViewLoad()
Screen.MousePointer = vbHourglass
On Error GoTo rsError
    uid = Trim$(UCase(uid))
    Fg.ColWidth(0) = 400
    Set n = tv.Nodes.Add(, , "ORACLE", " " & UCase(uid) & " ", 1, 2)
    n.Expanded = True

    'Load Table In Treeview
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Table", "Table", 5, 6)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & uid & "", Empty, "Table"))
        Do Until rs.EOF
            Set n = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
            rs.MoveNext
        Loop
        
    'Load View In Treeview
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "View", "View", 7, 8)
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & uid & "", Empty, "View"))
        Do Until rs.EOF
            Set n = tv.Nodes.Add("View", tvwChild, "VV" & rs!TABLE_NAME, rs!TABLE_NAME, 3, 4)
            rs.MoveNext
        Loop
        
    'Load Sequence
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Sequence", "Sequence", 9, 10)
    Set rs = New ADODB.Recordset
        rs.Open "Select * from ALL_SEQUENCES where SEQUENCE_OWNER='" & uid & "'", db, adOpenDynamic, adLockBatchOptimistic
    Do Until rs.EOF
        Set n = tv.Nodes.Add("Sequence", tvwChild, "SQ" & rs!SEQUENCE_NAME, rs!SEQUENCE_NAME, 3, 4)
        rs.MoveNext
    Loop
    
    'Load Synonyms
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Synonyms", "Synonyms", 11, 12)
    Set rs = New ADODB.Recordset
    rs.Open "Select * from ALL_SYNONYMS Where OWNER='" & uid & "'", db, adOpenDynamic, adLockBatchOptimistic
    Do Until rs.EOF
        Set n = tv.Nodes.Add("Synonyms", tvwChild, "SS" & rs!SYNONYM_NAME, rs!SYNONYM_NAME, 3, 4)
        rs.MoveNext
    Loop
    
    'Load Procedure
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Procedure", "Procedure", 5, 6)
    Set rs = db.OpenSchema(adSchemaProcedures, Array(Empty, "" & uid & "", Empty, Empty))
    Do Until rs.EOF
        Set n = tv.Nodes.Add("Procedure", tvwChild, "PP" & rs!PROCEDURE_NAME, rs!PROCEDURE_NAME, 3, 4)
        rs.MoveNext
    Loop
    
    'Load Package
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Package", "Package", 7, 8)
    Set rs = New ADODB.Recordset
    rs.Open "Select * from ALL_SOURCE Where OWNER='" & uid & "' AND Type='PACKAGE'", db, adOpenDynamic, adLockBatchOptimistic
        Do Until rs.EOF
            On Error Resume Next
            Set n = tv.Nodes.Add("Package", tvwChild, "PK" & rs!Name, rs!Name, 3, 4)
            rs.MoveNext
            On Error Resume Next
        Loop
    
    'Load Package Body
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "PackageBody", "Package Body", 9, 10)
    Set rs = New ADODB.Recordset
    rs.Open "Select * From ALL_SOURCE Where OWNER='" & uid & "' AND Type='PACKAGE BODY'", db, adOpenDynamic, adLockBatchOptimistic
        Do Until rs.EOF
            On Error Resume Next
            Set n = tv.Nodes.Add("PackageBody", tvwChild, "PKD" & rs!Name, rs!Name, 3, 4)
            rs.MoveNext
            On Error Resume Next
        Loop
    
    'Load Type
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Type", "Type", 11, 12)
    Set rs = New ADODB.Recordset
        rs.Open "Select * From ALL_TYPES where OWNER='" & uid & "' ", db, adOpenDynamic, adLockBatchOptimistic
    Do Until rs.EOF
        Set n = tv.Nodes.Add("Type", tvwChild, "TP" & rs!TYPE_NAME, rs!TYPE_NAME, 3, 4)
        rs.MoveNext
    Loop
    
    If rs.State = 1 Then
        rs.Close
    End If
    
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error Occured While Processing : " & Err.Description & " ", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Sub FlexGridRowNo()
    Dim i
    For i = 1 To Fg.Rows - 1
        Fg.TextMatrix(i, 0) = i
    Next i
End Sub

Private Sub rt_Change()
    Dim l As Long
    l = rt.SelStart
    modRtColor.DoColor rt
    rt.SelStart = l
    rt.SelColor = vbBlack
End Sub

Private Sub Timer1_Timer()
    Call TreeViewLoad
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
            Call FlexGridRowNo
            lblTable.Caption = "Table Name:" & Node.Text & " "
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "VV" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            Call FlexGridRowNo
            lblTable.Caption = "View Name: " & Node.Text & " "
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "SQ" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            Call FlexGridRowNo
            lblTable.Caption = "Sequence Name: " & Node.Text & " "
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "SS" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set Fg.DataSource = rs
            Call FlexGridRowNo
            lblTable.Caption = "Synonym Name : " & Node.Text & " "
            tvSelectedItem.Text = ""
            tvSelectedItem.Text = Node.Text
            Call Picture1_Click
            
    ElseIf Node.Key = "PP" & Node.Text Then
            Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & uid & "' And Type = 'PROCEDURE' And NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rt.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rt.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProcedure.Caption = "Procedure Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "PK" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & uid & "' And Type = 'PACKAGE' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rt.SelColor = vbBlack
            rt.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rt.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProcedure.Caption = "Package Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "PKD" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & uid & "' And Type = 'PACKAGE BODY' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rt.SelColor = vbBlack
            rt.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rt.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProcedure.Caption = "Package Body Name : " & Node.Text & " "
            Call Picture2_Click
            
    ElseIf Node.Key = "TP" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_TYPE_VERSIONS where OWNER='" & uid & "' And TYPE_NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
            rt.SelColor = vbBlack
            rt.Text = ""
                For i = 1 To fgRt.Rows - 1
                    rt.SelText = fgRt.TextMatrix(i, 1)
                Next i
            lblProcedure.Caption = "Type Name : " & Node.Text & " "
            Call Picture2_Click
    
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox "Error : " & Err.Description & " ", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub
Private Sub FormToolTp()
    modToolTip.CreateBalloon lblMenu, lblMenu.hwnd, "Click Here For Menu", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdSearch, cmdSearch.hwnd, "Search Data", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdSave, cmdSave.hwnd, "Save Data Of Treeview", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblExprtToExl, dLblExprtToExl.hwnd, "Export Grid Data To Microsoft Excel", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblSaveAs, dLblSaveAs.hwnd, "Save As Text", szBalloon, False, Me.Caption, etiInfo
End Sub
