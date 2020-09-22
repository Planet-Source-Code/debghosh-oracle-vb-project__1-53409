VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCreateTable 
   Caption         =   "Create Table"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5580
      Left            =   75
      ScaleHeight     =   5550
      ScaleWidth      =   7260
      TabIndex        =   9
      Top             =   1125
      Width           =   7290
      Begin DebaFrmCtl.DLbl dLblHelp 
         Height          =   240
         Left            =   330
         TabIndex        =   18
         Top             =   4920
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   423
         BackColor       =   16777215
         Caption         =   "HELP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnMouseMoveForeColor=   16711935
         Picture         =   "frmCreateTable.frx":0000
      End
      Begin DebaFrmCtl.DLbl dLblShowSQL 
         Height          =   195
         Left            =   1380
         TabIndex        =   17
         Top             =   4890
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   344
         BackColor       =   16777215
         Caption         =   "Show SQL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnMouseMoveForeColor=   16761024
         PictureVisible  =   0   'False
      End
      Begin RichTextLib.RichTextBox rtCond 
         Height          =   975
         Left            =   255
         TabIndex        =   16
         Top             =   3825
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   1720
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmCreateTable.frx":3082
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
      Begin VB.TextBox txtSQL 
         Height          =   300
         Left            =   3930
         TabIndex        =   15
         Top             =   5115
         Visible         =   0   'False
         Width           =   2250
      End
      Begin RichTextLib.RichTextBox rtSQL 
         Height          =   1800
         Left            =   3930
         TabIndex        =   14
         Top             =   3210
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   3175
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmCreateTable.frx":30FE
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
      Begin VB.TextBox txtFgCell 
         Height          =   360
         Left            =   3915
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cmbCol 
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
         Left            =   3930
         TabIndex        =   12
         Top             =   2445
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3930
         TabIndex        =   11
         Top             =   2115
         Visible         =   0   'False
         Width           =   2100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Fg 
         Height          =   2910
         Left            =   270
         TabIndex        =   10
         Top             =   885
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   5133
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
      Begin VB.Label lblExtraCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Write Extra Code Here"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   19
         Top             =   1755
         Width           =   1635
      End
   End
   Begin DebaFrmCtl.DGrad cmdLoad 
      Height          =   360
      Left            =   2805
      TabIndex        =   8
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      OnMouseMoveGradient=   3
      Caption         =   "&Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCreateTable.frx":317A
      MousePointer    =   99
      ScaleHeight     =   24
      ScaleMode       =   0
      ScaleWidth      =   81
   End
   Begin VB.ComboBox cmbColNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   7
      Top             =   645
      Width           =   1605
   End
   Begin VB.TextBox txtTableName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   990
      TabIndex        =   5
      Top             =   90
      Width           =   3240
   End
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   7185
      TabIndex        =   3
      Top             =   525
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   7020
      TabIndex        =   2
      Top             =   315
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   7365
      TabIndex        =   1
      Top             =   345
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin VB.Label lblSelectCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Column Nos."
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
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   630
      Width           =   930
   End
   Begin VB.Label lblTableName 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Table Name"
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
      Height          =   420
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   870
   End
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh


Option Explicit
Dim cmb As New clsCombo
Dim pp As New clsPicturePaint
Dim l_Click As Boolean
Private Sub dLblHelp_Click()
    frmCreateTblHelp.Show vbModal, Me
End Sub
Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    Set pp = New clsPicturePaint
    Picture1.ScaleMode = vbPixels
    pp.PictureIni Picture1, 0, 150, "Create Table"
    Set cmb = New clsCombo
    Set cmb.Combo1 = cmbCol
    Call modCreateTable.OnFormLoad
    dLblShowSQL.Enabled = False
    l_Click = False
    Call FormToolTip
    OFrmCtl1.SysTrayInfo = "User Name :    " & uid & " " & vbCrLf _
                        & "Database :    " & d_Base & " "
End Sub
Private Sub cmbCol_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode fg, cmbCol, KeyCode, Shift
End Sub
Private Sub cmbCol_LostFocus()
    cmbCol.Text = UCase(cmbCol.Text)
End Sub
Private Sub cmdLoad_Click()
    l_Click = True
    Call modCreateTable.OnLoadClick
    dLblShowSQL.Enabled = True
End Sub
Private Sub dLblShowSQL_Click()
    If txtTableName.Text <> "" Then
        rtSQL.Text = ""
        modCreateTable.TextSQL
        If rtSQL.Text <> "" Then
            frmCreateTblSQL.Show vbModal, Me
        End If
    End If
End Sub
Private Sub Fg_Click()
    If l_Click = True Then
    If fg.Col = 1 Then
        MSHFlexGridEdit fg, txtEdit, 32
    Else
        Call modCreateTable.ColData(fg.Col)
        MSHFlexGridEdit fg, cmbCol, 32
    End If
    Else
        txtTableName.SetFocus
        gDmsg.DebMsgBox "Please Enter Table Name And Select Column Name then Click On Load", "Error", DebmsgExclamation
    End If
End Sub
Private Sub Fg_GotFocus()
    If fg.Col = 1 Then
        If txtEdit.Visible = False Then Exit Sub
            fg = txtEdit
            txtEdit.Visible = False
    Else
        If cmbCol.Visible = False Then Exit Sub
            fg = cmbCol
            cmbCol.Visible = False
    End If
End Sub
Private Sub Fg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call Fg_Click
    End If
End Sub
Private Sub Fg_LeaveCell()
    If fg.Col = 1 Then
        If txtEdit.Visible = False Then Exit Sub
            fg = txtEdit
            txtEdit.Visible = False
    Else
        If cmbCol.Visible = False Then Exit Sub
            fg = cmbCol
            cmbCol.Visible = False
    End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Me.ScaleMode = vbPixels
    If Me.Width < 12000 Then
        Me.Width = 12000
    End If
    If Me.Height < 7500 Then
        Me.Height = 7500
    End If
    lblTableName.Move 180, 45
    txtTableName.Move lblTableName.Left + lblTableName.Width + 5, 43
    lblSelectCol.Move txtTableName.Left + txtTableName.Width + 15, 43
    cmbColNo.Move lblSelectCol.Left + lblSelectCol.Width + 5, 43
    cmdLoad.Move cmbColNo.Left + cmbColNo.Width + 15, 43
    Picture1.Move 16, 43, Me.ScaleWidth - 36, Me.ScaleHeight - 90
    pp.PictureClick Picture1, 0, 150, "Create Table"
    fg.Move 10, 60, Picture1.ScaleWidth - 36, Picture1.ScaleHeight - 200
    lblExtraCode.Move 10, fg.Top + fg.Height + 5
    rtCond.Move 10, lblExtraCode.Top + lblExtraCode.Height + 5, fg.Width, Picture1.ScaleHeight - fg.Top - fg.Height - 60
    dLblHelp.Move 10, rtCond.Top + rtCond.Height + 8
    dLblShowSQL.Move fg.Left + fg.Width - dLblShowSQL.Width - 15, dLblHelp.Top
    Picture1.ScaleMode = vbTwips
    If d_Base <> "" Then
        OFrmCtl1.StatusCaption = " " & uid & " /  " & d_Base & " ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    Else
        OFrmCtl1.StatusCaption = " " & uid & " /  ???? ---- " & Format(Now, "dddd,  dd MMM YYYY") & " "
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set cmb = Nothing
    Set pp = Nothing
    frmMain.Show
End Sub

Private Sub rtCond_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rtTabPress As Boolean
    rtTabPress = (KeyCode = vbKeyTab)
    If rtTabPress Then
        rtCond.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode fg, txtEdit, KeyCode, Shift
End Sub
Private Sub txtTableName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 35, 36, 48 To 57, 65 To 90, 95, 97 To 122
        Case Else
            gDmsg.DebMsgBox "Please Enter A -- Z,a -- z,0 -- 9, _ ,$,# Or Click On Help", "Error", DebmsgCritical
            KeyAscii = 0
    End Select
End Sub
Sub FormToolTip()
    modToolTip.CreateBalloon txtTableName, txtTableName.hwnd, "Enter Table Name", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmbColNo, cmbColNo.hwnd, "Select No. Of Columns", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdLoad, cmdLoad.hwnd, "Click Here To Load Column(s) In Grid", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon txtEdit, txtEdit.hwnd, "Enter Column Name", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmbCol, cmbCol.hwnd, "Select Text from Drop Down List or Enter as per Your Requirement", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblHelp, dLblHelp.hwnd, "Click Here For Help", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon dLblShowSQL, dLblShowSQL.hwnd, "Click Here To Show SQL", szBalloon, False, Me.Caption, etiInfo
End Sub

