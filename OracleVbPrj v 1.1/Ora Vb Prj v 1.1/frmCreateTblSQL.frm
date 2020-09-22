VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateTblSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Table"
   ClientHeight    =   6615
   ClientLeft      =   960
   ClientTop       =   1395
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   5595
      TabIndex        =   7
      Top             =   315
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   6180
      TabIndex        =   6
      Top             =   570
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   6030
      TabIndex        =   5
      Top             =   495
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   4530
      TabIndex        =   4
      Top             =   375
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4500
      Top             =   6675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DebaFrmCtl.DGrad cmdExecute 
      Height          =   315
      Left            =   8430
      TabIndex        =   3
      Top             =   5800
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      OnMouseMoveGradient=   1
      Caption         =   "Execute"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCreateTblSQL.frx":0000
      MousePointer    =   99
      ScaleHeight     =   21
      ScaleMode       =   3
      ScaleWidth      =   82
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdSaveas 
      Height          =   330
      Left            =   1665
      TabIndex        =   2
      Top             =   5800
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      DefaultGradient =   3
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
      MouseIcon       =   "frmCreateTblSQL.frx":0162
      MousePointer    =   99
      ScaleHeight     =   22
      ScaleMode       =   3
      ScaleWidth      =   94
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdReload 
      Height          =   315
      Left            =   315
      TabIndex        =   1
      Top             =   5800
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      DefaultGradient =   3
      Caption         =   "Reload"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCreateTblSQL.frx":02C4
      MousePointer    =   99
      ScaleHeight     =   21
      ScaleMode       =   3
      ScaleWidth      =   88
   End
   Begin RichTextLib.RichTextBox rtSQL 
      Height          =   4740
      Left            =   315
      TabIndex        =   0
      Top             =   990
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   8361
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmCreateTblSQL.frx":0426
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL (no need to put semicolon"";"" at the end of statement)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   345
      TabIndex        =   8
      Top             =   660
      Width           =   4965
   End
End
Attribute VB_Name = "frmCreateTblSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh


Option Explicit
Dim tblRs  As New ADODB.Recordset
Private Sub cmdExecute_Click()
On Error GoTo rsError
    Set tblRs = New ADODB.Recordset
    tblRs.Open "" & rtSQL.Text & "", db, adOpenDynamic, adLockBatchOptimistic
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error ( Creating Table ), Please Check The SQL Again " & Err.Description & " ", "Error", DebmsgCritical
    Else
        gDmsg.DebMsgBox "" & frmCreateTable.txtTableName.Text & " Table Created Successfully", "Successful", DebmsgInformation
    End If
End Sub
Private Sub cmdReload_Click()
    rtSQL.Text = ""
    rtSQL.Text = frmCreateTable.rtSQL.Text
End Sub
Private Sub cmdSaveAs_Click()
    If rtSQL.Text <> "" Then
        modRtSaveAs.SaveTextAs rtSQL, cd
    End If
End Sub
Private Sub Form_Load()
    modRtColor.InitWords
    modRtColor.DoColor rtSQL
    rtSQL.SelStart = 0
    rtSQL.Text = frmCreateTable.rtSQL.Text
    Call FormToolTip
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo rsError
    'Checking State Whether Is It Open Or Close
    If tblRs.State = 1 Then
        tblRs.Close
    End If
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description
        Exit Sub
    End If
End Sub
Private Sub rtSQL_Change()
    Dim lCursor As Long
    lCursor = rtSQL.SelStart
    modRtColor.DoColor rtSQL
    rtSQL.SelStart = lCursor
    rtSQL.SelColor = vbBlack
End Sub
Sub FormToolTip()
    modToolTip.CreateBalloon cmdExecute, cmdExecute.hwnd, "Execute", szBalloon, True
    modToolTip.CreateBalloon cmdSaveas, cmdSaveas.hwnd, "Save As (Text File, SQL File, Rich Text File)", szBalloon, False, Me.Caption, etiInfo
    modToolTip.CreateBalloon cmdReload, cmdReload.hwnd, "Reload SQL Again", szBalloon, False, Me.Caption, etiInfo
End Sub

Private Sub rtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rtTabPress As Boolean
    rtTabPress = (KeyCode = vbKeyTab)
    If rtTabPress Then
        rtSQL.SelText = vbTab
        KeyCode = 0
    End If
End Sub

