VERSION 5.00
Object = "*\A..\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Begin VB.Form frmLogOn 
   BorderStyle     =   0  'None
   Caption         =   "LogOn"
   ClientHeight    =   4620
   ClientLeft      =   2220
   ClientTop       =   2400
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   4425
      TabIndex        =   12
      Top             =   30
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   5625
      TabIndex        =   11
      Top             =   15
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   150
      Left            =   5505
      TabIndex        =   10
      Top             =   90
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   265
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   210
      Left            =   5325
      TabIndex        =   9
      Top             =   45
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   370
   End
   Begin DebaFrmCtl.DGrad cmdCancel 
      Height          =   360
      Left            =   5070
      TabIndex        =   4
      Top             =   3345
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      OnMouseMoveGradient=   1
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   24
      ScaleMode       =   0
      ScaleWidth      =   90
      OnMouseMoveForeColor=   12648447
   End
   Begin DebaFrmCtl.DGrad cmdOK 
      Height          =   360
      Left            =   3690
      TabIndex        =   3
      Top             =   3345
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      OnMouseMoveGradient=   3
      Caption         =   "&Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   24
      ScaleMode       =   0
      ScaleWidth      =   90
      OnMouseMoveForeColor=   12648447
   End
   Begin VB.TextBox txtDb 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2000
      TabIndex        =   2
      Top             =   2520
      Width           =   4440
   End
   Begin VB.TextBox txtPwd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1995
      Width           =   4440
   End
   Begin VB.TextBox txtUserId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2000
      TabIndex        =   0
      Top             =   1485
      Width           =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   285
      X2              =   6600
      Y1              =   3210
      Y2              =   3195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   390
      TabIndex        =   8
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   375
      TabIndex        =   7
      Top             =   2085
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   1575
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter USER ID, PASSWORD and DATABASE."
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
      Height          =   360
      Left            =   555
      TabIndex        =   5
      Top             =   780
      Width           =   5790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   255
      X2              =   6585
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3270
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "frmLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh

Option Explicit
Dim t As New clsTextSubclass ' Subclassing Text Box
Private Sub cmdCancel_Click()
On Error GoTo eError
    Unload Me
eError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        End
    End If
End Sub
Private Sub cmdOk_Click()
    If txtUserId.Text = "" Then
        txtUserId.SetFocus
        gDmsg.DebMsgBox "Please Enter USER ID", "Error", DebmsgExclamation
        Exit Sub
    Else
        If txtPwd.Text = "" Then
            txtPwd.SetFocus
            gDmsg.DebMsgBox "Please Enter Password", "Error", DebmsgExclamation
            Exit Sub
        End If
    End If
    Call modLogOn.OracleConnect
End Sub
Private Sub Form_Load()
    Set t = New clsTextSubclass
    t.TxtPop txtUserId
    t.TxtPop txtPwd
    t.TxtPop txtDb
    Call FormToolTip
End Sub
Sub FormToolTip()
    txtUserId.TabIndex = 0
    modToolTip.CreateBalloon txtUserId, txtUserId.hwnd, "Enter USER ID Here", szBalloon, False, Me.Caption, etiInfo
    txtPwd.TabIndex = 1
    modToolTip.CreateBalloon txtPwd, txtPwd.hwnd, "Enter Password Here", szBalloon, False, Me.Caption, etiInfo
    txtDb.TabIndex = 2
    modToolTip.CreateBalloon txtDb, txtDb.hwnd, "Enter Database Name Here", szBalloon, False, Me.Caption, etiInfo
    cmdOK.TabIndex = 3
    modToolTip.CreateBalloon cmdOK, cmdOK.hwnd, "Click Here To Connect ORACLE", szBalloon, False, Me.Caption, etiInfo
    cmdCancel.TabIndex = 4
    modToolTip.CreateBalloon cmdCancel, cmdCancel.hwnd, "Exit Project", szBalloon, False, Me.Caption, etiError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set t = Nothing
End Sub
