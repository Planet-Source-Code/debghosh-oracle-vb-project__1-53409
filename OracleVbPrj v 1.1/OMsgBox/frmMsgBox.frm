VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MsgBox"
   ClientHeight    =   4320
   ClientLeft      =   2775
   ClientTop       =   2160
   ClientWidth     =   5490
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OmsgBox.DClose DClose1 
      Height          =   255
      Left            =   4770
      TabIndex        =   10
      Top             =   1080
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1155
      Picture         =   "frmMsgBox.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3255
      Width           =   1000
   End
   Begin VB.CommandButton cmdRetry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Retry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1140
      Picture         =   "frmMsgBox.frx":306C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2835
      Width           =   1000
   End
   Begin VB.CommandButton cmdIgnore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ignore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   90
      Picture         =   "frmMsgBox.frx":5F8E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2835
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   75
      Picture         =   "frmMsgBox.frx":8EB0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3705
      Width           =   1000
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   90
      Picture         =   "frmMsgBox.frx":BDD2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3270
      Width           =   1000
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1125
      Picture         =   "frmMsgBox.frx":ECF4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2415
      Width           =   1000
   End
   Begin VB.CommandButton cmdAbort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Abort"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   90
      Picture         =   "frmMsgBox.frx":11C16
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2415
      Width           =   1000
   End
   Begin VB.Label lblHShade 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shade"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   1800
      TabIndex        =   9
      Top             =   -135
      Width           =   705
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3390
      TabIndex        =   8
      Top             =   1710
      Width           =   375
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Heading"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   435
      TabIndex        =   7
      Top             =   -120
      Width           =   645
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   1260
      Picture         =   "frmMsgBox.frx":14B38
      Top             =   3705
      Width           =   480
   End
   Begin VB.Image imgInformation 
      Height          =   480
      Left            =   1260
      Picture         =   "frmMsgBox.frx":1597A
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image imgExclamation 
      Height          =   480
      Left            =   1260
      Picture         =   "frmMsgBox.frx":167BC
      Top             =   3675
      Width           =   480
   End
   Begin VB.Image imgCritical 
      Height          =   480
      Left            =   1275
      Picture         =   "frmMsgBox.frx":175FE
      Top             =   3705
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DebMsgResult As Long
Private Sub cmdAbort_Click()
    DebMsgResult = DebmsgAbort
    Unload Me
End Sub

Private Sub cmdAbort_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdAbort.FontBold = True
End Sub

Private Sub cmdCancel_Click()
    DebMsgResult = 2
    Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdCancel.FontBold = True
End Sub

Private Sub cmdIgnore_Click()
    DebMsgResult = DebmsgIgnore
    Unload Me
End Sub

Private Sub cmdIgnore_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   cmdIgnore.FontBold = True
End Sub

Private Sub cmdNo_Click()
    DebMsgResult = 1
    Unload Me
End Sub

Private Sub cmdNo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdNo.FontBold = True
End Sub

Private Sub cmdOk_Click()
    DebMsgResult = 3
    Unload Me
End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdOk.FontBold = True
End Sub

Private Sub cmdRetry_Click()
    DebMsgResult = 4
    Unload Me
End Sub

Private Sub cmdRetry_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdRetry.FontBold = True
End Sub

Private Sub cmdYes_Click()
    DebMsgResult = 0
    Unload Me
End Sub

Private Sub cmdYes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdYes.FontBold = True
End Sub

Private Sub Form_Load()
    Beep
    DClose1.ToolTipText = "Close"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdAbort.FontBold = False
    cmdCancel.FontBold = False
    cmdIgnore.FontBold = False
    cmdNo.FontBold = False
    cmdOk.FontBold = False
    cmdRetry.FontBold = False
    cmdYes.FontBold = False
End Sub

Private Sub lblHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage frmMsgBox.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub lblHShade_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage frmMsgBox.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
