VERSION 5.00
Begin VB.Form frmProcess 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   1755
   ClientTop       =   2955
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7575
      Top             =   1845
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Process"
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
      Height          =   765
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   6180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ..................................."
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
      Height          =   390
      Left            =   765
      TabIndex        =   0
      Top             =   345
      Width           =   5490
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   405
      Left            =   6810
      Top             =   135
      Width           =   390
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Shape1.Move 0, 0, Me.Width, Me.Height
    lblText.Caption = "Retrieving Data.........."
    Timer1.Enabled = False
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
    Call modBrowseData.BrowseDataForm
    Timer1.Enabled = False
    Unload Me
End Sub


