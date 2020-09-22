VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditRecHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   6555
   ClientLeft      =   1995
   ClientTop       =   1545
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin DebaFrmCtl.DGrad DGrad1 
      Height          =   345
      Left            =   5550
      TabIndex        =   5
      Top             =   5505
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      OnMouseMoveGradient=   1
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmEditRecHelp.frx":0000
      MousePointer    =   99
      ScaleHeight     =   23
      ScaleMode       =   3
      ScaleWidth      =   71
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4425
      Left            =   495
      TabIndex        =   4
      Top             =   990
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   7805
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      FileName        =   "G:\MyPSCSubmitCode\OracleVbPrj v 1.1\Icon\EditHelp.rtf"
      TextRTF         =   $"frmEditRecHelp.frx":0162
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
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   5865
      TabIndex        =   3
      Top             =   60
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   150
      Left            =   6180
      TabIndex        =   2
      Top             =   180
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   265
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   6255
      TabIndex        =   1
      Top             =   45
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   6345
      TabIndex        =   0
      Top             =   15
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
End
Attribute VB_Name = "frmEditRecHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DGrad1_Click()
    Unload Me
End Sub
