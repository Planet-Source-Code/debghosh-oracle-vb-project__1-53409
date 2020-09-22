VERSION 5.00
Object = "*\A..\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5715
   ClientLeft      =   2325
   ClientTop       =   1710
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin DebaFrmCtl.DGrad DGrad1 
      Height          =   330
      Left            =   4650
      TabIndex        =   5
      Top             =   5130
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
      DefaultGradient =   3
      Caption         =   "&Enter"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   22
      ScaleMode       =   0
      ScaleWidth      =   110
      OnMouseMoveForeColor=   12640511
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4230
      Left            =   300
      TabIndex        =   4
      Top             =   795
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7461
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      FileName        =   "G:\MyPSCSubmitCode\OracleVbPrj v 1.1\Icon\About.rtf"
      TextRTF         =   $"frmAbout.frx":0000
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
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   5385
      TabIndex        =   3
      Top             =   -210
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   15
      Left            =   4785
      TabIndex        =   2
      Top             =   -30
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   -26
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   30
      Left            =   5265
      TabIndex        =   1
      Top             =   -75
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   53
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DGrad1_Click()
    Unload Me
    frmLogOn.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogOn.Show
End Sub
