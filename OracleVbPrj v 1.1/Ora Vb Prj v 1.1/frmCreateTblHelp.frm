VERSION 5.00
Object = "*\A..\..\..\MYPSCS~1\ORACLE~1.1\Deba Frm Ctl v 1.1\Deba Frm Ctl v1.1.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCreateTblHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   6615
   ClientLeft      =   1860
   ClientTop       =   1530
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DebaFrmCtl.DGrad DGrad1 
      Height          =   360
      Left            =   5550
      TabIndex        =   5
      Top             =   5775
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
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
      ForeColor       =   12648447
      MouseIcon       =   "frmCreateTblHelp.frx":0000
      MousePointer    =   99
      ScaleHeight     =   24
      ScaleMode       =   0
      ScaleWidth      =   85
      OnMouseMoveForeColor=   16777215
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   4875
      Left            =   570
      TabIndex        =   4
      Top             =   810
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   8599
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      FileName        =   "G:\MyPSCSubmitCode\OracleVbPrj v 1.1\Icon\CREATE TABLE.rtf"
      TextRTF         =   $"frmCreateTblHelp.frx":0162
   End
   Begin DebaFrmCtl.EdgeReg EdgeReg1 
      Height          =   270
      Left            =   6645
      TabIndex        =   3
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin DebaFrmCtl.EdgeBottom EdgeBottom1 
      Height          =   90
      Left            =   6990
      TabIndex        =   2
      Top             =   360
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   159
   End
   Begin DebaFrmCtl.EdgeRight EdgeRight1 
      Height          =   120
      Left            =   7305
      TabIndex        =   1
      Top             =   360
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   212
   End
   Begin DebaFrmCtl.OFrmCtl OFrmCtl1 
      Height          =   330
      Left            =   6525
      TabIndex        =   0
      Top             =   30
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      SysTrayInfo     =   ""
      StatusCaption   =   ""
   End
End
Attribute VB_Name = "frmCreateTblHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh



Private Sub DGrad1_Click()
    Unload Me
End Sub
