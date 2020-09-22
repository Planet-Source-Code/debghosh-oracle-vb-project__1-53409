VERSION 5.00
Begin VB.UserControl EdgeBottom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   MousePointer    =   7  'Size N S
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   19
   ToolboxBitmap   =   "EdgeBottom.ctx":0000
End
Attribute VB_Name = "EdgeBottom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Dim rs As Boolean
Dim aX As Integer, aY As Integer
Private Sub frm_Load()
    frm.AutoRedraw = True
    frm.ScaleMode = vbPixels
    MoveWindow UserControl.hwnd, 0, frm.ScaleHeight - 3, frm.ScaleWidth - 8, 3, 1
    If frm.BorderStyle = 2 Or frm.BorderStyle = 5 Then
        UserControl.Enabled = True
    Else
        UserControl.Enabled = False
    End If
End Sub
Private Sub frm_Resize()
    MoveWindow UserControl.hwnd, 0, frm.ScaleHeight - 3, frm.ScaleWidth - 8, 3, 1
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo picError
    Dim Result As Long
    Dim Pos As POINTAPI
    Dim c As Long
        rs = True
        Do
        Result = GetCursorPos(Pos)
        aX% = Pos.X
        aY% = Pos.Y
        DoEvents
        Result = GetCursorPos(Pos)
        'frm.Width = frm.Width + (Pos.X - aX%) * 20
        frm.Height = frm.Height + (Pos.Y - aY%) * 20
        'Pic.Left = frm.ScaleWidth - Pic.ScaleWidth
        'Pic.Top = frm.ScaleHeight - Pic.ScaleHeight
        Loop Until rs = False
        Exit Sub
picError:
        rs = False
        Exit Sub
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    rs = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
End Sub
