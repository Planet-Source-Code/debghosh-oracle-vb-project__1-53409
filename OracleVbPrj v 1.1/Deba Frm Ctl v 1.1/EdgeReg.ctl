VERSION 5.00
Begin VB.UserControl EdgeReg 
   BackColor       =   &H00C00000&
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   MousePointer    =   8  'Size NW SE
   ScaleHeight     =   195
   ScaleWidth      =   195
   ToolboxBitmap   =   "EdgeReg.ctx":0000
End
Attribute VB_Name = "EdgeReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
Dim hr1, hr2, hr3 As Long
Dim Bx, Gx, i As Integer
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Dim Rs As Boolean
Dim aX As Integer, aY As Integer

Private Sub UCpaint()
    With UserControl
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Cls
    End With
    hr1 = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    hr2 = CreateRoundRectRgn(-10, -10, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
    hr3 = CreateRoundRectRgn(-10, -10, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6, 20, 20)
    CombineRgn hr1, hr2, hr3, RGN_XOR
    SetWindowRgn UserControl.hwnd, hr1, True
    UserControl.Refresh
    'Call BackCol
End Sub
Private Sub BackCol()
    UserControl.BackColor = RGB(0, 0, 160)
End Sub

Private Sub frm_Load()
    frm.AutoRedraw = True
    frm.ScaleMode = vbPixels
    MoveWindow UserControl.hwnd, frm.ScaleWidth - UserControl.ScaleWidth, frm.ScaleHeight - UserControl.ScaleHeight + 2, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
    If frm.BorderStyle = 2 Or frm.BorderStyle = 5 Then
        UserControl.Enabled = True
    Else
        UserControl.Enabled = False
    End If
End Sub

Private Sub frm_Resize()
    MoveWindow UserControl.hwnd, frm.ScaleWidth - UserControl.ScaleWidth, frm.ScaleHeight - UserControl.ScaleHeight + 2, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
End Sub

Private Sub frm_Unload(Cancel As Integer)
    DeleteObject hr1
    DeleteObject hr2
    DeleteObject hr3
End Sub
Private Sub UserControl_Initialize()
    Call UCpaint
End Sub
Private Sub UserControl_Resize()
    Call UCpaint
    UserControl.Width = 300
    UserControl.Height = 275
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo picError
    Dim Result As Long
    Dim Pos As POINTAPI
    Dim c As Long
        Rs = True
        Do
        Result = GetCursorPos(Pos)
        aX% = Pos.X
        aY% = Pos.Y
        DoEvents
        Result = GetCursorPos(Pos)
        frm.Width = frm.Width + (Pos.X - aX%) * 20
        frm.Height = frm.Height + (Pos.Y - aY%) * 20
        'Pic.Left = frm.ScaleWidth - Pic.ScaleWidth
        'Pic.Top = frm.ScaleHeight - Pic.ScaleHeight
        Loop Until Rs = False
        Exit Sub
picError:
        Rs = False
        Exit Sub
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rs = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
End Sub
