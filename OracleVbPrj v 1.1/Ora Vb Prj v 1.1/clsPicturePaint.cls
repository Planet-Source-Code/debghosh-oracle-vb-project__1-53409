VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicturePaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh

'Using A Picture Box  like a Tab Control.
Option Explicit
'Initializing Picture Box
Public Sub PictureIni(pict As PictureBox, st As Integer, en As Integer, txt As String)
    Dim r As Long
    Dim rcU As Long
    Dim rcL As Long
    With pict
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .BorderStyle = 0
        .Cls
        .BackColor = vbGreen
        .ZOrder 1
    End With
    r = CreateRectRgn(0, 0, pict.ScaleWidth, pict.ScaleHeight)
    rcU = CreateRoundRectRgn(st, 0, en, 50, 12, 12)
    rcL = CreateRoundRectRgn(0, 36, pict.ScaleWidth, pict.ScaleHeight, 12, 12)
    CombineRgn r, rcU, rcL, RGN_OR
    pict.Font = "Verdana"
    pict.FontSize = 10
    TextOut pict.hdc, st + 4, 10, txt, Len(txt)
    SetWindowRgn pict.hwnd, r, True
    DeleteObject r
    DeleteObject rcU
    DeleteObject rcL
End Sub
Public Sub PictureClick(pict As PictureBox, st As Integer, en As Integer, txt As String)
    Dim r As Long
    Dim rcU As Long
    Dim rcL As Long
    With pict
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .BorderStyle = 0
        .Cls
        .BackColor = vbWhite
        .ZOrder 0
    End With
    r = CreateRectRgn(0, 0, pict.ScaleWidth, pict.ScaleHeight)
    rcU = CreateRoundRectRgn(st, 0, en, 50, 12, 12)
    rcL = CreateRoundRectRgn(0, 36, pict.ScaleWidth, pict.ScaleHeight, 12, 12)
    CombineRgn r, rcU, rcL, RGN_OR
    FrameRgn pict.hdc, r, CreateSolidBrush(&H80FF&), 2, 2
    pict.Font = "Verdana"
    pict.FontSize = 10
    TextOut pict.hdc, st + 4, 10, txt, Len(txt)
    SetWindowRgn pict.hwnd, r, True
    DeleteObject r
    DeleteObject rcU
    DeleteObject rcL
End Sub



