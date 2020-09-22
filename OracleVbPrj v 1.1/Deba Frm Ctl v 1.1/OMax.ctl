VERSION 5.00
Begin VB.UserControl OMax 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   MouseIcon       =   "OMax.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   570
   ScaleWidth      =   675
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2200
      Top             =   45
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   255
      Width           =   315
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   480
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   525
   End
End
Attribute VB_Name = "OMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Form Maximize And Restore Down Control
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Public frmState As Boolean
Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.X, p.Y) = UserControl.hwnd Then
         MouseOver = True
    End If
End Function
Private Sub Timer1_Timer()
    If Not MouseOver Then
    If UserControl.Enabled = True Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbGreen
            End With
        End If
        Else
            With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = &HE0E0E0
        End With
        Shape3.Visible = False
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = &HC0C0C0
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    End If
        Timer1.Enabled = False
    End If
End Sub
Public Sub UcState()
    If frmState = True Then
        Shape2.Visible = True
        Shape3.Visible = True
        Shape2.Move 2, 2, 10, 9
        Shape3.Move 6, 6, Shape2.Width, Shape2.Height
    Else
        Shape3.Visible = False
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call UCpaint
End Sub
Private Sub UCpaint()
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Cls
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If UserControl.Enabled = True Then
        Shape3.Visible = False
        With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = vbWhite
        End With
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = vbGreen
        End With
        With Shape3
            .ZOrder 0
            .FillStyle = 0
            .FillColor = vbGreen
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    Else
        If UserControl.Enabled = False Then
        With Shape1
            .ZOrder 1
            .FillStyle = 0
            .FillColor = &HE0E0E0
        End With
        Shape3.Visible = False
        With Shape2
            .ZOrder 0
            .FillStyle = 0
            .FillColor = &HC0C0C0
        End With
        Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = &HE0E0E0
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = &HE0E0E0
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = &H808080
            End With
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X > 0 Or X < UserControl.ScaleWidth Or Y > 0 Or Y < UserControl.ScaleHeight Then
        If frmState = True Then
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
            With Shape3
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
        Else
            With Shape1
                .ZOrder 1
                .FillStyle = 0
                .FillColor = vbGreen
            End With
            With Shape2
                .Visible = True
                .ZOrder 0
                .FillStyle = 0
                .FillColor = vbWhite
            End With
        End If
    Else
        Timer1.Enabled = False
        Call UCpaint
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If frmState = True Then
            frmState = False
            Shape2.Visible = True
            Shape3.Visible = True
            Shape2.Move 2, 2, 10, 9
            Shape3.Move 6, 6, Shape2.Width, Shape2.Height
        Else
            frmState = True
            Shape3.Visible = False
            Shape2.Move 3, 3, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
        End If
        RaiseEvent Click
    End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 255
    UserControl.Width = 270
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call UCpaint
    PropertyChanged "Enabled"
End Property

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
