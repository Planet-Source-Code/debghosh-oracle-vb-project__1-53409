VERSION 5.00
Begin VB.UserControl DLbl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   MouseIcon       =   "DLbl.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   70
   ToolboxBitmap   =   "DLbl.ctx":0152
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   135
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debasis"
      Height          =   195
      Left            =   1110
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
   Begin VB.Image img 
      Height          =   60
      Left            =   1290
      Top             =   420
      Width           =   75
   End
End
Attribute VB_Name = "DLbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim m_d_Fc As OLE_COLOR
Dim m_Fc As OLE_COLOR
Dim img_Visible As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Private Sub UCpaint()
    Dim maxW, maxH As Integer
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    UserControl.Cls
    maxH = img.Height
    If maxH < lblCaption.Height Then
        maxH = lblCaption.Height
    End If
    If UserControl.Enabled = True Then
        If img_Visible = True Then
            lblCaption.Enabled = True
            img.Visible = True
            UserControl.Height = maxH * 15 'lblCaption.Height * 15
            UserControl.Width = (img.Width + lblCaption.Width + 4) * 15
            img.Move 0, (UserControl.ScaleHeight - img.Height) / 2, img.Width, img.Height
            lblCaption.Move img.Left + img.Width + 4, (UserControl.ScaleHeight - lblCaption.Height) / 2, lblCaption.Width, lblCaption.Height
        Else
            img.Visible = False
            lblCaption.Enabled = True
            lblCaption.Move 0, 0
            UserControl.Height = (lblCaption.Height * 15) 'lblCaption.Height * 15
            UserControl.Width = lblCaption.Width * 15
        End If
    Else
        lblCaption.Enabled = False
        img.Visible = False
        lblCaption.Move 0, 0
        UserControl.Height = maxH * 15 'lblCaption.Height * 15
        UserControl.Width = lblCaption.Width * 15 '(img.Width + lblCaption.Width + 4) * 15
    End If
End Sub
Private Function MouseOver() As Boolean
    Dim typPoint As POINTAPI
    Dim dumpAway As Long
    On Error Resume Next
    dumpAway = GetCursorPos(typPoint)
    If WindowFromPoint(typPoint.X, typPoint.Y) = UserControl.hwnd Then
        MouseOver = True
    End If
End Function
Private Sub DefLabelVal()
    lblCaption.ForeColor = m_d_Fc
    lblCaption.FontUnderline = False
End Sub
Private Sub LabelChange()
    lblCaption.ForeColor = m_Fc
    lblCaption.FontUnderline = True
End Sub
Private Sub img_Click()
    'If UserControl.Enabled = True Then
        'RaiseEvent Click
    'End If
End Sub
Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call UserControl_MouseDown(Button, Shift, X, Y)
    End If
End Sub
Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call UserControl_MouseUp(Button, Shift, X, Y)
    End If
End Sub
Private Sub lblCaption_Click()
    'If UserControl.Enabled = True Then
        'RaiseEvent Click
    'End If
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call UserControl_MouseDown(Button, Shift, X, Y)
    End If
End Sub
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call UserControl_MouseUp(Button, Shift, X, Y)
    End If
End Sub
Private Sub Timer1_Timer()
    If Not MouseOver Then
        Call UCpaint
        Call DefLabelVal
        Timer1.Enabled = False
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call UCpaint
    PropertyChanged "BackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    Call UCpaint
    PropertyChanged "Caption"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Call UCpaint
    PropertyChanged "Font"
End Property
Public Property Get PictureVisible() As Boolean
    PictureVisible = img_Visible
End Property
Public Property Let PictureVisible(ByVal New_Visible As Boolean)
    img_Visible = New_Visible
    Call UCpaint
    PropertyChanged "PictureVisible"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    Call UCpaint
    PropertyChanged "ForeColor"
End Property
Public Property Get OnMouseMoveForeColor() As OLE_COLOR
    OnMouseMoveForeColor = m_Fc
End Property
Public Property Let OnMouseMoveForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_Fc = New_ForeColor
    PropertyChanged "OnMouseMoveForeColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Initialize()
    img.Visible = img_Visible
    Call UCpaint
    Call DefLabelVal
End Sub

Private Sub UserControl_InitProperties()
    img_Visible = True
    Call UCpaint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X > 0 Or X < UserControl.ScaleWidth Or Y > 0 Or Y < UserControl.ScaleHeight Then
        Call LabelChange
    Else
        Call DefLabelVal
        Timer1.Enabled = False
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent Click
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub
Sub ShowAboutBox()
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Debasis")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    m_d_Fc = lblCaption.ForeColor
    m_Fc = PropBag.ReadProperty("OnMouseMoveForeColor", m_d_Fc)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    img_Visible = PropBag.ReadProperty("PictureVisible", True)
End Sub

Private Sub UserControl_Resize()
    Call UCpaint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Debasis")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HFF0000)
    Call PropBag.WriteProperty("OnMouseMoveForeColor", m_Fc, m_d_Fc)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PictureVisible", img_Visible, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=img,img,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = img.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set img.Picture = New_Picture
    Call UCpaint
    PropertyChanged "Picture"
End Property


