VERSION 5.00
Begin VB.UserControl OFrmCtl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ToolboxBitmap   =   "OFrmCtl.ctx":0000
   Begin DebaFrmCtl.OClose OClose1 
      Height          =   255
      Left            =   570
      TabIndex        =   2
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin DebaFrmCtl.OMax OMax1 
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin DebaFrmCtl.OMin OMin1 
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   45
      Top             =   30
      Width           =   255
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   90
      Left            =   1860
      TabIndex        =   4
      Top             =   435
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   225
      Left            =   1725
      TabIndex        =   3
      Top             =   375
      Width           =   360
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main Menu"
      Begin VB.Menu mnuMinWin 
         Caption         =   "Minimize Window"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendTray 
         Caption         =   "Send To Tray"
      End
   End
End
Attribute VB_Name = "OFrmCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Form Minimize Control
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh

Option Explicit

Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

Private WithEvents lbStatus As Label
Attribute lbStatus.VB_VarHelpID = -1

Private shpStatus As Shape

Private t_Icon As NOTIFYICONDATA

Dim trayInfo As String

Dim stCaption As String

Private w_State As Long

Private Sub Init()
    ' Initialize The Control And Parent
    Dim l As Long
    Dim hMenu As Long
    Dim cnt As Long
    Dim hRgn As Long
    Dim fR As Long
    Dim r As RECT
    Dim fx As Integer, fy As Integer
    
    With frm
        .AutoRedraw = True
        .ScaleMode = vbPixels
    End With
    
    ' Get the current window style of the form.
    l = GetWindowLong(frm.hwnd, GWL_STYLE)
        
    ' Set the window style
    l = l And Not (WS_BORDER Or WS_CAPTION Or WS_THICKFRAME)
    SetWindowLong frm.hwnd, GWL_STYLE, l
    
    hMenu = GetSystemMenu(frm.hwnd, False)
    If hMenu Then
        cnt = GetMenuItemCount(hMenu)
        If cnt Then
            RemoveMenu hMenu, cnt - 3, MF_BYPOSITION Or MF_REMOVE  'Remove Maximize
            'RemoveMenu hMenu, cnt - 5, MF_BYPOSITION Or MF_REMOVE
            'RemoveMenu hMenu, cnt - 6, MF_BYPOSITION Or MF_REMOVE 'Remove Move Window
            DrawMenuBar frm.hwnd
        End If
    End If
    
    
    ' Move And Size the Window
    fR = GetWindowRect(frm.hwnd, r)
    fx = r.Right - r.Left
    fy = r.Bottom - r.Top - GetSystemMetrics(SM_CYCAPTION)
    fR = MoveWindow(frm.hwnd, r.Left, r.Top, fx%, fy%, True)
    With Image1
        .Picture = frm.Icon
        .Stretch = True
        .Width = 18
        .Height = 18
    End With
    
    ' Form Caption
    With lblCaption
        .Caption = frm.Caption
        .Font = "Verdana"
        .FontBold = True
        .FontSize = 10
        .ForeColor = vbWhite
    End With
    
    With lblS
        .Caption = lblCaption.Caption
        .Font = lblCaption.Font
        .FontBold = lblCaption.FontBold
        .FontSize = lblCaption.FontSize
    End With
    Call PositionCtl
    Call PaintForm
    
End Sub
Sub PositionCtl()
    Select Case frm.BorderStyle
        Case 0
            OClose1.Visible = False
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = False
            
        Case 1
            OClose1.Visible = True
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = True
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 6, 2
            
        Case 2
            OClose1.Visible = True
            OMax1.Visible = True
            OMin1.Visible = True
            Image1.Visible = True
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 6, 2
            OMax1.Move OClose1.Left - OMax1.Width - 1, 2
            OMin1.Move OMax1.Left - OMin1.Width - 1, 2
            
        Case 3
            OClose1.Visible = True
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = True
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 4, 2
            
        Case 4
            OClose1.Visible = True
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = False
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 4, 2
        
        Case 5
            OClose1.Visible = True
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = False
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 4, 2
            
        Case Else
            OClose1.Visible = True
            OMax1.Visible = False
            OMin1.Visible = False
            Image1.Visible = False
            OClose1.Move UserControl.ScaleWidth - OClose1.Width - 4, 2
    End Select
    Image1.Move 2, 2
    lblCaption.Move 0, 4, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblS.Move 1, 5, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblCaption.ZOrder 0
End Sub
Sub FormSysmenu()
    ' Sysmenu
    Dim cm As Long
    Dim hm As Long
    Dim lf As Long
    Dim r As RECT
    Dim ef As MenuControlConstants
    Dim pt As POINTAPI
    Dim l As Long
    
    GetCursorPos pt
    'Get System Menu
    hm = GetSystemMenu(frm.hwnd, &H0&)
    
    If hm <> 0 Then
    
        lf = ef Or (TPM_RETURNCMD)
        cm = TrackPopupMenu(hm, lf, pt.X, pt.Y, &H0&, frm.hwnd, r)
    End If
    
    If cm <> 0 Then
        Call PostMessage(frm.hwnd, WM_SYSCOMMAND, cm, hm)
    End If
 End Sub
Sub PaintForm()
    Dim frmWidth As Integer, frmHeight As Integer
    Dim hRgn As Long
    Dim X, Y
    Dim i
    Dim rc As RECT
    Dim max As Long
    Dim p As POINTAPI
    Dim l As Long
    Dim Gx, Bx As Integer
    frmWidth = frm.ScaleWidth
    frmHeight = frm.ScaleHeight
    Gx = 220
        For i = 0 To 24
            UserControl.Line (0, i)-(frmWidth, i), RGB(0, Gx, 0)
            Gx = Gx - 4
            If Gx < 4 Then
                Gx = 4
            End If
        Next i
    Bx = 160
        For i = 24 To 38
            frm.Line (0, i)-(frmWidth, i), RGB(0, 0, Bx)
            Bx = Bx + 10
        Next i
        For i = 39 To frmHeight
            frm.Line (0, i)-(frmWidth, i), RGB(0, 0, Bx)
        Next i
   
    Bx = 200
        For Y = 0 To 5
            frm.Line (Y, 24)-(Y, frmHeight), RGB(0, 0, Bx)
            Bx = Bx + 10
        Next Y
    Bx = 150
        For Y = frmHeight To frmHeight - 10 Step -1
            frm.Line (0, Y)-(frmWidth, Y), RGB(0, 0, Bx)
             Bx = Bx + 10
        Next Y
         Bx = 220
         For Y = frmWidth - 7 To frmWidth
            frm.Line (Y, 24)-(Y, frmHeight), RGB(0, 0, Bx)
            Bx = Bx - 10
            If Bx < 10 Then
                Bx = 10
            End If
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, frm.ScaleWidth, frm.ScaleHeight, 10, 10)
        SetWindowRgn frm.hwnd, hRgn, True
End Sub
Private Sub frm_Load()
    Call Init
    OMin1.ToolTipText = "Minimize Window Or Send To Tray"
    OClose1.ToolTipText = "Close"
    Call PositionCtl
    If frm.BorderStyle = 2 Or frm.BorderStyle = 5 Then
        Set lbStatus = frm.Controls.Add("VB.Label", "lbStatus", frm)
        lbStatus.Move 14, frm.ScaleHeight - 30, frm.ScaleWidth - 26, 15
        lbStatus.Visible = True
        With lbStatus
            .Caption = StatusCaption
            .BackStyle = 0
            .Font = "Tahoma"
            .FontSize = 9
            .FontBold = True
            .ForeColor = vbWhite
            .Alignment = 0
        End With
        
        Set shpStatus = frm.Controls.Add("VB.Shape", "shpStatus", frm)
        With shpStatus
            .Shape = 4
            .BorderWidth = 2
            .BorderColor = &HC00000
            .Move 8, frm.ScaleHeight - 40, frm.ScaleWidth - 16, 24
            .Visible = True
        End With
    End If
End Sub
Private Sub frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Tray Icon
    Dim Result As Long
        Dim Msg As Long
        If frm.ScaleMode = vbPixels Then
            Msg = X
        Else
            Msg = X / Screen.TwipsPerPixelX
        End If
    Select Case Msg
        
        Case WM_LBUTTONUP
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon ' Delete Icon From SysTray
        
        Case WM_LBUTTONDBLCLK
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
        
        Case WM_RBUTTONUP
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
        
        Case WM_RBUTTONDBLCLK
            frm.WindowState = w_State
            Result = SetForegroundWindow(frm.hwnd)
            frm.Show
            Shell_NotifyIcon NIM_DELETE, t_Icon
       End Select
End Sub

Private Sub frm_Resize()
    Dim i As Integer
    'Checking Window State
    If frm.WindowState = vbNormal Then
        OMax1.frmState = False
        Call OMax1.UcState
        OMax1.ToolTipText = "Maximize"
    Else
        OMax1.frmState = True
        Call OMax1.UcState
        OMax1.ToolTipText = "Restore Down"
    End If
    'Movw Usercontrol To Top Of The Form
    MoveWindow UserControl.hwnd, 0, 0, frm.ScaleWidth, 24, 1
    Call PositionCtl
    Call PaintForm
    On Error Resume Next
    lbStatus.Move 14, frm.ScaleHeight - 30, frm.ScaleWidth - 26, 15
    shpStatus.Move 8, frm.ScaleHeight - 30, frm.ScaleWidth - 16, 20
End Sub

Private Sub lblCaption_DblClick()
If frm.BorderStyle = 2 Then
    If frm.WindowState = vbNormal Then
        ShowWindow frm.hwnd, SW_MAXIMIZE
    Else
        ShowWindow frm.hwnd, SW_RESTORE
    End If
End If
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call FormSysmenu
    End If
End Sub

Private Sub lblS_DblClick()
If frm.BorderStyle = 2 Then
    If frm.WindowState = vbNormal Then
        ShowWindow frm.hwnd, SW_MAXIMIZE
    Else
        ShowWindow frm.hwnd, SW_RESTORE
    End If
End If
End Sub
Private Sub lblS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
Private Sub lblS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call FormSysmenu
    End If
End Sub

Private Sub mnuMinWin_Click()
    ShowWindow frm.hwnd, SW_MINIMIZE
End Sub

Private Sub mnuSendTray_Click()
    With t_Icon
        .cbSize = Len(t_Icon)
        .hwnd = frm.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .szTip = "Developed By Debasis Ghosh (debughosh@vsnl.net)" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = trayInfo
        .szInfoTitle = "" & frm.Caption & ""
        .dwInfoFlags = NIIF_INFO
        .uTimeout = 3000
   End With
        Shell_NotifyIcon NIM_ADD, t_Icon 'Add Icon To Systray
        w_State = frm.WindowState ' Hold Window State
        frm.Hide
End Sub

Private Sub OClose1_Click()
    On Error Resume Next
    Unload frm
End Sub

Private Sub OMax1_Click()
    On Error Resume Next
    If frm.WindowState = vbNormal Then
        OMax1.frmState = False
        Call OMax1.UcState
        ShowWindow frm.hwnd, SW_MAXIMIZE
    Else
        OMax1.frmState = True
        Call OMax1.UcState
        ShowWindow frm.hwnd, SW_RESTORE
    End If
End Sub

Private Sub OMin1_Click()
    PopupMenu mnuMain, , , , mnuMinWin
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
    trayInfo = PropBag.ReadProperty("SysTrayInfo", "Debasis Ghosh")
    stCaption = PropBag.ReadProperty("StatusCaption", "Debasis Ghosh")
End Sub
Public Property Get SysTrayInfo() As String
    SysTrayInfo = trayInfo
End Property
Public Property Let SysTrayInfo(ByVal New_SysTrayInfo As String)
    trayInfo = New_SysTrayInfo
    PropertyChanged "SysTrayInfo"
End Property
Public Property Get StatusCaption() As String
    StatusCaption = stCaption
End Property
Public Property Let StatusCaption(ByVal New_Caption As String)
    stCaption = New_Caption
    PropertyChanged "StatusCaption"
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SysTrayInfo", trayInfo, "Debasis Ghosh")
    Call PropBag.WriteProperty("StatusCaption", stCaption, "Debasis Ghosh")
End Sub
