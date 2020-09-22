Attribute VB_Name = "modTextSubclass"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Const GWL_WNDPROC = -4
    Public Const WM_RBUTTONUP = &H205
    Private Const WM_COPY As Long = &H301
    Private Const WM_PASTE As Long = &H302
    Global lpCmb As Long
    Global lpPrevWndProc As Long
    Global gHW As Long
    Global gCmb As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
           (ByVal lpPrevWndFunc As Long, _
            ByVal hwnd As Long, _
            ByVal Msg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
           (ByVal hwnd As Long, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As Long) As Long
Public Sub Hook()
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
                                     AddressOf WindowProc)
End Sub
    Public Sub UnHook()
        Dim lngReturnValue As Long
        lngReturnValue = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
    End Sub
    Function WindowProc(ByVal hw As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

        Select Case uMsg
            Case WM_RBUTTONUP
                gDmsg.DebMsgBox "Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", "ORACLE VB Project", DebmsgInformation
            Case WM_COPY
                gDmsg.DebMsgBox "Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", "ORACLE VB Project", DebmsgInformation
            Case WM_PASTE
                gDmsg.DebMsgBox "Project Developed By Debasis Ghosh. Feel Free to comment on this at debughosh@vsnl.net", "ORACLE VB Project", DebmsgInformation
            Case Else
                WindowProc = CallWindowProc(lpPrevWndProc, hw, _
                                           uMsg, wParam, lParam)
        End Select
    End Function
Public Sub FlatBorder(ByVal hwnd As Long)
    Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub



