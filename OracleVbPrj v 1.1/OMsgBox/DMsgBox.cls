VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim hRgn As Long
Public Enum DebaMsgStyle
    DebmsgYesNo = 0
    DebmsgYesNoCancel = 1
    DebmsgOkCancel = 2
    DebmsgOkOnly = 3
    DebmsgAbortRetryIgnore = 4
    DebmsgCritical = 5
    DebmsgDefault = 6
    DebmsgExclamation = 7
    DebmsgInformation = 8
    DebmsgQuestion = 9
    DebmsgRetryCancel = 10
End Enum
Public Enum msgResult
    DebmsgYes = 0
    DebmsgNo = 1
    DebmsgCancel = 2
    DebmsgOk = 3
    DebmsgRetry = 4
    DebmsgAbort = 5
    DebmsgIgnore = 6
End Enum
Dim rndRG As Long
Dim y, Y1, Y2 As Integer
Public Function DebMsgBox(ByVal Msg As String, Optional Title As String = "Debasis", Optional Style As DebaMsgStyle = DebmsgOkOnly) As Integer
    m_Width = 0
    With frmMsgBox
        .AutoRedraw = True
        .ScaleMode = vbPixels
    End With
    Call DefBtn
    Select Case Style
            Case DebmsgOkCancel   'msgOkCancel
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 200 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 50
                    cmdTop = 130
                Else
                    m_Width = 230
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width - 10, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .cmdCancel.Visible = True
                .cmdCancel.Move (m_Width) / 2 + 10, .cmdOk.Top, .cmdCancel.Width, .cmdCancel.Height
                .cmdCancel.TabIndex = 1
                .DClose1.Enabled = False
            End With
            
            Case DebmsgYesNo         'msgYesNo
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 200 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 50
                    cmdTop = 130
                Else
                    m_Width = 230
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdYes.Visible = True
                .cmdYes.Move (m_Width) / 2 - .cmdYes.Width - 10, cmdTop, .cmdYes.Width, .cmdYes.Height
                .cmdYes.TabIndex = 0
                .cmdNo.Visible = True
                .cmdNo.Move (m_Width) / 2 + 10, .cmdYes.Top, .cmdNo.Width, .cmdNo.Height
                .cmdNo.TabIndex = 1
                .DClose1.Enabled = False
            End With
            
            Case DebmsgRetryCancel   'msgRetryCancel
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 200 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 50
                    cmdTop = 130
                Else
                    m_Width = 230
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdRetry.Visible = True
                .cmdRetry.Move (m_Width) / 2 - .cmdRetry.Width - 10, cmdTop, .cmdRetry.Width, .cmdRetry.Height
                .cmdRetry.TabIndex = 0
                .cmdCancel.Visible = True
                .cmdCancel.Move (m_Width) / 2 + 10, .cmdRetry.Top, .cmdCancel.Width, .cmdCancel.Height
                .cmdRetry.TabIndex = 1
                .DClose1.Enabled = False
            End With
            
            Case DebmsgYesNoCancel 'msgYesNoCancel
            
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 260 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 80
                    cmdTop = 130
                Else
                    m_Width = 320
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdNo.Visible = True
                .cmdNo.Move (m_Width) / 2 - .cmdNo.Width / 2, cmdTop, .cmdNo.Width, .cmdNo.Height
                .cmdYes.Visible = True
                .cmdYes.Move .cmdNo.Left - .cmdNo.Width - 15, .cmdNo.Top, .cmdYes.Width, .cmdYes.Height
                .cmdCancel.Visible = True
                .cmdCancel.Move .cmdNo.Left + .cmdNo.Width + 15, .cmdNo.Top, .cmdCancel.Width, .cmdCancel.Height
                .DClose1.Enabled = False
            End With
            With frmMsgBox
                .cmdYes.TabIndex = 0
                .cmdNo.TabIndex = 1
                .cmdCancel.TabIndex = 2
            End With
        
        Case DebmsgAbortRetryIgnore 'msgAbortRetryIgnore
            
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 260 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 80
                    cmdTop = 130
                Else
                    m_Width = 320
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdRetry.Visible = True
                .cmdRetry.Move (m_Width) / 2 - .cmdRetry.Width / 2, cmdTop, .cmdRetry.Width, .cmdRetry.Height
                .cmdAbort.Visible = True
                .cmdAbort.Move .cmdRetry.Left - .cmdAbort.Width - 15, .cmdRetry.Top, .cmdAbort.Width, .cmdAbort.Height
                .cmdIgnore.Visible = True
                .cmdIgnore.Move .cmdRetry.Left + .cmdRetry.Width + 15, .cmdRetry.Top, .cmdIgnore.Width, .cmdIgnore.Height
                .DClose1.Enabled = False
            End With
            With frmMsgBox
                .cmdAbort.TabIndex = 0
                .cmdRetry.TabIndex = 1
                .cmdIgnore.TabIndex = 2
            End With
            
        Case DebmsgOkOnly  'msgOkonly
            
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 200 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 50
                    cmdTop = 130
                Else
                    m_Width = 260
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
            
            Case DebmsgCritical   'msgCritical
            
            frmMsgBox.imgCritical.Visible = True
            With frmMsgBox
                .imgCritical.Left = 20
                .imgCritical.Top = 50
                .lblText.Left = .imgCritical.Left + .imgCritical.Width + 20
                .lblText.Top = .imgCritical.Top
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 240 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 70
                    cmdTop = 130
                Else
                    m_Width = 330
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
            
            Case DebmsgInformation 'msgInformation
            
            frmMsgBox.imgInformation.Visible = True
            With frmMsgBox
                .imgInformation.Left = 20
                .imgInformation.Top = 50
                .lblText.Left = .imgInformation.Left + .imgInformation.Width + 20
                .lblText.Top = .imgInformation.Top
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 240 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 70
                    cmdTop = 130
                Else
                    m_Width = 330
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
            
        Case DebmsgQuestion  'msgQuestion
            
            frmMsgBox.imgQuestion.Visible = True
            
            With frmMsgBox
                .imgQuestion.Left = 20
                .imgQuestion.Top = 50
                .lblText.Left = .imgQuestion.Left + .imgQuestion.Width + 20
                .lblText.Top = .imgQuestion.Top
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 240 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 70
                    cmdTop = 130
                Else
                    m_Width = 330
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
            
            Case DebmsgExclamation 'Exclamation
                
            frmMsgBox.imgExclamation.Visible = True
            With frmMsgBox
                .imgExclamation.Left = 20
                .imgExclamation.Top = 50
                .lblText.Left = .imgExclamation.Left + .imgExclamation.Width + 20
                .lblText.Top = .imgExclamation.Top
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 240 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 70
                    cmdTop = 130
                Else
                    m_Width = 330
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
            
        Case Else
            With frmMsgBox
                .lblText.Left = 30
                .lblText.Top = 50
                .lblText.Caption = Msg
            End With
            If frmMsgBox.lblText.Width < 450 Then
                If frmMsgBox.lblText.Width > 200 Then
                    m_Width = frmMsgBox.lblText.Left + frmMsgBox.lblText.Width + 50
                    cmdTop = 130
                Else
                    m_Width = 260
                    cmdTop = 130
                End If
            Else
                Call WrapLabel
                m_Width = 455
                cmdTop = frmMsgBox.lblText.Top + frmMsgBox.lblText.Height + 40
            End If
            With frmMsgBox
                .cmdOk.Visible = True
                .cmdOk.Move (m_Width) / 2 - .cmdOk.Width / 2, cmdTop, .cmdOk.Width, .cmdOk.Height
                .cmdOk.TabIndex = 0
                .DClose1.Enabled = True
            End With
    End Select
    frmMsgBox.Width = m_Width * 15
    frmMsgBox.Height = (cmdTop + 45) * 15
    With frmMsgBox
        .lblHead.Caption = Title
        .lblHead.Move 0, 2, .ScaleWidth, 30
        .lblHShade.Caption = Title
        .lblHShade.Move 1, 3, .ScaleWidth, 30
        .lblHead.ZOrder 0
        .DClose1.Move .ScaleWidth - .DClose1.Width - 7, 2
    End With
    FormResize frmMsgBox, frmMsgBox.ScaleWidth, frmMsgBox.ScaleHeight
        For y = cmdTop - 10 To cmdTop - 9
            frmMsgBox.Line (6, y)-(frmMsgBox.ScaleWidth - 6, y), RGB(0, 250, 10)
        Next y
    frmMsgBox.Show vbModal
    'Beep
    DebMsgBox = frmMsgBox.DebMsgResult
End Function
Private Sub WrapLabel()
    Dim XStr As String
    Dim Break As Integer
    Dim x As Integer
    Dim y As Integer
    Dim Z As Integer
    Break = 62 ' the width to break the text
   With frmMsgBox.lblText
      If .Width > 425 Then
           Do Until .Width <= 425
               x = 0
               Do Until x + Break >= Len(.Caption)
                   If .Width < 425 Then Exit Do
                   Do Until InStr(Left(Right(.Caption, Len(.Caption) - x), Break), vbCr) = 0 And Left(Right(.Caption, Len(.Caption) - x), 1) <> vbLf
                       XStr = Left(Right(.Caption, Len(.Caption) - x - 1), Break)
                       x = x + 1
                   Loop
                    
                   If x > 0 Then
                       XStr = Left(Right(.Caption, Len(.Caption) - x), Break)
                   Else
                       XStr = Left(.Caption, Break)
                   End If
                   XStr = Left(XStr, InStrRev(XStr, " "))

                   If XStr = "" Then
                       y = 1
                       Do
                           Z = 15
                           XStr = Left(Right(.Caption, Len(.Caption) - x), Break / y - Z)
                            
                           Do Until InStr(Left(Right(.Caption, Len(.Caption) - x - Len(XStr)), 10), vbCr) = 0
                               Z = Z + Break / 1.5
                               XStr = Left(Right(.Caption, Len(.Caption) - x), Break / y - Z)
                           Loop
                            
                           If Right(Left(.Caption, Len(XStr) + 3), 3) = "-" & vbCrLf Then
                               y = y + 1
                           Else
                               Exit Do
                           End If
                           y = y + 1
                       Loop
                       XStr = XStr & "-" & vbCrLf
                        
                       .Caption = Left(.Caption, x) & XStr & Right(.Caption, Len(.Caption) - Len(XStr) - x + 3)
                   Else
                       XStr = Left(XStr, Len(XStr) - 1) & vbCrLf
                        
                       .Caption = Left(.Caption, x) & XStr & Right(.Caption, Len(.Caption) - Len(XStr) - x + 1)
                   End If
                    
                   x = x + Len(XStr)
                   If .Width <= 425 Then GoTo Done
               Loop
               If .Width <= 425 Then GoTo Done
               Break = Break - 15
           Loop
Done:
       End If
   End With
End Sub
Private Sub DefBtn()
    Dim i
    With frmMsgBox
        .cmdAbort.Visible = False
        .cmdCancel.Visible = False
        .cmdIgnore.Visible = False
        .cmdNo.Visible = False
        .cmdOk.Visible = False
        .cmdRetry.Visible = False
        .cmdYes.Visible = False
        .imgExclamation.Visible = False
        .imgInformation.Visible = False
        .imgQuestion.Visible = False
        .imgCritical.Visible = False
    End With
    
End Sub


Private Sub FormResize(frmName As Form, frmWidth As Integer, frmHeight As Integer)
    Dim x, y
    Dim i
    Dim rc As RECT
    Dim rndRG As Long
    Dim max As Long
    Dim p As POINTAPI
    Dim l As Long
    Dim Gx, Bx As Integer
    Static sB As Integer
    frmName.AutoRedraw = True
    frmName.ScaleMode = vbPixels
    'frmName.Cls
    Gx = 220
        For i = 0 To 24
            frmName.Line (0, i)-(frmWidth, i), RGB(0, Gx, 0)
            Gx = Gx - 4
            If Gx < 4 Then
                Gx = 4
            End If
        Next i
    Bx = 160
        For i = 24 To 31
            frmName.Line (0, i)-(frmWidth, i), RGB(0, 0, Bx)
            Bx = Bx + 10
        Next i
        For i = 31 To frmHeight
            frmName.Line (0, i)-(frmWidth, i), RGB(0, 0, Bx)
        Next i
   
    Bx = 200
        For y = 0 To 5
            frmName.Line (y, 24)-(y, frmHeight), RGB(0, 0, Bx)
            Bx = Bx + 10
        Next y
    Bx = 150
        For y = frmHeight To frmHeight - 10 Step -1
            frmName.Line (0, y)-(frmWidth, y), RGB(0, 0, Bx)
             Bx = Bx + 10
        Next y
         Bx = 220
         For y = frmWidth - 7 To frmWidth
            frmName.Line (y, 24)-(y, frmHeight), RGB(0, 0, Bx)
            Bx = Bx - 10
            If Bx < 10 Then
                Bx = 10
            End If
        Next y
    rndRG = CreateRoundRectRgn(0, 0, frmWidth, frmHeight, 20, 20)
    SetWindowRgn frmName.hWnd, rndRG, True
    DeleteObject rndRG
End Sub


