VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh


Option Explicit
Public WithEvents Combo1 As ComboBox
Attribute Combo1.VB_VarHelpID = -1
Const WM_USER = &H400
Const CB_SHOWDROPDOWN = &H14F
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim sComboText As String
     Dim iLoop As Integer
     Dim sTempString As String
     Dim lReturn As Long
     If KeyCode >= KEY_A And KeyCode <= KEY_Z Then
       'only look at letters A-Z
       sTempString = Combo1.Text
       If Len(sTempString) = 1 Then sComboText = sTempString
       lReturn = SendMessage(Combo1.hwnd, WM_SETREDRAW, False, 0&)
       For iLoop = 0 To (Combo1.ListCount - 1)
         If UCase((sTempString & Mid$(Combo1.List(iLoop), _
           Len(sTempString) + 1))) = UCase(Combo1.List(iLoop)) Then
           Combo1.ListIndex = iLoop
           Combo1.Text = Combo1.List(iLoop)
           Combo1.SelStart = Len(sTempString)
           Combo1.SelLength = Len(Combo1.Text) - (Len(sTempString))
           sComboText = sComboText & Mid$(sTempString, Len(sComboText) + 1)
           Exit For
         Else
           If InStr(UCase(sTempString), UCase(sComboText)) Then
             sComboText = sComboText & Mid$(sTempString, Len(sComboText) _
             + 1)
             Combo1.Text = sComboText
             Combo1.SelStart = Len(Combo1.Text)
           Else
             sComboText = sTempString
             End If
         End If
       Next iLoop
       lReturn = SendMessage(Combo1.hwnd, _
       WM_SETREDRAW, True, 0&)
     End If
End Sub




