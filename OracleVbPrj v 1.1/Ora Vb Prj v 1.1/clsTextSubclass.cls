VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh

'Subclassing Textbox Control.
Option Explicit
Public Sub TxtPop(txt As TextBox)
    FlatBorder txt.hwnd
    gHW = txt.hwnd
    Hook
End Sub
Private Sub Class_Terminate()
    UnHook
End Sub
Public Sub txtUnhook()
    UnHook
End Sub
