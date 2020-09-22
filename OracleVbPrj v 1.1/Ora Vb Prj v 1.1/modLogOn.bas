Attribute VB_Name = "modLogOn"
Option Explicit
Public db As New ADODB.Connection
Public uid As String
Public pwd As String
Public d_Base As String
Public gDmsg As New OmsgBox.DMsgBox
Public Sub OracleConnect()
    Screen.MousePointer = vbHourglass
    On Error GoTo logonError
    Set gDmsg = New OmsgBox.DMsgBox
    Dim Conn As String
    Dim drv As String
    uid = Trim$(frmLogOn.txtUserId.Text)
    pwd = Trim$(frmLogOn.txtPwd.Text)
    d_Base = Trim$(frmLogOn.txtDb.Text)
    Set db = New ADODB.Connection
        If frmLogOn.txtDb.Text <> "" Then
            Conn = "UID= " & uid & ";PWD=" & pwd & ";DRIVER={Microsoft ODBC For Oracle};" _
            & "SERVER=" & d_Base & ";"
        Else
            Conn = "UID= " & uid & ";PWD=" & pwd & ";DRIVER={Microsoft ODBC For Oracle};"
        End If
            drv = "Microsoft ODBC For Oracle"
        With db
            .ConnectionString = Conn
            .CursorLocation = adUseClient
            .Open
        End With
    Screen.MousePointer = vbDefault
logonError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        gDmsg.DebMsgBox Err.Description, "Error", DebmsgCritical
        With frmLogOn
            .txtUserId.Text = ""
            .txtPwd.Text = ""
            .txtDb.Text = ""
            .txtUserId.SetFocus
        End With
    Else
        Screen.MousePointer = vbDefault
        db.BeginTrans
        Unload frmLogOn
        frmMain.Show
    End If
End Sub




