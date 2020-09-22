Attribute VB_Name = "modRtSaveAs"
Option Explicit
Public Sub SaveTextAs(rtB As RichTextBox, cdB As CommonDialog)
    If rtB.Text <> "" Then
    cdB.CancelError = True
On Error GoTo rtError
    cdB.Flags = cdlOFNOverwritePrompt
    cdB.DialogTitle = "Save As (Save In RichText Format)"
    cdB.Filter = "Text File (*.Txt)|*.txt|SQL File (*.SQL)|*.SQL|Rich Text File (*.Rtf)|*.Rtf|"
    cdB.ShowSave
    rtB.SaveFile cdB.FileName, rtfRTF
    Else
        gDmsg.DebMsgBox "Text Box Empty"
    End If
rtError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description, "Error", DebmsgCritical
    End If
End Sub
Public Sub OpenRtText(rt As RichTextBox, cd As CommonDialog)
    cd.CancelError = True
On Error GoTo rtError
    cd.DialogTitle = "Open File"
    cd.Filter = "Text File (*.Txt)|*.txt|SQL File (*.SQL)|*.SQL|Rich Text File (*.Rtf)|*.Rtf|"
    cd.ShowOpen
    rt.LoadFile cd.FileName
rtError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox Err.Description, "Error", DebmsgCritical
    End If
End Sub


