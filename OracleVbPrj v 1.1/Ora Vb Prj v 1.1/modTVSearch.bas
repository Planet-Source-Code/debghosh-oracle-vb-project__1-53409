Attribute VB_Name = "modTVSearch"
Option Explicit
Const intIndentSpaces = 4
Dim strFile
Dim ff As Integer
Dim intIndentLevel As Integer
Public Sub FindNode(TView As TreeView)
    Dim i As Integer
    Dim FindPos As Integer
    Dim r
    Dim Found As Boolean
    Found = False
    r = InputBox("Enter Search String Here", "Search Data")
    If r = "" Then
        Exit Sub
    End If
    
    For i = 1 To TView.Nodes.Count
        FindPos = InStr(1, TView.Nodes(i).Text, r, vbTextCompare)
        If FindPos <> 0 Then
            TView.SetFocus
            TView.Nodes(i).Selected = True
            FindPos = 0
            Found = True
        Else
            If i = TView.Nodes.Count And Found = False Then
                gDmsg.DebMsgBox "Not found"
            End If
        End If
    Next i
End Sub
Public Sub SaveData(cd As CommonDialog, TView As TreeView)
    Dim objNode As Node
    cd.CancelError = True
On Error GoTo cdError
    cd.Flags = cdlOFNOverwritePrompt
    cd.DialogTitle = "Save As"
    cd.Filter = "Text File(*.Txt)|*.Txt|"
    cd.ShowSave
    strFile = cd.FileName
    ff = FreeFile
    Open strFile For Output As #ff
    intIndentLevel = 0
    Set objNode = TView.Nodes(1)
    ParseTree objNode
    Close #ff
cdError:
    Exit Sub
End Sub
Private Sub ParseTree(objNode As Node)
    Print #ff, Space(intIndentLevel * intIndentSpaces) & objNode.Text
    If objNode.Children > 0 Then
        intIndentLevel = intIndentLevel + 1
        ParseTree objNode.Child
    End If
    Set objNode = objNode.Next
    If TypeName(objNode) <> "Nothing" Then
        ParseTree objNode
    Else
        intIndentLevel = intIndentLevel - 1
    End If
End Sub



