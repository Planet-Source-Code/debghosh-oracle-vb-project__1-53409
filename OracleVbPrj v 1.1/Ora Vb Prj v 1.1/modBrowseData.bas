Attribute VB_Name = "modBrowseData"
Option Explicit
Dim sr As New ADODB.Recordset
Sub BrowseDataForm()
    With frmBrowseData
    Screen.MousePointer = vbHourglass
        .lblSelectSchemaName.Visible = False
        .cmbSchema.Visible = False
        .cmdLoad.Visible = False
    Dim c As Integer
On Error GoTo rsError
    .Fg.ColWidth(0) = 400
    .pb.Visible = True
    .pb.Min = 0
    Set sr = db.OpenSchema(adSchemaSchemata, Array(Empty, Empty, Empty))
    .pb.Max = sr.RecordCount
    c = 1
    Do Until sr.EOF
        .cmbSchema.AddItem sr!SCHEMA_NAME
        .pb.Value = c
        c = c + 1
        sr.MoveNext
    Loop
        .cmbSchema.ListIndex = 0
        .pb.Visible = False
        .lblSelectSchemaName.Visible = True
        .cmbSchema.Visible = True
        .cmdLoad.Visible = True
    Screen.MousePointer = vbDefault
    End With
    sr.Close
rsError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error Occured While Processing : " & Err.Description & " "
    End If
End Sub


