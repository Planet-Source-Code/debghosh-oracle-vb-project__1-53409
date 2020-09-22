Attribute VB_Name = "modExportToExcel"
Option Explicit
Public Sub Exp2Exl(flg As MSHFlexGrid, bkName As String)
On Error GoTo fgError
If flg.TextMatrix(0, 1) <> "" Then
    Dim exc As Excel.Application
    Dim d As Excel.Workbook
    Dim w As Excel.Worksheet
    Set exc = CreateObject("Excel.Application")
    Set d = exc.Workbooks.Add
    Set w = d.Worksheets(1)
    exc.Visible = True
    Dim i, j
    With flg
        For i = 0 To .Rows - 1
            For j = 1 To .Cols - 1
                exc.Cells(i + 1, j) = .TextMatrix(i, j)
                exc.Cells(i + 1, j).Borders.LineStyle = xlDouble
            Next j
        Next i
        exc.Range("A1:" & Chr(65 + j) & 1).Font.Bold = True
        exc.Columns("$A:" & "$" & Chr(65 + j)).AutoFit
    End With
    w.Name = bkName
End If
fgError:
    If Err.Number <> 0 Then
        gDmsg.DebMsgBox "Error: " & Err.Description & " ", "Error", DebmsgCritical
        Exit Sub
    End If
End Sub


