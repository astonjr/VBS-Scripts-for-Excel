Option Explicit

Sub Merge_Same_Cells_SpecificHeader()

    Application.DisplayAlerts = False

    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim row As Long
    Dim startRow As Long
    Dim header As String
    
    ' Define the header you are looking for
    header = "Email ID"
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If ws.Cells(1, col).Value = header Then
            startRow = 2 ' Initialize startRow
            For row = 2 To lastRow + 1 ' Add 1 to handle the last group
                If ws.Cells(row, col).Value <> ws.Cells(row - 1, col).Value Or row = lastRow + 1 Then
                    ' If the current cell is different from the previous or it's the last row
                    If row - startRow > 1 Then
                        ' If there is more than one cell to merge
                        Set rng = ws.Range(ws.Cells(startRow, col), ws.Cells(row - 1, col))
                        rng.Merge
                        rng.HorizontalAlignment = xlCenter
                        rng.VerticalAlignment = xlTop
                    End If
                    startRow = row ' Reset startRow to the current row
                End If
            Next row
        End If
    Next col

    Application.DisplayAlerts = True

End Sub
