Option Explicit

Sub SortAndMergeEventsByEmail()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim uniqueEmails As Object
    Dim email As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim emailCol As Long
    Dim startDateCol As Long
    Dim headerRow As Long
    Dim i As Long
    Dim rng As Range

    Set ws = ActiveSheet
    headerRow = 1 ' Assuming the headers are in the first row

    ' Find last row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find the column indices based on the headers
    emailCol = 0
    startDateCol = 0

    For i = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(headerRow, i).Value = "Email ID" Then
            emailCol = i
        ElseIf ws.Cells(headerRow, i).Value = "Start Date" Then
            startDateCol = i
        End If
    Next i

    ' Check if the columns were found
    If emailCol = 0 Or startDateCol = 0 Then
        MsgBox "Error: 'Email ID' or 'Start Date' column not found."
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Unmerge all cells in the Email ID column
    ws.Columns(emailCol).UnMerge

    ' Create a dictionary to store unique email IDs
    Set uniqueEmails = CreateObject("Scripting.Dictionary")

    ' Loop through the rows to identify unique email IDs
    For i = headerRow + 1 To lastRow ' Assuming data starts from row after header
        email = ws.Cells(i, emailCol).Value ' Get the email ID value
        If Not uniqueEmails.exists(email) Then
            uniqueEmails.Add email, i ' Store the starting row for each unique email
        End If
    Next i

    ' Loop through unique email IDs and sort events by start date
    For Each email In uniqueEmails.Keys
        startRow = uniqueEmails(email)
        
        ' Find the end row for the current email
        For endRow = startRow To lastRow
            If ws.Cells(endRow + 1, emailCol).Value <> email Then Exit For
        Next endRow
        
        ' Sort the range for the current email
        ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, ws.Columns.Count)).Sort _
            Key1:=ws.Cells(startRow, startDateCol), _
            Order1:=xlAscending, _
            Header:=xlNo ' Assuming Start Date is in the identified column

        ' Merge cells in the Email ID column for the current email
        If endRow - startRow > 0 Then
            Set rng = ws.Range(ws.Cells(startRow, emailCol), ws.Cells(endRow, emailCol))
            rng.Merge
            rng.HorizontalAlignment = xlCenter
            rng.VerticalAlignment = xlTop
        End If
    Next email

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub