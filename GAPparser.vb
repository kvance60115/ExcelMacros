Sub ImportFixedWidthData()
    Dim dataFilePath As String
    Dim fieldLengthsFilePath As String
    Dim fileDialog As fileDialog
    Dim fieldLengths As Variant
    Dim data As String
    Dim line As String
    Dim startPos As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ws As Worksheet
    Dim currentRow As Long
    
    ' Prompt for the data file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.Title = "Select the Data File"
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Text Files", "*.txt", 1
    If fileDialog.Show = -1 Then
        dataFilePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "No file selected. Exiting..."
        Exit Sub
    End If
   
    ' Prompt for the field lengths file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.Title = "Select the Field Lengths CSV File"
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "CSV Files", "*.csv", 1
    If fileDialog.Show = -1 Then
        fieldLengthsFilePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "No file selected. Exiting..."
        Exit Sub
    End If
    
    ' Read the field lengths from the CSV file
    fieldLengths = ReadFieldLengths(fieldLengthsFilePath)
    If IsEmpty(fieldLengths) Then
        MsgBox "Failed to read field lengths. Exiting..."
        Exit Sub
    End If
    
    ' Read and process the data file
    Open dataFilePath For Input As #1
    Set ws = ThisWorkbook.Sheets(1) ' Change to your desired sheet
    currentRow = 1
    Do While Not EOF(1)
        Line Input #1, line
        startPos = 1
        j = 1
        Dim thisLength As Long
        For i = LBound(fieldLengths) To UBound(fieldLengths) - 1
            thisLength = Abs(startPos - fieldLengths(i + 1)) + 1
            ws.Cells(currentRow, j).Value = Mid(line, startPos, thisLength)
            startPos = startPos + thisLength
            j = j + 1
            ' assert j < 990
        Next i
        currentRow = currentRow + 1
    Loop
    Close #1
    
    MsgBox "Data import complete."
End Sub

Function ReadFieldLengths(filePath As String) As Variant
    Dim fieldLengths() As Integer
    Dim line As String
    Dim lengths() As String
    Dim i As Integer
    
    Open filePath For Input As #1
    If Not EOF(1) Then
        Line Input #1, line
        lengths = Split(line, ",")
        ReDim fieldLengths(LBound(lengths) To UBound(lengths))
        For i = LBound(lengths) To UBound(lengths)
            fieldLengths(i) = CInt(lengths(i))
        Next i
    End If
    Close #1
    
    ReadFieldLengths = fieldLengths
End Function
