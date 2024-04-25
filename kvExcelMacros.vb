Option Explicit

Sub doSeperateBom()
    'comment to cause a conflict 2 windows
    ' insert a comment for git commit test
    ' inserts a sperater row between BOMs on BOM import.  or any tabular data with groupings for that matter.  user picks the col to base the sepration on

    Dim column As Integer
column = InputBox("which col?", "which col")
separateBOMs (column)

End Sub



Sub separateBOMs(changeCol As Long)

    Dim lastRow As Long
    Dim currentVal As String
    Dim nextVal As String
    Dim currVal As String
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim nextRow As Long
    Dim insertRow As Long
    Dim headerCol As Long
    
    headerCol = Cells(1, Columns.Count).End(xlToLeft).column
    
    ' Determine the last row with data in the specified column
    lastRow = Cells(Rows.Count, changeCol).End(xlUp).Row
    
    ' Initialize current value with the value in the first row
    currentVal = Cells(2, changeCol).Value
    
    ' Loop through each row starting from the second row
    'For i = 2 To lastRow
    i = 2
    Do While i <= lastRow
        ' Read the value in the next row
       Cells(i, changeCol).Activate
        nextVal = Cells(i + 1, changeCol).Value
        currVal = Cells(i, changeCol).Value
        ' Check if the value changes
        'Debug.Print i
        'Debug.Print nextVal
        If currentVal <> nextVal Then
        ' And Not (currentVal <> "" And nextVal <> "") Then ' breaks my brain
            ' Define the row numbers for the current range
            nextRow = i + 1 ' Row after4 the change
            endRow = i ' Row where the change occurs
            
            ' Update current value to the new value
            If nextVal <> "" And currentVal <> "" Then
                Rows(nextRow).Insert
                     lastRow = lastRow + 1
                        If changeCol = 1 Then
                            Dim cell As Range
                            ' color the cells in the inserted row red.  But do not color the entire row. only those that correspond to the used range colomn number, which will be the last used header row (happens to be X in this particular data, but could change
                            For Each cell In Range(Cells(nextRow, 1), Cells(nextRow, headerCol))
                            cell.Interior.Color = RGB(255, 0, 0) ' Red color
                            Next cell
                            Else
                            
                             Dim cell2 As Range
                            ' color the cells in the inserted row yellow if not based on col1.  But do not color the entire row. only those that correspond to the used range colomn number, which will be the last used header row (happens to be X in this particular data, but could change
                            For Each cell2 In Range(Cells(nextRow, 1), Cells(nextRow, headerCol))
                            cell2.Interior.Color = RGB(255, 255, 153) ' Yeller color
                            Next cell2
                        End If
            End If
        End If
   ' Next i
   currentVal = nextVal
  
   i = i + 1
   Loop
    
    ' Insert a new row after processing the current range
    'insertRow = endRow + 1

End Sub

Sub ConsolidateTabs()
    Dim ws As Worksheet
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim lastColumn As Long
    Dim copyRange As Range
    Dim headerRow As Range
    Dim fieldNames() As String
    Dim foundSheet As Boolean
    foundSheet = False

    ' unsophistacated merge of worksheets into a single merged worksheet.  just looks for field name (user provided) in A1 and merged if found, skips the sheet if not.
    ' Set the main worksheet where all data will be consolidated
    
  On Error Resume Next
    Set wsMain = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
 On Error GoTo 0
    
    If Not wsMain Is Nothing Then
        Dim tabName As String
        
        wsMain.Name = GetUniqueSheetName("MERGED")
    End If
    'check if first cell is fieldname - to identify sheets needing copied
    Dim field1 As String
    field1 = InputBox("name  of first field", "identify field 1")
    'let user specify number of cols
    lastColumn = InputBox("enter number of columns to copy", "num cols?")
    
    ' Loop through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets
        ' Check if the current worksheet is the main worksheet
        If ws.Name <> wsMain.Name Then
            ' Check if the first cell is "ItemName"
            If foundSheet = False Then
                If ws.Cells(1, 1).Value = field1 Then
                    'define header row field names
                    Dim i As Long
                    ReDim fieldNames(1 To lastColumn)
                    For i = 1 To lastColumn
                        fieldNames(i) = ws.Cells(1, i).Value
                    Next i
                End If
                'set wsmain header row
                For i = 1 To UBound(fieldNames)
                    wsMain.Cells(1, i) = fieldNames(i)
                Next i
                foundSheet = True
            End If
            
            ' we found a sheet and set a header row.  now populate we merge to main
            'get last row so we know where to put the data:
            Dim j As Long
            Dim colMapping As Object
            Set colMapping = CreateObject("Scripting.Dictionary")
            Dim wsMainColumn As Long
            'get num cols in ws
            Dim rowLastColumn As Long
           
            rowLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

            
            For j = 1 To rowLastColumn
                Dim columnLastRow As Long
                columnLastRow = wsMain.Cells(wsMain.Rows.Count, j).End(xlUp).Row
                ' Update lastRow if the current column's last row is greater
                If columnLastRow > lastRow Then
                    lastRow = columnLastRow
                End If
                            
            
                'while we are here, lets find the colindex where ws data goes
                wsMainColumn = 0
                 'For j = 1 To rowLastColumn
                     For i = 1 To UBound(fieldNames)
                         If StrComp(Trim(ws.Cells(1, j).Value), Trim(fieldNames(i))) = 0 Then
                             'then col from ws corresponds to col j from wsmain
                             wsMainColumn = i
                             colMapping.Add ws.Cells(1, j).Value, wsMainColumn
                             Exit For
                         End If
                     Next i
                 
                 Next j
                 'find matching column then put colum data at last row
                 
                For j = 1 To rowLastColumn
                     wsMainColumn = colMapping(ws.Cells(1, j).Value)
                     ' Copy the data from ws to wsMain at lastRow, but only if the mapping exists
                     If wsMainColumn > 0 Then
                         ws.Range(ws.Cells(1, j), ws.Cells(ws.Rows.Count, j).End(xlUp)).Copy _
                         Destination:=wsMain.Cells(lastRow + 1, wsMainColumn)
                     End If
                Next j
        End If
    Next ws
    
    If Not foundSheet Then
        MsgBox ("no sheet with that field name found. quitting...")
    End If
    
    wsMain.Columns.AutoFit
End Sub

Function GetUniqueSheetName(baseName As String) As String
    Dim counter As Integer
    Dim sheetName As String
    
    ' Initialize counter
    counter = 1
    
    ' Loop until a unique name is found
    Do
        ' Construct potential sheet name
        If counter > 1 Then
            sheetName = baseName & " (" & counter & ")"
        Else
            sheetName = baseName
        End If
        
        ' Check if sheet name already exists
        If SheetExists(sheetName) Then
            counter = counter + 1
        Else
            ' Unique name found
            GetUniqueSheetName = sheetName
            Exit Function
        End If
    Loop
End Function


Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    ' Loop through all worksheets
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    
    ' If the loop completes without finding the sheet name, it doesn't exist
    SheetExists = False
End Function


Sub ConsolidateTabsFields()
    Dim ws As Worksheet
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim lastColumn As Long
    Dim copyRange As Range
    Dim headerRow As Range
    Dim fieldNames() As String
    Dim foundSheet As Boolean
    foundSheet = False
    Dim setHeaderRow As Boolean
    setHeaderRow = False
    
    
    
    ' Set the main worksheet where all data will be consolidated
    
  On Error Resume Next
    Set wsMain = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
 On Error GoTo 0
    
    If Not wsMain Is Nothing Then
        Dim tabName As String
        
        wsMain.Name = GetUniqueSheetName("MERGED")
    End If
    
    Dim fields As String
    fields = InputBox("enter field list csv")
    Dim fieldArray() As String
    fieldArray = Split(fields, ",")
    lastColumn = UBound(fieldArray) + 1
    
    'write header row
        Dim k As Long
        If setHeaderRow = False Then
            For k = 0 To lastColumn - 1
                wsMain.Cells(1, k + 1) = fieldArray(k)
            Next k
        
        setHeaderRow = True
        End If

                        
    
    
    
    ' Loop through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets
        ' Check if the current worksheet is the main worksheet
        If ws.Name <> wsMain.Name Then
             
            ' we found a sheet and set a header row.  now populate we merge to main
            'get last row so we know where to put the data:
            Dim j As Long
            Dim colMapping As Object
            Set colMapping = CreateObject("Scripting.Dictionary")
            Dim wsMainColumn As Long
            'get num cols in ws
            Dim rowLastColumn As Long
           
            rowLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column

            
            For j = 1 To rowLastColumn
                Dim columnLastRow As Long
                columnLastRow = wsMain.Cells(wsMain.Rows.Count, j).End(xlUp).Row
                ' Update lastRow if the current column's last row is greater
                If columnLastRow > lastRow Then
                    lastRow = columnLastRow
                End If
                            
            
                'while we are here, lets find the colindex where ws data goes
                wsMainColumn = 0
                 'For j = 1 To rowLastColumn
                 Dim i As Long
                     For i = 0 To lastColumn - 1
                         If StrComp(Trim(ws.Cells(1, j).Value), Trim(fieldArray(i))) = 0 Then
                             'then col from ws corresponds to col j from wsmain
                             wsMainColumn = i + 1
                             colMapping.Add ws.Cells(1, j).Value, wsMainColumn
                             Exit For
                         End If
                     Next i
                 
                 Next j
                 'find matching column then put colum data at last row
                 
                For j = 1 To rowLastColumn
                
                     wsMainColumn = colMapping(ws.Cells(1, j).Value)
                     ' Copy the data from ws to wsMain at lastRow, but only if the mapping exists
                     If wsMainColumn > 0 Then
                         ws.Range(ws.Cells(1, j), ws.Cells(ws.Rows.Count, j).End(xlUp)).Copy _
                         Destination:=wsMain.Cells(lastRow + 1, wsMainColumn)
                     End If
                Next j
        End If
    Next ws
    
End Sub
Sub DumbMerge()
    Dim ws As Worksheet
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim lastColumn As Long
    Dim copyRange As Range

    'very unsophisicated merge of sheets - merges all sheets, regardless of fields names or order, into a merged sheet

    ' Create a new worksheet to consolidate all data
    Set wsMain = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    wsMain.Name = "ConsolidatedData"
    destRow = 1 ' Start from the first row in the destination sheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets
        ' Skip the new consolidated worksheet
        If ws.Name <> wsMain.Name Then
            ' Find the last used row in the current worksheet
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ' Find the last used column in the current worksheet
            lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            ' Define the range to copy (excluding headers)
            Set copyRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
            
            ' Copy data to the destination sheet
            copyRange.Copy Destination:=wsMain.Cells(destRow, 1)
            
            ' Update destination row for the next data set
            destRow = destRow + copyRange.Rows.Count
        End If
    Next ws
    
    MsgBox "Data has been consolidated successfully!", vbInformation
End Sub

Sub mergeByIndex()

' merges columns selected by user to a new merged sheet. user can enter csv or ranges. e.g. 1,2,3-5,6

    Dim ws As Worksheet
    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim lastMainCol As Long
    Dim copyRange As Range
    Dim fieldValues As Variant
    Dim fields As String
    Dim c As Variant
    Dim clong As Long
    Dim destCol As Long ' Destination column in wsMain

    ' First, add the new merged sheet
    Set wsMain = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    wsMain.Name = "MERGED"

    ' Get the field indices from the user
    fields = InputBox("Enter fields by column index separated by comma", "Field Index List")
   ' fieldValues = Split(fields, ",")
   fieldValues = parseInput(fields)
    lastMainCol = UBound(fieldValues) + 1

    ' Now iterate over each sheet and copy to merged sheet beginning at lastrow
    lastRow = 1
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "MERGED" Then
            ' Find last row used in "merged"
            For Each c In fieldValues
                clong = CLng(c) ' Get the field index as number
                ' Set destination column in wsMain
                destCol = destCol + 1
                ' Copy to the next column in merged
                ws.Cells(1, clong).Resize(ws.Cells(ws.Rows.Count, clong).End(xlUp).Row).Copy _
                    Destination:=wsMain.Cells(lastRow, destCol)
            Next c
            ' Increment lastRow for the next set of data
            lastRow = lastRow + ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ' Reset destination column for the next worksheet
            destCol = 0
        End If
    Next ws

End Sub


Function parseInput(ip As String) As String()
    Dim result() As String ' Declare an array to hold the result
    Dim arrInput() As String ' Declare an array to hold the input parts
    
    ' Split the input string by comma to get individual parts
    arrInput = Split(ip, ",")
    
    Dim nextIndex As Long ' Variable to track the next available index for appending elements
    
    Dim i As Long
    For i = 0 To UBound(arrInput)
        Dim subArr() As String
        Dim subStr As String
        If InStr(arrInput(i), "-") Then
            ' If a range is detected, split it into sub-array
            subStr = arrInput(i)
            subArr = makeSubArr(subStr) ' Assuming makeSubArr function returns an array of strings
            
            ' Append each element of subArr to result array
            Dim j As Long
            For j = 0 To UBound(subArr)
                ReDim Preserve result(0 To nextIndex) ' Resize result array
                result(nextIndex) = subArr(j) ' Append the element to result array
          
                    nextIndex = nextIndex + 1 ' Increment next available index
              
            Next j
        Else
            ' If it's a single value, directly append it to result array
            ReDim Preserve result(0 To nextIndex) ' Resize result array
            result(nextIndex) = arrInput(i) ' Append the element to result array
            nextIndex = nextIndex + 1 ' Increment next available index
        End If
    Next i
    
    parseInput = result ' Return the result array
End Function


Function makeSubArr(subStr As String) As String()
    
    Dim subArr() As String
    subArr = Split(subStr, "-")
    Dim startNum As Long
    startNum = subArr(0)
    
    Dim numNums As Long
    numNums = (subArr(1) - subArr(0)) + 1
    
    Dim retArr() As String
    ReDim retArr(numNums - 1)
    
    Dim i As Long
    
    For i = 0 To numNums - 1
     
       
      retArr(i) = startNum
      startNum = startNum + 1
      
    Next
makeSubArr = retArr

End Function

Function fixDate() As String

Dim currDate As String
Dim retDate As String




' currDate = ActiveCell.Offset(0, -2).Value ' value converts to data object .  use text:
currDate = ActiveCell.Offset(0, -1).Text

'determine if euro or  american

Dim sDate() As String
Dim pDate() As String

sDate = Split(currDate, "/")
If (sDate(0) > 12) Then
    'prob ok but to be safe
    'if not the first row, check previous row to guess if next in chronological order. this is just a heuristic

     pDate = Split(ActiveCell.Offset(-1, 0).Value, "/")
     If UBound(pDate) > 0 Then
         If (Abs((CLng(sDate(1)) - CLng(pDate(0)))) <= 1) Then
            'not in order - reformat
            retDate = sDate(1) & "/" & sDate(0) & "/" & sDate(2)
        End If
    End If
Else
    retDate = currDate
End If

fixDate = retDate


End Function

Sub doFixDate()

'insert col to hold new date
Dim scol As String
scol = InputBox("which column contains the dates", "pick a col to evaluate")
Dim lcol As Long
lcol = CLng(scol)

Columns(lcol + 1).Insert

Cells(1, lcol + 1).Select
ActiveCell.Value = "formatted date"
ActiveCell.Offset(1, 0).Select
    
    Debug.Print fixDate
    While ActiveCell.Offset(0, -1) <> ""
        ActiveCell.Value = "'" & fixDate
    
        ActiveCell.Offset(1, 0).Select
    'Debug.Print ActiveCell.Address
    Wend
    

End Sub

