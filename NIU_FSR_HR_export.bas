Sub CopyHRSDetail()
    
    'Copies to new tab for import into FWS Rec MSAccess
    Dim ws          As Worksheet
    'Set ws = ActiveSheet
    Set ws = Worksheets("HRSDetail")
    ws.Activate        'make sure correct file has focus
    Dim startCell   As Range
    Set startCell = ws.Range("E19")
    Dim firstRow    As Long
    Dim lastRow     As Long        'last row in col c
    Dim GroupFirstRow As Long
    Dim GroupLastRow As Long
    Dim groupName   As String
    Dim acctnum     As String
    
    Dim tempDict    As New Dictionary        ' to temporarily hold ID and amounts
    
    'initialize first group start num
    firstRow = startCell.Row
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    GroupFirstRow = startCell.Row
    Dim mycls       As clsAccountGroups
    Set mycls = New clsAccountGroups
    tempDict.RemoveAll
    'iterate over column c. use cell vals to create a 2D array once ubound is determined
    Dim i           As Long
    For i = firstRow To lastRow
        
        If Not ws.Rows(i).Hidden Then
            
            If (Right(ws.Cells(i, 3).Value, 5) <> "Total" And ws.Cells(i, 5).Value <> "") Then
                Debug.Print ws.Cells(i, 5).Value + " " + CStr(ws.Cells(i, 14).Value)
                
                tempDict.Add ws.Cells(i, 5).Value, _
                             CStr(ws.Cells(i, 14).Value)
            End If
            
            If ws.Cells(i, 5).Value = "" Then
                
                If i >= lastRow Then        ' a bit of  cludge to avoid edge cases a end of sheet
                GoTo endsub:
            End If
            If ws.Cells(i, 3).Value <> "" Then
                ' End group: display group name, start and end row
                GroupLastRow = i - 1
                acctnum = Left(ws.Cells(i, 3).Value, 6)
                If Left(acctnum, 5) <> "Grand" Then
                    Debug.Print acctnum & " " & CStr(GroupFirstRow) & " " & CStr(GroupLastRow)   'moved end if
                
                
                GroupFirstRow = getNextVisibleCell(i, ws)
                'set class dictionary to tempdict
                mycls.groupAcctNum = acctnum
                Select Case acctnum
                    Case "644050"        'Set mycls.dict644050 = tempDict   ' dictionaries are byRef, cannot be assigned.  must manually clone
                        Set mycls.dict644050 = New Dictionary
                        CloneDict tempDict, mycls.dict644050, "644050"
                    Case "647200"
                        Set mycls.dict647200 = New Dictionary
                        CloneDict tempDict, mycls.dict647200, "647200"
                    Case "648100"
                        Set mycls.dict648100 = New Dictionary
                        CloneDict tempDict, mycls.dict648100, "648100"
                    Case "648120"
                        Set mycls.dict648120 = New Dictionary
                        CloneDict tempDict, mycls.dict648120, "648120"
                    Case "660130"
                        Set mycls.dict660130 = New Dictionary
                        CloneDict tempDict, mycls.dict660130, "660130"
                    Case Else
                        MsgBox Err.Description
                End Select
                
                tempDict.RemoveAll
                End If
            End If
            
        End If
        
    End If
    
Next i
endsub: MsgBox "export To class complete"
makeArray mycls        ' use the class and its dicts to make an array to be assigned to range

End Sub

Sub CloneDict(ByRef sourceDict As Dictionary, ByRef destDict As Dictionary, acctnum As String)
    
    destDict.Add "name", acctnum
    
    Dim key         As Variant
    For Each key In sourceDict.Keys
        destDict.Add key, sourceDict(key)
    Next
    
End Sub

Function getNextVisibleCell(ByVal i As Long, ByRef ws As Worksheet) As Long
    
    If ws.Cells(i, 3).Value = "Grand Total" Then
        Exit Function
    Else
        While (ws.Rows(i).Hidden Or ws.Cells(i, 3) = "")
            i = i + 1
        Wend
    End If
    
    getNextVisibleCell = i
    
End Function

Sub makeArray(ByRef c As clsAccountGroups)
    ' todo: allow user to select which accounts they do/don't wan't to export
    
    'caclulate size of arr
    Dim l           As Long
    l = c.dict644050.Count - 1 + _
        c.dict647200.Count - 1 + _
        c.dict648100.Count - 1 + _
        c.dict648120.Count - 1 + _
        c.dict660130.Count - 1        'one element of the dict is name, so subtrace one to calc correct size
    
    MsgBox l
    
    Dim FSRArray()  As Variant
    ReDim FSRArray(1 To l, 1 To 3)
    
    Dim i           As Long
    Dim key         As Variant
    i = 1
    For Each key In c.dict644050.Keys
        If key <> "name" Then
            FSRArray(i, 1) = "644050"
            FSRArray(i, 2) = key
            FSRArray(i, 3) = c.dict644050(key)
            i = i + 1
        End If
        
    Next key
    
    For Each key In c.dict647200.Keys
        If key <> "name" Then
            FSRArray(i, 1) = "647200"
            FSRArray(i, 2) = key
            FSRArray(i, 3) = c.dict647200(key)
            i = i + 1
        End If
        
    Next key
    
    For Each key In c.dict648100.Keys
        If key <> "name" Then
            FSRArray(i, 1) = "648100"
            FSRArray(i, 2) = key
            FSRArray(i, 3) = c.dict648100(key)
            i = i + 1
        End If
        
    Next key
    
    For Each key In c.dict648120.Keys
        If key <> "name" Then
            FSRArray(i, 1) = "648120"
            FSRArray(i, 2) = key
            FSRArray(i, 3) = c.dict648120(key)
            i = i + 1
        End If
        
    Next key
    
    For Each key In c.dict660130.Keys
        If key <> "name" Then
            FSRArray(i, 1) = "648130"
            FSRArray(i, 2) = key
            FSRArray(i, 3) = c.dict660130(key)
            i = i + 1
        End If
        
    Next key
    
    addWS_FSR_worksheet (FSRArray)
End Sub

Sub addWS_FSR_worksheet(arr As Variant)
    
    Dim wb          As Workbook
    Set wb = ActiveWorkbook
    Dim WS_FSR      As Worksheet
    
    Set WS_FSR = wb.Worksheets.Add
    WS_FSR.Name = "WS_FSR"
    Dim fsr_range   As Range
    Dim lastRow     As Long, lastCol As Long
    
    lastRow = UBound(arr, 1)
    lastCol = UBound(arr, 2)
    
    Set fsr_range = WS_FSR.Range("A2").Resize(lastRow, lastCol)
    fsr_range.Value = arr
    WS_FSR.Range("A1").Value = "FSR_ACCT"
    WS_FSR.Range("B1").Value = "NAME"
    WS_FSR.Range("C1").Value = "FSR_AMT"
End Sub
