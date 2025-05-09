Option Compare Database

Sub changeQry(myParam As String)

Dim qry As DAO.QueryDef
Dim strSQL As String
strSQL = "SELECT Table1.ID,  Table1.fname FROM Table1 WHERE fname like '" & myParam & "';"
Set qry = CurrentDb.QueryDefs("testQry")
qry.SQL = strSQL
qry.Close
DoCmd.OpenQuery "testQry"


End Sub

Sub importFWS()
    Dim db As DAO.Database
    Dim xlApp As New Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim arrRange As Range
    Dim lastrow As Long
    Dim firstrow As Long
    Dim firstcol As Long
    Dim lastcol As Long

    On Error GoTo importError ' Error handler for unexpected errors

    Set db = CurrentDb

    ' Prompt user to select an Excel file
    With Application.FileDialog(3)
        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Exiting."
            GoTo cleanup
        End If
        filePath = .SelectedItems(1)
    End With

    ' Start Excel application
    xlApp.Visible = False   'set to true to debug

    ' Open the selected workbook
    Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)

    ' Reference the specific worksheet, isolate the error
    On Error Resume Next
    Set ws = wb.Worksheets("HRSDetail")
    On Error GoTo importError

    If ws Is Nothing Then
        MsgBox "Worksheet 'HRSDetail' not found."
        GoTo cleanup
    End If

   
   'this import takes some time to run - give user feedback via simple form:
    DoCmd.OpenForm "frmStatus", , , , , acWindowNormal
    Forms("frmStatus").Controls("lblStatus").Caption = "Cleaning up worksheet..."
    DoEvents
    cleanUpWorksheet ws
     ' Wait to ensure workbook is ready workaround. The ungrouping takes a while causing a race condition
    DoEvents
    xlApp.Wait Now + TimeValue("0:00:01")

    ' Define range from C16:Z(last row)
    firstrow = 16
    firstcol = 3
    lastcol = 26
    lastrow = ws.Cells(ws.Rows.Count, firstcol).End(xlUp).Row
    
    Forms("frmStatus").Controls("lblStatus").Caption = "Building Array..."
    DoEvents
    Set arrRange = ws.Range(ws.Cells(firstrow, firstcol), ws.Cells(lastrow, lastcol))
    Debug.Print "Range is: " & arrRange.Address
     Forms("frmStatus").Controls("lblStatus").Caption = "Building Array..."
    DoEvents
   rangeToArr arrRange, xlApp
    Forms("frmStatus").Controls("lblStatus").Caption = "Appending Records to table..."
    DoEvents
    arrayToTable rangeToArr(arrRange, xlApp)
    'MsgBox ("done")
cleanup:
    On Error Resume Next
    Forms("frmStatus").Controls("lblStatus").Caption = "HR Detail Import Complete"
    Forms("frmStatus").Controls("cmdCloseCounterForm").Enabled = True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Set db = Nothing
    Exit Sub

importError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume cleanup
End Sub



Sub cleanUpWorksheet(ws As Object)

    ' Remove subtotals
    On Error Resume Next
    ws.Cells.RemoveSubtotal
    On Error GoTo 0

    ' Ungroup all (rows and columns)
    On Error Resume Next
    ws.Rows.Ungroup
    ws.Columns.Ungroup
    On Error GoTo 0

    ' Show all rows and columns
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False

    
End Sub

Function rangeToArr(ByRef r As Range, a As Excel.Application) As Variant
    Dim rowcount As Long, colcount As Long
    rowcount = r.Rows.Count
    colcount = 5
    
    Dim myArr() As Variant
    ReDim myArr(1 To rowcount, 1 To colcount)
    
    Dim tmpTotal As Double
    Dim currID As String, nextID As String
    Dim currName As String, currAcct As String, currProj As String
    Dim currAmt As Double
    
    Dim i As Long, j As Long
    j = 1
    
    For i = 1 To rowcount
        currID = Trim(CStr(r.Cells(i, 24).Value))
        If currID = "" Then GoTo SkipRow
        
        currName = r.Cells(i, 3)
        currAcct = r.Cells(i, 1)
        currProj = r.Cells(i, 19)
        currAmt = CDbl(r.Cells(i, 12).Value)
        tmpTotal = tmpTotal + currAmt
        
        ' Look ahead to next *non-blank* row for ID
        nextID = ""
        Dim lookAhead As Long
        For lookAhead = i + 1 To rowcount
            nextID = Trim(CStr(r.Cells(lookAhead, 24).Value))
            If nextID <> "" Then Exit For
        Next lookAhead
        
        If currID <> nextID Then
            myArr(j, 1) = currID
            myArr(j, 2) = currName
            myArr(j, 3) = tmpTotal
            myArr(j, 4) = currAcct
            myArr(j, 5) = currProj
            j = j + 1
            tmpTotal = 0
        End If

SkipRow:
    Forms("frmStatus").Controls("txtCounter").Value = i
    Next i

    ' Trim array
    Dim retArray() As Variant
    ReDim retArray(1 To j - 1, 1 To 5)
    
    Dim k As Long, m As Long
    For k = 1 To j - 1
        For m = 1 To 5
            retArray(k, m) = myArr(k, m)
        Next m
    Next k

    rangeToArr = retArray
End Function


Sub arrayToTable(ByRef r As Variant)

    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rs As DAO.Recordset
    

    
    Set rs = db.OpenRecordset("WS_HR", dbOpenDynaset)
    
    ' delete existing records before adding new
    Dim strSQL As String
    strSQL = "delete from " & rs.Name
    db.Execute strSQL, dbFailOnError
    
    Dim i As Long
    For i = 1 To UBound(r, 1)
        rs.AddNew
            rs.Fields(0).Value = r(i, 1)
            rs.Fields(1).Value = r(i, 2)
            rs.Fields(2).Value = r(i, 3)
            rs.Fields(3).Value = r(i, 4)
            rs.Fields(4).Value = r(i, 5)
        rs.Update
        Forms("frmStatus").Controls("txtCounter").Value = i ' "added row " & i & " to table"
        Debug.Print "added row " & i; " to table1"
    Next i
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


Sub import_WS_PS()
  Dim db As DAO.Database
  Set db = CurrentDb
  Dim rs As DAO.Recordset
  
  db.Execute "Delete from WS_PS", dbFailOnError
  
  
  
  'Dim rs As DAO.Recordset
  'Set rs = db.OpenRecordset("WS_PS", dbOpenDynaset)
  Dim filePath As String
  
      With Application.FileDialog(3)
        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Exiting."
            GoTo cleanup
        End If
        filePath = .SelectedItems(1)
    End With
  
  
  DoCmd.TransferSpreadsheet _
    TransferType:=acImport, _
    SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
    TableName:="WS_PS", _
    filename:=filePath, _
    HasFieldNames:=True

  
cleanup:
  Set db = Nothing
  Set rs = Nothing
  MsgBox ("import complete")
End Sub


Function getSelectedAccounts() As String
    
    Dim varItem As Variant
    Dim strList As String

    With [Forms]![FWS Recon Util]![lstAccts]
        For Each varItem In .ItemsSelected
            strList = strList & "'" & .ItemData(varItem) & "',"
        Next varItem
    End With

    ' Remove the trailing comma
    If Len(strList) > 0 Then
        strList = Left(strList, Len(strList) - 1)
    End If

    getSelectedAccounts = strList
    Debug.Print strList

End Function

Sub updateSubform(sf As Form)
    Dim q As String
 'q = "SELECT WS_HR.sName, WS_HR.EmplId, Sum(WS_HR.Amt) AS SumOfAmt, " & _
    "Sum(WS_PS.[PS Earnings]) AS [SumOfPS Earnings], " & _
    "IIf(Sum(WS_HR.Amt)=Sum(WS_PS.[PS Earnings]),True,False) AS Balanced " & _
    "FROM WS_HR INNER JOIN WS_PS ON (WS_HR.Acct = WS_PS.AcctCD) " & _
    "AND (WS_HR.Project = WS_PS.ProjectID) AND (WS_HR.EmplId = WS_PS.ID) " & _
    "WHERE WS_HR.Acct IN (" & getSelectedAccounts() & ") " & _
    "AND WS_HR.Project = 'G7B69911' " & _
    "GROUP BY WS_HR.sName, WS_HR.EmplId " & _
    "ORDER BY WS_HR.sName;"
    
   'If sf.Name = "WS_HR subform" Then
   
   'need to add project id as a filter in the LOJs.  or add to disply field to use as filer in
   q = "SELECT IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.EmplID, IDS_WITH_MISMATCHED_ACCT_SUM.Acct, Sum(IIf(Source='HR',Amt,0)) AS HR_Sum, Sum(IIf(Source='PS',Amt,0)) AS PS_Sum " & _
        "FROM IDS_WITH_MISMATCHED_ACCT_SUM " & _
        "WHERE (((IDS_WITH_MISMATCHED_ACCT_SUM.Acct)<>'660130'))  AND 1=1 " & _
        "AND IDS_WITH_MISMATCHED_ACCT_SUM.Acct IN(" & getSelectedAccounts() & ")" & _
        "GROUP BY IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.EmplID, IDS_WITH_MISMATCHED_ACCT_SUM.Acct " & _
        "HAVING (((Sum(IIf([Source]='HR',[Amt],0))+Sum(IIf([Source]='PS',[Amt],0)))>0) AND ((Sum(IIf([Source]='HR',[Amt],0)))<>Sum(IIf([Source]='PS',[Amt],0)))) " & _
        " ORDER BY IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.Acct;"
    'ElseIf sf.Name = "WS_PS Subform" Then
    'q = "SELECT IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.EmplID, IDS_WITH_MISMATCHED_ACCT_SUM.Acct, Sum(IIf(Source='HR',Amt,0)) AS HR_Sum, Sum(IIf(Source='PS',Amt,0)) AS PS_Sum " & _
        "FROM IDS_WITH_MISMATCHED_ACCT_SUM " & _
        "WHERE (((IDS_WITH_MISMATCHED_ACCT_SUM.Acct)<>'660130')) " & _
        "AND IDS_WITH_MISMATCHED_ACCT_SUM.Acct IN(" & getSelectedAccounts() & ")" & _
        "GROUP BY IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.EmplID, IDS_WITH_MISMATCHED_ACCT_SUM.Acct " & _
        "HAVING (((Sum(IIf([Source]='PS',[Amt],0))+Sum(IIf([Source]='PS',[Amt],0)))>0) AND ((Sum(IIf([Source]='HR',[Amt],0)))<>Sum(IIf([Source]='PS',[Amt],0)))) " & _
        " ORDER BY IDS_WITH_MISMATCHED_ACCT_SUM.sName, IDS_WITH_MISMATCHED_ACCT_SUM.Acct;"
    'Else
     '    q = "SELECT WS_HR.sName, WS_HR.EmplId, Sum(WS_HR.Amt) AS SumOfAmt, " & _
    "Sum(WS_PS.[PS Earnings]) AS [SumOfPS Earnings], " & _
    "IIf(Sum(WS_HR.Amt)=Sum(WS_PS.[PS Earnings]),True,False) AS Balanced " & _
    "FROM WS_HR INNER JOIN WS_PS ON (WS_HR.Acct = WS_PS.AcctCD) " & _
    "AND (WS_HR.Project = WS_PS.ProjectID) AND (WS_HR.EmplId = WS_PS.ID) " & _
    "WHERE WS_HR.Acct IN (" & getSelectedAccounts() & ") " & _
    "AND WS_HR.Project = 'G7B69911' " & _
    "GROUP BY WS_HR.sName, WS_HR.EmplId " & _
    "ORDER BY WS_HR.sName;"
    
'End If
    
Debug.Print q

    sf.RecordSource = q
    sf.Requery
End Sub


