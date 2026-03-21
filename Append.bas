Attribute VB_Name = "Module1_Append_Trades"
Option Explicit

'========================================================
' MAIN PROCEDURE
'========================================================
Public Sub Append_ToTabReg_AndArchive()

    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim wsE As Worksheet, wsR As Worksheet
    Set wsE = wb.Worksheets("Entries")
    Set wsR = wb.Worksheets("Reg")

    Dim rngTransaction As Range, rngTranParam As Range
    Dim rngTradeArchive As Range, rngEntryList As Range

    Set rngTransaction = wsE.Range("Transaction")   '3 column names of TabReg
    Set rngTranParam = wsE.Range("TranParam")       '3 numeric parameters
    Set rngTradeArchive = wsE.Range("TradeArchive") 'FIXED 4-column range
    Set rngEntryList = wsE.Range("EntryList")

    Dim loReg As ListObject
    Set loReg = wsR.ListObjects("TabReg")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    '--------------------------------------------
    ' Load arrays
    '--------------------------------------------
    Dim arrTx As Variant, arrPr As Variant
    arrTx = rngTransaction.Value2
    arrPr = rngTranParam.Value2

    Dim maxRows As Long
    maxRows = UBound(arrPr, 1)

    '--------------------------------------------
    ' Determine N (non-zero rows in TranParam)
    '--------------------------------------------
    Dim idx() As Long
    ReDim idx(1 To maxRows)

    Dim i As Long, n As Long
    n = 0

    For i = 1 To maxRows
        If IsNonZeroRow3(arrPr, i) Then
            n = n + 1
            idx(n) = i
        End If
    Next i

    If n = 0 Then
        rngEntryList.ClearContents
        GoTo CleanOK
    End If

    '--------------------------------------------
    ' Shift TradeArchive internally (4 columns only)
    '--------------------------------------------
    ShiftTradeArchiveDown_Fixed4 rngTradeArchive, n

    '--------------------------------------------
    ' Fill first N rows of TradeArchive
    ' col1 = TranParam col1
    ' col2 = Transaction col1
    ' col3 = TranParam col2
    ' col4 = TranParam col3
    '--------------------------------------------
    Dim r As Long, rowSrc As Long

    For r = 1 To n
        rowSrc = idx(r)

        rngTradeArchive.Cells(r, 1).Value2 = arrPr(rowSrc, 1)
        rngTradeArchive.Cells(r, 2).Value2 = Trim$(CStr(arrTx(rowSrc, 1)))
        rngTradeArchive.Cells(r, 3).Value2 = arrPr(rowSrc, 2)
        rngTradeArchive.Cells(r, 4).Value2 = arrPr(rowSrc, 3)
    Next r

    '--------------------------------------------
    ' Append values into TabReg
    '--------------------------------------------
    AppendToTabReg loReg, arrTx, arrPr, idx, n

    '--------------------------------------------
    ' Clear EntryList only
    '--------------------------------------------
    rngEntryList.ClearContents

CleanOK:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Post_Transactions failed:" & vbCrLf & _
           "Err " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


'========================================================
' SHIFT ARCHIVE (FIXED 4 COLUMNS ONLY)
'========================================================
Private Sub ShiftTradeArchiveDown_Fixed4(ByVal rngArchive As Range, ByVal n As Long)

    If n <= 0 Then Exit Sub

    Dim rowsCount As Long
    rowsCount = rngArchive.Rows.Count

    If rngArchive.Columns.Count <> 4 Then
        Err.Raise vbObjectError + 501, , _
        "TradeArchive must be exactly 4 columns."
    End If

    If n >= rowsCount Then
        rngArchive.ClearContents
        Exit Sub
    End If

    Dim arrIn As Variant, arrOut As Variant
    arrIn = rngArchive.Value2
    arrOut = arrIn

    Dim r As Long, c As Long

    'Move bottom-up
    For r = rowsCount To 1 Step -1
        If r + n <= rowsCount Then
            For c = 1 To 4
                arrOut(r + n, c) = arrIn(r, c)
            Next c
        End If
    Next r

    'Clear top N rows
    For r = 1 To n
        For c = 1 To 4
            arrOut(r, c) = vbNullString
        Next c
    Next r

    rngArchive.Value2 = arrOut
End Sub


'========================================================
' APPEND VALUES INTO TABREG
'========================================================
Private Sub AppendToTabReg(ByVal lo As ListObject, _
                           ByVal arrTx As Variant, _
                           ByVal arrPr As Variant, _
                           idx() As Long, _
                           ByVal n As Long)

    Dim r As Long, i As Long
    Dim name1 As String, name2 As String, name3 As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant

    If lo.ListRows.Count = 0 Then lo.ListRows.Add

    For r = 1 To n
        i = idx(r)

        name1 = Trim$(CStr(arrTx(i, 1)))
        name2 = Trim$(CStr(arrTx(i, 2)))
        name3 = Trim$(CStr(arrTx(i, 3)))

        v1 = arrPr(i, 1)
        v2 = arrPr(i, 2)
        v3 = arrPr(i, 3)

        RequireColumn lo, name1
        RequireColumn lo, name2
        RequireColumn lo, name3

        AppendValueUnderLastNonZero lo, name1, v1
        AppendValueUnderLastNonZero lo, name2, v2
        AppendValueUnderLastNonZero lo, name3, v3
    Next r
End Sub


Private Sub RequireColumn(ByVal lo As ListObject, ByVal colName As String)
    If Len(colName) = 0 Then
        Err.Raise vbObjectError + 601, , "Empty TabReg column name."
    End If
    On Error GoTo NotFound
    Dim tmp As ListColumn
    Set tmp = lo.ListColumns(colName)
    Exit Sub
NotFound:
    Err.Raise vbObjectError + 602, , _
    "TabReg missing column: " & colName
End Sub


Private Sub AppendValueUnderLastNonZero(ByVal lo As ListObject, _
                                        ByVal colName As String, _
                                        ByVal newVal As Variant)

    Dim lc As ListColumn
    Set lc = lo.ListColumns(colName)

    If lo.ListRows.Count = 0 Then lo.ListRows.Add

    Dim rngCol As Range
    Set rngCol = lc.DataBodyRange

    Dim lastIdx As Long
    lastIdx = LastNonZeroIndex(rngCol)

    Dim targetIdx As Long
    targetIdx = lastIdx + 1
    If targetIdx < 1 Then targetIdx = 1

    Do While targetIdx > lo.ListRows.Count
        lo.ListRows.Add
        Set rngCol = lc.DataBodyRange
    Loop

    rngCol.Cells(targetIdx, 1).Value2 = newVal
End Sub


Private Function LastNonZeroIndex(ByVal rngCol As Range) As Long

    Dim n As Long: n = rngCol.Rows.Count
    Dim k As Long, v As Variant

    For k = n To 1 Step -1
        v = rngCol.Cells(k, 1).Value2

        If IsNumeric(v) Then
            If CDbl(v) <> 0# Then
                LastNonZeroIndex = k
                Exit Function
            End If
        End If
    Next k

    LastNonZeroIndex = 0
End Function


Private Function IsNonZeroRow3(ByVal arr As Variant, ByVal i As Long) As Boolean

    Dim j As Long
    For j = 1 To 3
        If IsNumeric(arr(i, j)) Then
            If CDbl(arr(i, j)) <> 0# Then
                IsNonZeroRow3 = True
                Exit Function
            End If
        End If
    Next j

End Function
