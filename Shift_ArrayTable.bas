Attribute VB_Name = "Module_Shift_ArrayTable"
Option Explicit

' Create a circular buffer writing data from ws Calc array at day index "slot" for all tickers
' Leaves 2 empty rows above (forparameter/column name and slot/date)
' Leaves first column reserved for tickers (when needed)

Sub UpdateHistoryCubeArray()

    Dim wsC As Worksheet
    Dim wsH As Worksheet
    Dim slot As Long
    Dim nTick As Long
    Dim nParam As Long
    Dim destCol As Long

    Set wsC = Sheets("Calc")
    Set wsH = Sheets("Hist")

    slot = wsC.Range("Slot").Value

    nTick = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row - 1
    nParam = wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column - 1

    destCol = 2 + (slot - 1) * nParam

    'Copy matrix directly (no transpose)
    wsH.Cells(3, destCol).Resize(nTick, nParam).Value = _
        wsC.Cells(2, 2).Resize(nTick, nParam).Value

End Sub


' Create a circular buffer writing TabCalc at day index "slot" for all tickers
' Leaves 2 empty rows above (forparameter/column name and slot/date)
' Leaves first column empty (reserved for tickers - when needed)

Public Sub UpdateHistoryCubeTable()

    Const HIST_FIRSTROW As Long = 3     'first data row in Hist
    Const HIST_FIRSTCOL As Long = 2     'column where slot1 begins
    
    Dim wsC As Worksheet
    Dim wsD As Worksheet
    Dim wsH As Worksheet
    Dim lo As ListObject
    
    Dim slot As Long
    Dim nTick As Long
    Dim nParam As Long
    Dim destCol As Long
    
    Dim arr As Variant

    
    Set wsC = ThisWorkbook.Worksheets("Calc")
    Set wsD = ThisWorkbook.Worksheets("DashBoard")
    Set wsH = ThisWorkbook.Worksheets("Hist")
    
    'Table containing current calculations
    Set lo = wsC.ListObjects("TabCalc")
    
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    'Current circular slot
    slot = wsD.Range("Slot").Value
    
    'Table size
    nTick = lo.DataBodyRange.Rows.Count
    nParam = lo.DataBodyRange.Columns.Count
    
    'Destination column in circular cube
    destCol = HIST_FIRSTCOL + (slot - 1) * nParam
    
    'Read table into memory (fast)
    arr = lo.DataBodyRange.Value
    
    'Ensure Hist sheet is large enough
    Call EnsureHistCapacity(wsH, HIST_FIRSTROW, destCol, nTick, nParam)
    
    'Write matrix in one operation
    wsH.Cells(HIST_FIRSTROW, destCol).Resize(nTick, nParam).Value = arr
    
    Debug.Print "History updated | Slot:", slot, _
                "| Rows:", nTick, "| Cols:", nParam

End Sub


Private Sub EnsureHistCapacity(ws As Worksheet, firstRow As Long, _
                               destCol As Long, nTick As Long, nParam As Long)

    Dim lastRowNeeded As Long
    Dim lastColNeeded As Long
    
    lastRowNeeded = firstRow + nTick - 1
    lastColNeeded = destCol + nParam - 1
    
    'Expand rows if needed
    If ws.Rows.Count < lastRowNeeded Then
        ws.Rows(lastRowNeeded).EntireRow.Insert
    End If
    
    'Expand columns if needed
    If ws.Columns.Count < lastColNeeded Then
        ws.Columns(lastColNeeded).EntireColumn.Insert
    End If

End Sub

