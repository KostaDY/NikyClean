Attribute VB_Name = "Module2_Shift_Summary"
Option Explicit

Public Sub All_DataTablesDailyShift()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim calcMode As XlCalculation

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Worksheets("DataShift")
    ws.Range("LastShift").Value = Now

    For Each lo In ws.ListObjects
        ShiftTableLeft_KeepLastFormula lo
    Next lo

CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculate

End Sub
Private Sub ShiftTableLeft_KeepLastFormula(ByVal lo As ListObject)

    Dim db As Range
    Dim src As Variant
    Dim colArr() As Variant
    Dim r As Long, c As Long
    Dim nRows As Long, nCols As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set db = lo.DataBodyRange
    nRows = db.Rows.Count
    nCols = db.Columns.Count
    If nCols < 2 Then Exit Sub

    ' Read all values once
    src = db.Value

    ' Shift column-by-column
    For c = 1 To nCols - 1
        
        ' Prepare one-column array
        ReDim colArr(1 To nRows, 1 To 1)
        
        For r = 1 To nRows
            colArr(r, 1) = src(r, c + 1)
        Next r
        
        ' Write this column only
        db.Columns(c).Value = colArr
        
    Next c

End Sub
