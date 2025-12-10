VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private oldValue As Variant

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   If testFlag Then Debug.Print "Entered: " & "Private Sub Worksheet_SelectionChange ws NewTrader"

    Dim tbl As ListObject
    Dim colClickTick As ListColumn
    Dim rngClickTick As Range
    Dim rngAggrUps As Range
    Dim rngAggrDowns As Range
    Dim rngChangeMax As Range
    Dim rngChangeMin As Range
    Dim rngSaldoUps As Range
    Dim rngSaldoDowns As Range
    'Dim rngFree As Range
 ' _________________________________________ Select rowcode
 
        If Target.CountLarge <> 1 Then Exit Sub
        
        Application.EnableEvents = False
        
        ' --- SelRowNum ---
        On Error Resume Next
        ThisWorkbook.Names("SelRowNum").Delete
        ThisWorkbook.Names.Add Name:="SelRowNum", RefersTo:="=" & Target.Row
        
        ' --- SelRow ---
        On Error Resume Next
        ThisWorkbook.Names("SelRow").Delete
        ThisWorkbook.Names.Add Name:="SelRow", RefersTo:="=" & Target.Row
        
        Application.EnableEvents = True
         
         On Error GoTo CleanExit         '  Check for sticky error

    ' Reference the table and column
    Set tbl = Me.ListObjects("Tab_NewTRADER")
    Set colClickTick = tbl.ListColumns("Show")
    Set rngClickTick = colClickTick.DataBodyRange
    Set rngAggrUps = Me.Range("AggrChangeUps")
    Set rngAggrDowns = Me.Range("AggrChangeDowns")
    Set rngChangeMax = Me.Range("ChangeMax")
    Set rngChangeMin = Me.Range("ChangeMin")
    Set rngSaldoUps = Me.Range("SaldoMax")
    Set rngSaldoDowns = Me.Range("SaldoMin")
    ' Check if the clicked cell is in one of the clickable column
    If Not Intersect(Target, rngClickTick) Is Nothing Or _
        Not Intersect(Target, rngAggrUps) Is Nothing Or _
        Not Intersect(Target, rngAggrDowns) Is Nothing Or _
        Not Intersect(Target, rngChangeMax) Is Nothing Or _
        Not Intersect(Target, rngChangeMin) Is Nothing Or _
        Not Intersect(Target, rngSaldoUps) Is Nothing Or _
        Not Intersect(Target, rngSaldoDowns) Is Nothing Then
        If Target.Cells.Count = 1 Then
            ThisWorkbook.Worksheets("ID").Range("IDTICKER").Value = Target.Value
        End If
    End If
    Exit Sub
CleanExit:
    MsgBox "Exit with error from Trader selection change"
End Sub

