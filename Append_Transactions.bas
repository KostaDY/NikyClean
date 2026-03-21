Attribute VB_Name = "ModuleFIFO"
Option Explicit

Public Sub RunStrategyFIFO()
    Dim originalCalc As XlCalculation
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, j As Long, n As Long
    Dim colNumber As Long, colPrice As Long
    Dim colRemain As Long, colPnL As Long

    Dim Number() As Double, Price() As Double
    Dim Remaining() As Double, PnL() As Double
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.ScreenUpdating = False
    On Error GoTo CleanFail
    'If testFlag Then Debug.Print "Entered: " & "Sub RunInventoryStrategyFIFOSILENT()"

    ' Store original calculation mode and turn off UI updates
    originalCalc = Application.Calculation

    Application.Calculation = xlCalculationManual

    ' === Begin core FIFO logic ===
    Set ws = ThisWorkbook.Sheets("ID")
    Set tbl = ws.ListObjects("TabFIFO")
    n = tbl.ListRows.Count
    If n = 0 Then GoTo CleanExit

    colNumber = tbl.ListColumns("Number").Index
    colPrice = tbl.ListColumns("Price").Index
    colRemain = tbl.ListColumns("Remaining").Index
    colPnL = tbl.ListColumns("Profit/Loss").Index

    ReDim Number(1 To n)
    ReDim Price(1 To n)
    ReDim Remaining(1 To n)
    ReDim PnL(1 To n)

    For i = 1 To n
        Number(i) = tbl.DataBodyRange.Cells(i, colNumber).Value2
        Price(i) = tbl.DataBodyRange.Cells(i, colPrice).Value2
        Remaining(i) = 0
        PnL(i) = 0
    Next i

    For i = 1 To n
        If Price(i) > 0 Then ' sale
            Dim demand As Double: demand = Number(i)
            Dim salePrice As Double: salePrice = Price(i)
            Dim totalCost As Double: totalCost = 0
            Dim matchedQty As Double: matchedQty = 0

            For j = 1 To i - 1
                If Price(j) < 0 Then ' purchase
                    Dim available As Double
                    available = Number(j) - Remaining(j)
                    If available <= 0 Then GoTo NextPurchase

                    Dim qtyUsed As Double
                    qtyUsed = Application.Min(demand, available)
                    demand = demand - qtyUsed
                    Remaining(j) = Remaining(j) + qtyUsed

                    totalCost = totalCost + qtyUsed * Abs(Price(j))
                    matchedQty = matchedQty + qtyUsed

                    If demand <= 0 Then Exit For
                End If
NextPurchase:
            Next j

            If matchedQty > 0 Then
                PnL(i) = matchedQty * salePrice - totalCost
            Else
                PnL(i) = 0
            End If
        End If
    Next i

    For i = 1 To n
        tbl.DataBodyRange.Cells(i, colRemain).Value2 = IIf(Price(i) < 0, Number(i) - Remaining(i), "")
        tbl.DataBodyRange.Cells(i, colPnL).Value2 = IIf(Price(i) > 0, PnL(i), 0)
    Next i
    ' === End core logic ===
    Application.Calculate
CleanExit:
    ' Always restore application state
    Application.Calculation = originalCalc
    'Application.DisplayAlerts = True
    'Application.EnableEvents = True
    'Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "FIFO strategy failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
