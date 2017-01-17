Attribute VB_Name = "Test"
Public activeRow As Long
Public activeCol As Long

Sub copiacelle()
Dim message As String
If activeRow = 0 Then
    message = "No previous cell selected"
Else
    message = "Previous cell was " & CStr(activeRow) & " " & CStr(activeRow)
End If

MsgBox message
activeRow = ActiveCell.Row
activeCol = ActiveCell.Column
MsgBox CStr(activeRow) + " " + CStr(activeCol)

Dim currValue As String
currValue = ActiveSheet.Cells(activeRow, activeCol).Value
MsgBox currValue

Dim se As Boolean
se = SheetExists(currValue)
MsgBox se

' ----------------------
' Set data validation
' ----------------------
Dim ws As Worksheet
Dim range1 As Range, rng As Range
Set ws = ActiveWorkbook.Worksheets("Index")
Set range1 = ws.Range("A1:A5")
Set rng = ws.Range("B1")

With rng.Validation
    .Delete 'delete previous validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="='" & ws.Name & "'!" & range1.Address
End With

    'max_col = ActiveSheet.UsedRange.Columns.Count
    'For c = 1 To ActiveSheet.UsedRange.Columns.Count
    '   MsgBox ActiveSheet.Cells(r, c).Value
    'Next
    '
'
End Sub


