Attribute VB_Name = "Module1"
Sub split_rows()
Attribute split_rows.VB_ProcData.VB_Invoke_Func = " \n14"
'
' split_rows Macro
'
Function SplitRows(ByRef dataRng As Range, ByVal splitCol As Long, ByVal splitSep As String, _
                   Optional ByVal idCol As Long = 1) As Boolean
  SplitRows = True

  Dim oldUpd As Variant: oldUpd = Application.ScreenUpdating
  Dim oldCal As Variant: oldCal = Application.Calculation

  On Error GoTo err_sub

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual

  'Get the current number of data rows
  Dim rowCount As Long: rowCount = dataRng.Rows.Count

  'If an ID column is specified, use it to determine where the table ends by finding the first row
  '  with no data in that column
  If idCol > 0 Then
    With dataRng
      rowCount = .Offset(, idCol - 1).Resize(, 1).End(xlDown).Row - .Row + 1
    End With
  End If

  Dim splitArr() As String
  Dim splitLb As Long, splitUb As Long, splitI As Long
  Dim editedRowRng As Range

  'Loop through the data rows to split them as needed
  Dim r As Long: r = 0
  Do While r < rowCount
    r = r + 1

    'Split the string in the specified column
    splitArr = Split(dataRng.Cells(r, "O").Value & "", splitSep)
    splitLb = LBound(splitArr)
    splitUb = UBound(splitArr)

    'If the string was not split into more than 1 item, skip this row
    If splitUb <= splitLb Then GoTo splitRows_Continue

    'Replace the unsplit string with the first item from the split
    Set editedRowRng = dataRng.Resize(1).Offset(r - 1)
    editedRowRng.Cells(1, "O").Value = splitArr(splitLb)

    'Create the new rows
    For splitI = splitLb + 1 To splitUb
      editedRowRng.Offset(1).Insert 'Add a new blank row
      Set editedRowRng = editedRowRng.Offset(1) 'Move down to the next row
      editedRowRng.Offset(-1).Copy Destination:=editedRowRng 'Copy the preceding row to the new row
      editedRowRng.Cells(1, "O").Value = splitArr(splitI) 'Place the next item from the split string

      'Account for the new row in the counters
      r = r + 1
      rowCount = rowCount + 1
    Next

splitRows_Continue:
  Loop

exit_sub:
  On Error Resume Next

  'Resize the original data range to reflect the new, full data range
  If rowCount <> dataRng.Rows.Count Then Set dataRng = dataRng.Resize(rowCount)

  'Restore the application settings
  If Application.ScreenUpdating <> oldUpd Then Application.ScreenUpdating = oldUpd
  If Application.Calculation <> oldCal Then Application.Calculation = oldCal
  Exit Function

err_sub:
  SplitRows = False
  Resume exit_sub
End Function

Sub Main()
Call SplitRows(Range("A2:T20175"), 15, ",")
End Sub



'
End Sub
