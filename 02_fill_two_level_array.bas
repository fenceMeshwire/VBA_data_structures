Option Explicit

Sub fill_two_level_array()

Dim intCounter As Integer
Dim lngRow, lngRowMax As Long
Dim varCollector, varStorage As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

lngRowMax = wksSheet.UsedRange.Rows.Count
ReDim varStorage(intCounter)
ReDim varCollector(intCounter)

For lngRow = 2 To lngRowMax
    varCollector = Split(wksSheet.Cells(lngRow, 3).Value, "&") ' Array is filled by Split() method.
  varStorage(intCounter) = varCollector
  intCounter = intCounter + 1
    
  ReDim Preserve varCollector(intCounter)
  ReDim Preserve varStorage(intCounter)
Next lngRow

ReDim varCollector(0)
ReDim Preserve varStorage(UBound(varStorage) - 1)

End Sub
