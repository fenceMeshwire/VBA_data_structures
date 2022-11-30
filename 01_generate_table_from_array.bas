Option Explicit

' ______________________________________________________________________
Sub generate_table_from_array()

Dim intValue, intRow As Integer
Dim varValues As Variant
Dim wksSheet As Worksheet

Set wksSheet = Tabelle1

varValues = generate_values()

intRow = 1

With wksSheet

  For intValue = LBound(varValues) To UBound(varValues)
  
    wksSheet.Cells(intRow, 1).Value = varValues(intValue)(1)
    wksSheet.Cells(intRow, 2).Value = varValues(intValue)(2)
    wksSheet.Cells(intRow, 3).Value = CCur(varValues(intValue)(3))
    wksSheet.Cells(intRow, 4).Value = varValues(intValue)(4)
    intRow = intRow + 1
    
  Next intValue
  
  .Columns.AutoFit
  
End With

End Sub

' ______________________________________________________________________
Function generate_values() As Variant

Dim varValues(1 To 6) As Variant

' varValues: (6 x 1)-Matrix
Dim varValue01, varValue02, varValue03, varValue04, varValue05, varValue06 As Variant

varValue01 = [{"Property 1", 50, 215.45, "Description 1"}]
varValue02 = [{"Property 2", 75, 750.60, "Description 2"}]
varValue03 = [{"Property 3", 100, 178.60, "Description 3"}]
varValue04 = [{"Property 4", 125, 205.10, "Description 4"}]
varValue05 = [{"Property 5", 150, 672.30, "Description 5"}]
varValue06 = [{"Property 6", 175, 142.80, "Description 6"}]

varValues(1) = varValue01
varValues(2) = varValue02
varValues(3) = varValue03
varValues(4) = varValue04
varValues(5) = varValue05
varValues(6) = varValue06

generate_values = varValues

End Function
