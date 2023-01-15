Option Explicit

' ______________________________________________________________________
Sub create_data_for_chart()

Dim intRandomSuccess%, intRandomRisk%, intRandomCost%
Dim intRow%, intRowMax%, intCol%, intColMax%
Dim strSuccess$
Dim varHeader, varOptions As Variant

varHeader = [{"Option", "Chance", "Risk", "Gain"}]
varOptions = [{"Alpha", "Beta", "Gamma", "Delta", "Epsilon"}]

Data.UsedRange.Clear
Call assign_header(varHeader)
Call assign_options(varOptions)

intRowMax% = Data.UsedRange.Rows.Count
intColMax% = Data.UsedRange.Columns.Count

For intCol% = 2 To intColMax%

  For intRow% = 2 To intRowMax%
    If intCol% = 2 Then
      intRandomSuccess% = Application.WorksheetFunction.RandBetween(15, 70)
      Data.Cells(intRow, intCol).Value = intRandomSuccess%
    ElseIf intCol% = 3 Then
      intRandomRisk% = Application.WorksheetFunction.RandBetween(20, 90)
      Data.Cells(intRow, intCol).Value = intRandomRisk%
    ElseIf intCol% = 4 Then
      intRandomCost% = flatten_to_full_hundreds(Application.WorksheetFunction.RandBetween(1000, 5000))
      Data.Cells(intRow, intCol).Value = intRandomCost%
    End If
  Next intRow
  
Next intCol%

Call assign_results(get_best_options)
Call create_bubble_diagram

End Sub

' ______________________________________________________________________
Function assign_header(ByRef varHeader As Variant)

Dim intCounter%

For intCounter% = LBound(varHeader) To UBound(varHeader)
  Data.Cells(1, intCounter%) = varHeader(intCounter)
Next intCounter%

End Function

' ______________________________________________________________________
Function assign_options(ByRef varOptions As Variant)

Dim intCounter%

For intCounter% = LBound(varOptions) To UBound(varOptions)
  Data.Cells(intCounter% + 1, 1) = varOptions(intCounter)
Next intCounter%

End Function

' ______________________________________________________________________
Function flatten_to_full_hundreds(ByVal intRandomNumber%) As Integer

Dim intResult%
Dim strRandom$, strRemainder$, strResult$
Dim intLengthRandom%

strRandom$ = intRandomNumber%
intLengthRandom% = Len(strRandom$)

strRemainder$ = Left(strRandom$, intLengthRandom - 2)
strResult$ = strRemainder$ + "00"
intResult% = CInt(strResult$)

Debug.Print intResult%

flatten_to_full_hundreds = intResult%

End Function

' ______________________________________________________________________
Function get_best_options() As Variant

Dim intRow%, intRowMax%
Dim intHighestChance%, intLowestRisk%, lngGreatestGain&
Dim varTarget(1 To 3) As Variant

intRowMax% = Data.UsedRange.Rows.Count
intHighestChance% = Application.WorksheetFunction.Max(Data.Range("B2:B" & intRowMax%))
intLowestRisk% = Application.WorksheetFunction.Min(Data.Range("C2:C" & intRowMax%))
lngGreatestGain& = Application.WorksheetFunction.Max(Data.Range("D2:D" & intRowMax))

For intRow% = 2 To intRowMax%
  If Data.Cells(intRow%, 2).Value = intHighestChance% Then Exit For
Next intRow%

varTarget(1) = Data.Cells(intRow%, 1).Value

For intRow% = 2 To intRowMax%
  If Data.Cells(intRow%, 3).Value = intLowestRisk% Then Exit For
Next intRow%

varTarget(2) = Data.Cells(intRow%, 1).Value

For intRow% = 2 To intRowMax%
  If Data.Cells(intRow%, 4).Value = lngGreatestGain& Then Exit For
Next intRow%

varTarget(3) = Data.Cells(intRow%, 1).Value

get_best_options = varTarget

End Function
' ______________________________________________________________________

Function assign_results(ByRef varResults As Variant)

Visualization.Cells(1, 1).Value = "Highest Chance:"
Visualization.Cells(1, 2).Value = varResults(1)
Visualization.Cells(2, 1).Value = "Lowest Risk:"
Visualization.Cells(2, 2).Value = varResults(2)
Visualization.Cells(3, 1).Value = "Greatest Gain:"
Visualization.Cells(3, 2).Value = varResults(3)
Visualization.Columns.AutoFit

End Function

' ______________________________________________________________________
Function create_bubble_diagram()

Dim intRowMax%

Dim shpClear As Shape
Dim shpChart As Shape
Dim shpDiagram As Chart

For Each shpClear In Visualization.Shapes
  shpClear.Delete
Next

Set shpChart = Visualization.Shapes.AddChart2(404, xlBubble)
Set shpDiagram = shpChart.Chart

intRowMax% = Data.UsedRange.Rows.Count

With shpDiagram
  .SetSourceData Source:=Data.Range("B1:D" & intRowMax%)
  .HasTitle = True
  .ChartTitle.Text = "Crystal Ball of Decision" ' Title of the Chart
  With .ChartTitle.Format.TextFrame2.TextRange
    .Font.Size = 14
    .Font.Fill.Solid
    .Font.Bold = msoTrue
  End With
  
  .Axes(xlCategory, xlPrimary).HasTitle = True ' Designation of the axes
  .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Chance - Percent"
  .Axes(xlValue, xlPrimary).HasTitle = True
  .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Risk - Percent"
  
  .HasLegend = False ' Hide legend
  .ApplyDataLabels   ' Show labels
  
  With .FullSeriesCollection(1).DataLabels
    .Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "=Data!$A$2:$A$" & intRowMax, 0
    .ShowRange = True
    .ShowValue = False
    .Position = xlLabelPositionCenter
    .Font.Bold = True
  End With
  
  .ClearToMatchStyle
  .ChartStyle = 313 ' Motor City: 313
End With

Set shpChart = Nothing
Set shpDiagram = Nothing

End Function
