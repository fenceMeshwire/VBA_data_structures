Option Explicit

' ______________________________________________________________________
Sub main_program()

Call create_worksheets_data_and_diagram
Call create_random_portfolio_data
Call create_pie_chart

End Sub

' ______________________________________________________________________
Sub create_worksheets_data_and_diagram()

Dim bleData As Boolean, bleVisualization As Boolean
Dim wksSheet As Worksheet

For Each wksSheet In ThisWorkbook.Worksheets
  If wksSheet.Name = "Data" Then bleData = True
  If wksSheet.Name = "Visualization" Then bleVisualization = True
Next wksSheet

If bleData = False Then ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).Name = "Data"
If bleVisualization = False Then ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).Name = "Visualization"

Application.DisplayAlerts = False
For Each wksSheet In ThisWorkbook.Worksheets
  If wksSheet.Name <> "Data" And wksSheet.Name <> "Visualization" And ThisWorkbook.Sheets.Count > 1 Then
    wksSheet.Delete
  End If
Next wksSheet
Application.DisplayAlerts = True

End Sub

' ______________________________________________________________________
Sub create_random_portfolio_data()

Dim Data As Worksheet
Dim intCounter%, intChkSum%, intSumTotal%, intRand%
Dim varRandom(1 To 10) As Variant

Set Data = ThisWorkbook.Worksheets("Data")

For intCounter% = LBound(varRandom) To UBound(varRandom)
  intRand% = Application.WorksheetFunction.RandBetween(5, 20)
  intChkSum = intChkSum + intRand%
  
  If intChkSum <= 100 And intSumTotal% < 100 And intCounter% < 10 Then
    varRandom(intCounter%) = intRand%
    intSumTotal% = intSumTotal% + intRand%
  Else
    If intChkSum > 100 And intSumTotal% <= 100 Then
      varRandom(intCounter%) = Abs(100 - intSumTotal%)
      intSumTotal% = 100
    ElseIf intChkSum < 100 And intSumTotal% < 100 Then
      varRandom(intCounter%) = Abs(intSumTotal% - 100)
      intSumTotal% = 100
    Else
      varRandom(intCounter%) = 0
    End If
  End If
  
Next intCounter%

Data.Cells(1, 1).Value = "Weight"
Data.Cells(1, 2).Value = "Share"

For intCounter% = LBound(varRandom) To UBound(varRandom)
  Data.Cells(intCounter% + 1, 1).Value = varRandom(intCounter%)
  Data.Cells(intCounter% + 1, 2).Value = "Share " & intCounter%
Next intCounter

End Sub
  
' ______________________________________________________________________
Sub create_pie_chart()

Dim intRowMax%
Dim rngData As Range
Dim pieChart As Chart
Dim shpObject As Object
Dim strData As String
Dim Data As Worksheet
Dim Visualization As Worksheet

Set Data = ThisWorkbook.Worksheets("Data")
Set Visualization = ThisWorkbook.Worksheets("Visualization")

For Each shpObject In Visualization.Shapes
  shpObject.Delete
Next shpObject

intRowMax% = Data.Cells(Data.Rows.Count, 1).End(xlUp).Row

Set rngData = Data.Range("A2:A" & intRowMax%)

Set shpObject = Visualization.Shapes.AddChart2(313, xlPie)
Set pieChart = shpObject.Chart

With pieChart
  .SetSourceData Source:=rngData
  .SetElement (msoElementDataLabelOutSideEnd)
  ' Diagram title
  .HasTitle = True
  .ChartTitle.Text = "Portofolio Distribution"
  .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
  ' Legend
  .SetElement (msoElementLegendRight)
  .FullSeriesCollection(1).XValues = "=Data!$B2:$B$" & intRowMax%
  .FullSeriesCollection(1).DataLabels.ShowPercentage = True
  .FullSeriesCollection(1).DataLabels.ShowValue = False
  ' Size
  shpObject.ScaleWidth 1.04, msoFalse, msoScaleFromTopLeft
End With

Set shpObject = Nothing
Set pieChart = Nothing

Visualization.Select

End Sub
