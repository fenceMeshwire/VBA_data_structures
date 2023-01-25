Option Explicit

' Create two charts: Data and Visualization.
' __________________________________________________________________________________________
Sub generate_data_and_radar_map()

Call generate_data
Call create_radar_map_diagram

End Sub

' __________________________________________________________________________________________
Sub generate_data()

Dim lngRow&
Dim varDesignation As Variant
Dim varUrgency As Variant

varDesignation = [{"Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta"}]
varUrgency = [{"Mandatory", "Required", "Optional"}]

With Data
  .UsedRange.Rows.Clear
  .Cells(1, 1).Value = "Category"
  .Cells(1, 2).Value = "Urgency"
  .Cells(1, 3).Value = "Fast"
  .Cells(1, 4).Value = "Good"
  .Cells(1, 5).Value = "Cheap"
  For lngRow& = LBound(varDesignation) To UBound(varDesignation)
    .Cells(lngRow& + 1, 1).Value = varDesignation(lngRow&)
    .Cells(lngRow& + 1, 2).Value = varUrgency(Application.WorksheetFunction.RandBetween(1, 3))
    .Cells(lngRow& + 1, 3).Value = Application.WorksheetFunction.RandBetween(1, 100) / 100
    .Cells(lngRow& + 1, 3).NumberFormat = "0%"
    .Cells(lngRow& + 1, 4).Value = Application.WorksheetFunction.RandBetween(1, 100) / 100
    .Cells(lngRow& + 1, 4).NumberFormat = "0%"
    .Cells(lngRow& + 1, 5).Value = Application.WorksheetFunction.RandBetween(1, 100) / 100
    .Cells(lngRow& + 1, 5).NumberFormat = "0%"
  Next lngRow&
End With

End Sub

' __________________________________________________________________________________________
Sub create_radar_map_diagram()

Dim chrt As Object
Dim shpe As Object
Dim lngRow&, lngRowMax&
Dim rngData As Range

' Clear shapes from visualization table
For Each shpe In Visualization.Shapes
  shpe.Delete
Next shpe

lngRowMax& = Data.UsedRange.Rows.Count
Set rngData = Data.Range("A1:E" & lngRowMax&)

Set shpe = Visualization.Shapes.AddChart2(317, xlRadar)
Set chrt = shpe.Chart

On Error Resume Next
With chrt
  .SetSourceData Source:=rngData
  .HasTitle = True
  .ChartTitle.Text = "Visualization: Radar Map"
  .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
  .HasLegend = True
End With

Set shpe = Nothing
Set chrt = Nothing

End Sub
