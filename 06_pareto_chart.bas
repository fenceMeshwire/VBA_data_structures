Option Explicit

' Create two charts: Data and Visualization.
' __________________________________________________________________________________________
Sub generate_data_and_pareto()

Call generate_data
Call create_pareto

End Sub

' __________________________________________________________________________________________
Sub generate_data()

Dim lngRow&, lngRowMax&, lngCounter&
Dim varDesignation As Variant

varDesignation = [{"Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta"}]

With Data
  .UsedRange.Rows.Clear
  .Cells(1, 1).Value = "Error Category"
  .Cells(1, 2).Value = "Number of Errors"
  For lngRow& = LBound(varDesignation) To UBound(varDesignation)
    .Cells(lngRow& + 1, 1).Value = varDesignation(lngRow&)
    .Cells(lngRow& + 1, 2).Value = Application.WorksheetFunction.RandBetween(1, 100)
  Next lngRow&
End With

End Sub

' __________________________________________________________________________________________
Sub create_pareto()

Dim chrt As Object
Dim shpe As Object

Dim lngRow&, lngRowMax&, lngCounter&
Dim rngData As Range

' Clear shapes from visualization table
For Each shpe In Visualization.Shapes
  shpe.Delete
Next shpe

Set shpe = Visualization.Shapes.AddChart2(366, xlPareto)
Set chrt = shpe.chart

lngRowMax& = Data.UsedRange.Rows.Count
Set rngData = Data.Range("A1:B" & lngRowMax&)

On Error Resume Next
With chrt
  ' Set data source, which creates an error, which is handled by the error handler above.
  .SetSourceData Source:=rngData
  ' Title of the diagram
  .HasTitle = True
  .ChartTitle.Text = "Pareto chart"
  .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
  ' Designation of the axis
  .Axes(xlCategory).HasTitle = True
  .Axes(xlCategory).AxisTitle.Caption = Data.Cells(1, 1).Value
  .Axes(xlValue).HasTitle = True
  .Axes(xlValue).AxisTitle.Caption = Data.Cells(1, 2).Value
End With

Set shpe = Nothing
Set chrt = Nothing

End Sub
