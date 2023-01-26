Option Explicit

Sub distribute_budget_two_weights()

Dim curTotalCost as Currency
Dim curCostFirstProject As Currency
Dim curCostSecondProject As Currency

Dim dblTotalHours As Double
Dim dblHoursFirstProject As Double
Dim dblHoursSecondProject As Double
Dim dblWeight1 As Double
Dim dblWeight2 As Double

Dim lngRow&, lngRowMax&
Dim intFirstProject%, intSecondProject%, intTotalCost%
Dim strDifferentiator$

curTotalCost = 96000 ' Workers wage $60 per hour 1600 hours per year
dblTotalHours = 8000 ' 5 workers each working 1600 hours per year

dblWeight1 = 0.69
dblWeight2 = 1 - dblWeight1

With Sheet1

  lngRowMax& = .Cells(.Rows.Count, 4).End(xlUp).Row
  For lngRow& = 5 To lngRowMax&
    strDifferentiator$ = .Cells(lngRow&, 5).Value
    If strDifferentiator$ = "Project 1" Then intFirstProject% = intFirstProject% + 1
    If strDifferentiator$ = "Project 2" Then intSecondProject% = intSecondProject% + 1
    intTotalCost% = intTotalCost% + 1
  Next lngRow&
  
  ' Deploying weight factor:
  curCostFirstProject = curTotalCost * dblWeight1
  curCostSecondProject = curTotalCost * dblWeight2
  
  curCostFirstProject = curCostFirstProject / intFirstProject%
  curCostSecondProject = curCostSecondProject / intSecondProject%
  
  dblHoursFirstProject = dblArbeitsstunden * dblWeight1
  dblHoursFirstProject = Round((dblHoursFirstProject / intFirstProject%), 2)
  
  dblHoursSeconProject = dblArbeitsstunden * dblWeight2
  dblHoursSeconProject = Round((dblHoursSeconProject / intSecondProject%), 2)
  
  For lngRow& = 5 To lngRowMax&
    strDifferentiator$ = .Cells(lngRow&, 5).Value
    If strDifferentiator$ = "Project 1" Then
      .Cells(lngRow&, 2).Value = curCostFirstProject
      .Cells(lngRow&, 3).Value = dblHoursFirstProject
    End If
    If strDifferentiator$ = "Project 2" Then
      .Cells(lngRow&, 2).Value = curCostSecondProject
      .Cells(lngRow&, 3).Value = dblHoursSeconProject
    End If
    intTotalCost% = intTotalCost% + 1
  Next lngRow&
  
End With

End Sub
