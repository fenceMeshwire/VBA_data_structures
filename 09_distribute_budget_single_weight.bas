Option Explicit

Sub distribute_budget_single_weight()

Dim curTotal, curProject As Currency

Dim dblTotalHours As Double
Dim dblHours As Double

Dim lngRow&, lngRowMax&
Dim intTotal%

curTotal = 96000 ' Workers wage $60 per hour 1600 hours per year
dblTotalHours = 8000 ' 5 workers each working 1600 hours per year

With Sheet1

  lngRowMax& = .Cells(.Rows.Count, 1).End(xlUp).Row
  For lngRow& = 5 To lngRowMax&
    intTotal% = intTotal% + 1
  Next lngRow&
  
  curProject = Round(curTotal / intTotal%, 2)
  dblHours = Round(dblTotalHours / intTotal%, 2)
  
  For lngRow& = 5 To lngRowMax&
    .Cells(lngRow&, 2).Value = curProject
    .Cells(lngRow&, 3).Value = dblHours
  Next lngRow&
  
End With

End Sub
