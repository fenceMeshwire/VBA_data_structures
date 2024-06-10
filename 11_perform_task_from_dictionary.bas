Option Explicit

' 1. Create a VBA module "mdl_main". Master Data contains similar properties, 
'    e.g. "AB21", "XY31" in the first column.
' _______________________________________________________________________________________
Sub perform_task_from_dictionary()

Dim operations As New cls_operations

Dim dictPROP As Object
Dim varPROP As Variant
Dim wkSheet As Worksheet

Set wkSheet = Master_Data

varTypeCode = CallByName(operations, "get_PROP", VbMethod, wkSheet)
Set dictPROP = CallByName(operations, "get_dict_PROP", VbMethod, wkSheet, varPROP)

CallByName operations, "perform_task", VbMethod, wkSheet, dictPROP

End Sub


' 2. Create a VBA class module "cls_operations"
' _______________________________________________________________________________________
Public Function get_PROP(ByRef wkSheet As Worksheet) As Variant 

' This will create a list of duplicates, if there are any duplicates.

Dim intColPROP As Integer

Dim lngRow As Long, lngRowMax As Long
Dim lngCounterVar As Long
Dim lngCounter As Long
Dim strPROP As String
Dim varPROP As Variant

intColPROP = 1

ReDim varPROP(lngCounterVar)

With wkSheet
  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  
  For lngRow = 2 To lngRowMax
    strPROP = .Cells(lngRow, intColPROP).Value
    If Not IsNumeric(Application.Match(strPROP, varPROP, 0)) Then ' No duplicates
      varPROP(lngCounterVar) = strPROP
      lngCounterVar = lngCounterVar + 1
      ReDim Preserve varPROP(lngCounterVar)
    End If
  Next lngRow
  
  ReDim Preserve varPROP(UBound(varPROP) - 1)
  
End With

get_PROP = varPROP

End Function

' _______________________________________________________________________________________
Public Function get_dict_PROP(ByRef wkSheet As Worksheet, ByRef varPROP As Variant) As Object

' This will create a dictionary containing each property, counting the number of duplicates, if there are any duplicates.

Dim intField As Integer, intCounter As Integer
Dim intCounterCheck As Integer, intCounterVar As Integer, intCounterVisible As Integer
Dim strPROP As String

Dim ops As New cls_operations

Dim dict_PROP As Object
Dim varVisibleRows As Variant, varPROPcheck As Variant

Set dict_PROP = CreateObject("Scripting.Dictionary")

intField = 1

On Error GoTo restore_screen_update
Application.ScreenUpdating = False
For intCounterVar = LBound(varPROP) To UBound(varPROP)
  strPROP = varPROP(intCounterVar)
  
  CallByName ops, "apply_filter", VbMethod, strPROP, intField, wkSheet
  varPROPcheck = CallByName(ops, "get_visible_rows", VbMethod, wkSheet, intField)
  
  For intCounterCheck = LBound(varPROPcheck) To UBound(varPROPcheck)
    intCounter = intCounter + 1
  Next intCounterCheck
  dict_PROP.Add Key:=strPROP, Item:=intCounter
  intCounter = 0
  
  CallByName ops, "reset_autofilter", VbMethod, wkSheet
  
Next intCounterVar
Application.ScreenUpdating = True

Set get_dict_PROP = dict_PROP

Exit Function

On Error GoTo restore_screen_update:
Application.ScreenUpdating = True
Call reset_autofilter(wkSheet)

End Function

' ________________________________________________________________
Public Function apply_filter(ByVal strPROP As String, _
    ByVal intField As Integer, ByRef wkSheet As Worksheet)
    
If wkSheet.AutoFilterMode = False Then
  wkSheet.Range("A1").AutoFilter
End If

With wkSheet.Rows("1:1")
    .AutoFilter Field:=intField, Criteria1:=strPROP
End With

With wkSheet.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Function

' _______________________________________________________________________________________
Public Function get_visible_rows(ByRef wksInput As Worksheet, _
    ByVal intField As Integer) As Variant

Dim lngRow As Long, lngRowMax As Long
Dim lngIndex As Long
Dim varRange As Variant

lngRowMax = wksInput.Cells(wksInput.Rows.Count, 1).End(xlUp).Row

ReDim varRange(lngIndex)
For lngRow = 2 To lngRowMax
  If Not wksInput.Rows(lngRow).EntireRow.Hidden Then
    varRange(lngIndex) = lngRow
    lngIndex = lngIndex + 1
    ReDim Preserve varRange(lngIndex)
  End If
Next

ReDim Preserve varRange(lngIndex - 1)

get_visible_rows = varRange

End Function

' _______________________________________________________________________________________
Public Function reset_autofilter(ByRef wkSheet As Worksheet)

wkSheet.ShowAllData

End Function

' _______________________________________________________________________________________
Public Function perform_task(ByRef wkSheet As Worksheet, ByRef dictPROP)

' If there are less than three duplicates, color the 

Dim intCounter As Integer, intCounterMax As Integer
Dim lngRow As Long, lngRowMax As Long
Dim strPROP As String

With wkSheet
  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  intCounterMax = .UsedRange.Columns.Count
  For lngRow = 2 To lngRowMax
    strPROP = .Cells(lngRow, 1).Value
    If dictPROP.Exists(strPROP) Then
      If dictPROP(strPROP) < 3 Then
        For intCounter = 1 To intCounterMax
          .Cells(lngRow, intCounter).Interior.ColorIndex = 34
        Next intCounter
      End If
    End If
  Next lngRow
End With

End Function
