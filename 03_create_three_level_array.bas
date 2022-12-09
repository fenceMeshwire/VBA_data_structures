Option Explicit

Sub create_three_level_array()

Dim b() As Double

ReDim b(1, 2, 3)

' First digit/dimension:  0, 1
' Second digit/dimension: 0, 1, 2
' Third digit/dimension:  0, 1, 2, 3
' Number of digits: 2 * (3 * 4) = 24

b(0, 0, 0) = 1: b(0, 0, 1) = 2: b(0, 0, 2) = 3: b(0, 0, 3) = 4
b(0, 1, 0) = 5: b(0, 1, 1) = 6: b(0, 1, 2) = 7: b(0, 1, 3) = 8
b(0, 2, 0) = 9: b(0, 2, 1) = 10: b(0, 2, 2) = 11: b(0, 2, 3) = 12

b(1, 0, 0) = 13: b(1, 0, 1) = 14: b(1, 0, 2) = 15: b(1, 0, 3) = 16
b(1, 1, 0) = 17: b(1, 1, 1) = 18: b(1, 1, 2) = 19: b(1, 1, 3) = 20
b(1, 2, 0) = 21: b(1, 2, 1) = 22: b(1, 2, 2) = 23: b(1, 2, 3) = 24

End Sub
