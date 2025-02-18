Function f(a)
  Dim result
  result = 1
  If a > 1 Then
    For i = 2 To a
      result = result * i
    Next
  End If
  f = result
End Function
MsgBox f(5)