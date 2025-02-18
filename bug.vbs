Function f(a)
  If a = 1 Then
    f = 1
  Else
    f = a * f(a - 1)
  End If
End Function
MsgBox f(5)