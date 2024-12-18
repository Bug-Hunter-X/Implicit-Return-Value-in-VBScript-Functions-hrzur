Function f(x)
  If x > 10 Then
    f = x * 2
  Else
    f = x + 5
  End If
End Function
MsgBox f(2)

'Improved Version:
Function f2(x)
  Dim result
  If x > 10 Then
    result = x * 2
  Else
    result = x + 5
  End If
  f2 = result 'Explicitly assign the return value
End Function
MsgBox f2(2) 'This will now return 7, not an error