Function f(a, b)
  If IsEmpty(a) Then
    WScript.Echo "a is empty"
    Exit Function
  End If
  If IsEmpty(b) Then
    WScript.Echo "b is empty"
    Exit Function
  End If
  WScript.Echo a + b
End Function

f(1, 2)