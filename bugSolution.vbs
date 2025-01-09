Function f(a, b)
  If IsEmpty(a) Or IsNumeric(a) = False Then
    WScript.Echo "a is empty or not numeric"
    Exit Function
  End If
  If IsEmpty(b) Or IsNumeric(b) = False Then
    WScript.Echo "b is empty or not numeric"
    Exit Function
  End If
  WScript.Echo a + b
End Function

f(1, 2)
f(Empty, 2)
f(1, Empty)
f(Empty, Empty)

'Test with numeric variants containing zero
f(0,2)
f(2,0) 

'Test with other non-numeric variants to cover all cases
f("hello", 2) 
f(1, "world")
f("hello","world")
f(Null,2)
f(2,Null)