Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function

Sub CallMyFunction()
  On Error GoTo ErrHandler
  Dim result
  result = MyFunction("")
  Exit Sub
ErrHandler:
  MsgBox "Error: " & Err.Number & " - " & Err.Description
End Sub