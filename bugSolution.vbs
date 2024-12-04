Function MyFunction(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise 1001, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function

Sub CallMyFunction()
  On Error Resume Next
  MyFunction ""
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub