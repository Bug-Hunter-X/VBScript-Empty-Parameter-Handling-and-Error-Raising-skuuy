Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbErrorArgumentMissing, , "Both parameters are required"
  Else
    ' ... rest of the function
  End If
End Function