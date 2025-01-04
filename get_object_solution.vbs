Function GetObjectRobust(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    ' Attempt to create the object
    On Error Resume Next
    Set obj = CreateObject(progID)
    If Err.Number <> 0 Then
      ' Handle the creation error (e.g., log, display message)
      MsgBox "Error creating object: " & Err.Description, vbCritical
      Set obj = Nothing
    End If
    On Error GoTo 0
  End If
  On Error GoTo 0
  Set GetObjectRobust = obj
End Function

'Example usage:
Set objExcel = GetObjectRobust("Excel.Application")
if objExcel is nothing then
  msgbox "Excel is not running and could not be created"
exit sub
end if