Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

'Example usage:
Set objExcel = GetObject("Excel.Application")
if objExcel is nothing then
  msgbox "Excel is not running"
exit sub
end if