Attribute VB_Name = "Globals_mod"
Option Explicit





Public Function decodeFromValueList(strValue As String, strValueList As String) As Long

  Dim lngStartOffset As Long, lngEndOffset As Long
  
  lngEndOffset = InStr(strValueList, "," & strValue & ";")
  If lngEndOffset = 0 Then
    decodeFromValueList = -1
  Else
    lngStartOffset = InStrRev(strValueList, ";", lngEndOffset) + 1
    decodeFromValueList = CLng(Mid$(strValueList, lngStartOffset, lngEndOffset - lngStartOffset))
  End If

End Function


Public Function encodeFromValueList(lngValue As Long, strValueList As String) As String

  Dim lngStartOffset As Long, lngEndOffset As Long
  Dim strValue As String
  
  strValue = CStr(lngValue) & ","
  If Left$(strValueList, Len(strValue)) = strValue Then
    lngStartOffset = 1 + Len(strValue)
  Else
    lngStartOffset = InStr(strValueList, ";" & strValue)
    If lngStartOffset = 0 Then Exit Function
    lngStartOffset = lngStartOffset + Len(strValue) + 1
  End If
  
  lngEndOffset = InStr(lngStartOffset, strValueList, ";")
  encodeFromValueList = Mid$(strValueList, lngStartOffset, lngEndOffset - lngStartOffset)

End Function
