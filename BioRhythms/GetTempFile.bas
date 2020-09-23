Attribute VB_Name = "modGetTempFile"
Option Explicit

Public Declare Function GetTempPath Lib "kernel32" _
      Alias "GetTempPathA" _
      (ByVal nBufferLength As Long, _
       ByVal lpBuffer As String) As Long
       
Public Declare Function GetTempFilename Lib "kernel32" _
      Alias "GetTempFileNameA" _
      (ByVal lpszPath As String, _
       ByVal lpPrefixString As String, _
       ByVal wUnique As Long, _
       ByVal lpTempFileName As String) As Long
       
Private Const MAX_PATH& = 260
'

' returns temp path &  temp file name
' Can be optionaly be passed a filename Prefix
' ----------------------------------------------------------------------
Public Function fnGetTempPath(Optional strPrefix As String = "") As String
' ----------------------------------------------------------------------
  

Dim lngReturnVal As Long
Dim strTempPath As String * MAX_PATH
Dim strTempFileName As String * MAX_PATH
On Error GoTo Handle_Error
If Len(strPrefix) > 0 Then
   strPrefix = Trim(strPrefix)
Else
   strPrefix = ""
End If

lngReturnVal = GetTempPath(MAX_PATH, strTempPath)
lngReturnVal = GetTempFilename(strTempPath, strPrefix, 0, strTempFileName)

fnGetTempPath = strTempFileName

Exit_Point:
  Exit Function

Handle_Error:

  MsgBox "Unexpected Error Returned" & vbCrLf _
       & "Return value " & lngReturnVal & vbCrLf _
       & "Cannot retrieve temporary filename" & vbCrLf _
       & "Nmbr " & Err.Number & vbCrLf _
       & "Desc " & Err.Description & vbCrLf _
       & "Srce " & Err.Source & vbCrLf _
       & "Time " & Now & vbCrLf _
       & "Path " & App.Path, _
     16, "You have a Problem!"
     
  fnGetTempPath = ""
  Resume Exit_Point
End Function
' returns a temp file name for the current directory.
' ----------------------------------------------------------------------
Public Function fnGetTempFileName(Optional strPrefix As String = "") As String
' ----------------------------------------------------------------------
  On Error GoTo Handle_Error

Dim lngReturnVal As Long
Dim strTempPath As String * MAX_PATH
Dim strTempFileName As String * MAX_PATH

If Len(strPrefix) > 0 Then
   strPrefix = Trim(strPrefix)
Else
   strPrefix = ""
End If

strTempPath = CurDir & "\" & vbNullChar
lngReturnVal = GetTempFilename(strTempPath, strPrefix, 0, strTempFileName)

fnGetTempFileName = strTempFileName

 
Exit_Point:
  Exit Function

Handle_Error:

  MsgBox "Unexpected Error Returned" & vbCrLf _
       & "Return value " & lngReturnVal & vbCrLf _
       & "Cannot retrieve temporary filename" & vbCrLf _
       & "Nmbr " & Err.Number & vbCrLf _
       & "Desc " & Err.Description & vbCrLf _
       & "Srce " & Err.Source & vbCrLf _
       & "Time " & Now & vbCrLf _
       & "Path " & App.Path, _
     16, "You have a Problem!"
  fnGetTempFileName = ""
  Resume Exit_Point
End Function

