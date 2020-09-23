Attribute VB_Name = "modDateDiff"
Option Explicit
' Author: Louis Boldt 2/16/2002
' --------------------------------------------------------------------
Public Function fnAge(datFrom As Date, _
             Optional datTo As Variant) As String
' --------------------------------------------------------------------

Dim iYYYY As Integer 'final results go here
Dim iMM As Integer
Dim iDD As Integer
 
Dim iYYf As Integer ' from date work vars
Dim iMMf As Integer
Dim iDDf As Integer
Dim sMmDdf As Single
 
Dim iYYt As Integer ' to date work vars
Dim iMMt As Integer
Dim iDDt As Integer
Dim sMmDdt As Single
Dim datWork As Date
  On Error GoTo Handle_Error

  If IsMissing(datTo) Then ' variant used in parm list so IsMissnig will work
    datTo = Date           ' so set to System date
  End If
  
  If datFrom > datTo Then ' insure to date is later date
    datWork = datFrom
    datFrom = datTo
    datTo = datWork
  End If
  
  iYYf = Year(datFrom)
  iMMf = Month(datFrom)
  iDDf = Day(datFrom)
  iYYt = Year(datTo)
  iMMt = Month(datTo)
  iDDt = Day(datTo)
  
  sMmDdf = iMMf + (iDDf / 100) 'Set up for ez compare of month and day
  sMmDdt = iMMt + (iDDt / 100)
  
  If iYYf = iYYt Then ' nail down the year part
    iYYYY = 0
  Else
    iYYYY = iYYt - iYYf
    If sMmDdt < sMmDdf Then 'if todate lt fromdate
      iYYYY = iYYYY - 1
    End If
  End If ' year part is now what i want
  
  ' set datwork to a date that is less that one year from datTo
  datWork = DateAdd("yyyy", iYYYY, datFrom)
  
  iMM = DateDiff("m", datWork, datTo) 'counts how many time the month changes
  
  If iDDf > iDDt Then ' if from day gt to day months is to large
    iMM = iMM - 1
  End If
   
  ' set datwork to within i month of todate
  datWork = DateAdd("m", iMM, datWork)
  ' now get the days
  iDD = DateDiff("d", datWork, datTo)
  
  fnAge = iYYYY & " Year(s) " _
          & iMM & " Month(s) " _
          & iDD & " Day(s)"
  
Exit_Point:
  Exit Function

Handle_Error:

  MsgBox "Unexpected Error Returned" & vbCrLf _
       & "Nmbr " & Err.Number & vbCrLf _
       & "Desc " & Err.Description & vbCrLf _
       & "Srce " & Err.Source & vbCrLf _
       & "Time " & Now & vbCrLf _
       & "Path " & App.Path, _
     16, "You have a Problem!"
  fnAge = "Old Old Old"
  Resume Exit_Point
End Function



