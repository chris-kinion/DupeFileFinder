Attribute VB_Name = "Data_Format"
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: removeExtension
'Purpose: Removes any extensions from a file name
'***************************************************************************
Function removeExtension(fileName As String) As String
  If fileName = "" Then
    removeExtension = ""
    Exit Function
  ElseIf InStrRev(fileName, ".") = 0 Then
    removeExtension = fileName
    Exit Function
  Else
    removeExtension = Left(fileName, InStrRev(fileName, ".") - 1)
  End If
End Function
