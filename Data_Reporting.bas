Attribute VB_Name = "Data_Reporting"
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: WriteFilePathLine
'Purpose: Add new line of file data to worksheet
'***************************************************************************
Sub WriteFilePathLine(wksDestination As Worksheet, strFile As String, strPath As String, lngSize As Long, Optional strMessage As String)
  With wksDestination
    Dim lngLastRow As Long
    lngLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    .Cells(lngLastRow, 1) = strFile
    .Cells(lngLastRow, 2) = strPath
    .Cells(lngLastRow, 3) = lngSize
    .Cells(lngLastRow, 4) = removeExtension(strFile)
    If Not strMessage = "" Then
      Call statusBarDisplay(lngLastRow, strMessage)
    End If
  End With
  DoEvents
End Sub
