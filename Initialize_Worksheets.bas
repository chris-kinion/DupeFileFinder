Attribute VB_Name = "Initialize_Worksheets"
Option Explicit
Option Base 1

Public gwksControls As Worksheet
Public gwksSource As Worksheet
Public gwksDupes  As Worksheet
Public gwksPartials As Worksheet
Public lngCounter As Long
Public strSearchMessage As String

'***************************************************************************
'Procedure: InitSheets
'Purpose: Initialize global variables and prepare worksheets for use
'***************************************************************************
Sub InitSheets()
  With ThisWorkbook
    Set gwksControls = .Worksheets("Controls")
    Set gwksSource = .Worksheets("Source Files")
    Set gwksDupes = .Worksheets("Duplicate Files")
    Set gwksPartials = .Worksheets("Partial Matches")
  End With
  
  lngCounter = 0
  strSearchMessage = ""
  
  Dim wksThisSheet As Worksheet
  Dim arrSheetNames
  Dim i As Long
  
  arrSheetNames = Array(gwksSource.Name, gwksDupes.Name, gwksPartials.Name)
    
  For Each wksThisSheet In ThisWorkbook.Worksheets
    For i = LBound(arrSheetNames) To UBound(arrSheetNames)
      If wksThisSheet.Name = arrSheetNames(i) Then
        'Debug.Print wksThisSheet.Name
        wksThisSheet.Cells.Clear
        With wksThisSheet
          .Cells(1, 1).Value = "Filename"
          .Columns("A:A").ColumnWidth = 40
          .Cells(1, 2).Value = "Path"
          .Columns("B:B").ColumnWidth = 50
          .Cells(1, 3).Value = "Size"
          .Columns("C:C").ColumnWidth = 20
          .Cells(1, 4).Value = "Name Without Extension"
          .Columns("D:D").ColumnWidth = 40
        End With
      End If
    Next i
  Next wksThisSheet
  
End Sub

'***************************************************************************
'Procedure: InitSheets2
'Purpose: Set worksheet global variables only
'***************************************************************************
Sub InitSheets2()
  With ThisWorkbook
    Set gwksControls = .Worksheets("Controls")
    Set gwksSource = .Worksheets("Source Files")
    Set gwksDupes = .Worksheets("Duplicate Files")
    Set gwksPartials = .Worksheets("Partial Matches")
  End With
End Sub

'***************************************************************************
'Procedure: FinalizeWorksheets
'Purpose: Scroll to top and select A1 of each worksheet
'***************************************************************************
Sub FinalizeWorksheets()
  Dim wksThisSheet As Worksheet
  Dim arrSheetNames
  Dim i As Long
  
  arrSheetNames = Array(gwksSource.Name, gwksDupes.Name, gwksPartials.Name)
    
  For Each wksThisSheet In ThisWorkbook.Worksheets
    For i = LBound(arrSheetNames) To UBound(arrSheetNames)
      If wksThisSheet.Name = arrSheetNames(i) Then
        Call displayPageFromTop(wksThisSheet)
      End If
    Next i
  Next wksThisSheet
End Sub
