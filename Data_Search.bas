Attribute VB_Name = "Data_Search"
Option Explicit
Option Base 1

' WARNING!
' This tool uses Early Binding so you must include a reference to Microsoft Scripting Runtime:
' Tools > References > Microsoft Scripting Runtime

Dim FSO As Scripting.FileSystemObject
Dim blnNotFirstIteration As Boolean
Dim thisFile As file
Dim thisFolder As Folder
Dim subFolder As Folder

'***************************************************************************
'Procedure: ReportAndFilterFiles
'Purpose: Select up to two folders for file comparison
'***************************************************************************
Sub ReportAndFilterFiles()
  Dim startPath As String
  Dim nextPath As String
  startPath = ""
  nextPath = ""
  
  Call getSearchPath(startPath, "Select Folder to Search")
  If startPath = "" Then
    MsgBox "Search cancelled by user."
    Exit Sub
  End If
  Dim intOutput As Integer
  intOutput = MsgBox("Add Another Folder to Search?", vbYesNo, "Expand Search")
  
  Call InitSheets
  Call ReduceFunctionality
  If intOutput = 6 Then ' vbYes
    Call getSearchPath(nextPath, "Select Folder to Search")
    Call UpdateStatusBar("Searching " & startPath)
    Call SearchForAllFiles(startPath)
    If Not nextPath = "" Then
      strSearchMessage = "Searching " & nextPath
      Call UpdateStatusBar(strSearchMessage)
      Call SearchForAllFiles(nextPath)
    End If
  ElseIf intOutput = 7 Then ' vbNo
    strSearchMessage = "Searching " & startPath
    Call UpdateStatusBar(strSearchMessage)
    Call SearchForAllFiles(startPath)
  End If
  Call FilterRowData
  
  Call FinalizeWorksheets ' Formating and Wrap-up
  Call displayPageFromTop(gwksSource)
  Call RestoreFunctionality
  Call UpdateStatusBar("")
  
End Sub

'***************************************************************************
'Procedure: SearchForAllFiles
'Purpose: Recursively find all files in folders and subfolders
'Comments: Adapted from https://wellsr.com/vba/2018/excel/list-files-in-folder-and-subfolders-with-vba-filesystemobject/
'***************************************************************************
Sub SearchForAllFiles(ByVal HostFolder As String)
  If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
  Set thisFolder = FSO.GetFolder(HostFolder)
  
  If Not blnNotFirstIteration Then
    For Each thisFile In thisFolder.Files
      lngCounter = lngCounter + 1
      Call WriteFilePathLine(gwksSource, thisFile.Name, thisFile.Path, thisFile.Size, strSearchMessage)
    Next thisFile
    
    If Not thisFolder.SubFolders Is Nothing Then
      blnNotFirstIteration = True
      Call SearchForAllFiles(HostFolder)
    End If
  Else ' Not first iteration
    For Each subFolder In thisFolder.SubFolders
      For Each thisFile In subFolder.Files
        lngCounter = lngCounter + 1
        Call WriteFilePathLine(gwksSource, thisFile.Name, thisFile.Path, thisFile.Size, strSearchMessage)
      Next thisFile
      
      If Not subFolder.SubFolders Is Nothing Then
        Call SearchForAllFiles(HostFolder & "\" & subFolder.Name)
      End If
    Next subFolder
    blnNotFirstIteration = False
  End If
End Sub

'***************************************************************************
'Procedure: FilterRowData
'Purpose: Remove duplicate entries, sort, report duplicate and partial matches
'***************************************************************************
Sub FilterRowData(Optional thisWks As Worksheet)
  Dim strStatus As String
  strStatus = "Filtering Results"
  
  Call UpdateStatusBar("Sorting Results")
  If thisWks Is Nothing Then
    Set thisWks = gwksSource
  End If
  
  thisWks.Range("A1").CurrentRegion.RemoveDuplicates Columns:=2, Header:=xlYes
  With thisWks.Sort.SortFields
    .Add Key:=Range("$D:$D"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Add Key:=Range("$A:$A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Add Key:=Range("$C:$C"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  End With
  With thisWks.Sort
    .SetRange Range("$A:$D")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlSortColumns
    .SortMethod = xlPinYin
    .Apply
  End With
  
  Dim lngLastRow As Long
  Dim lngRowNum As Long
  With thisWks
    lngLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
  End With
  If lngLastRow > 3 Then
    With thisWks
      Call UpdateStatusBar(strStatus)
      If .Cells(2, 1) = .Cells(3, 1) And .Cells(2, 3) = .Cells(3, 3) Then
        gwksDupes.Cells(2, 1) = .Cells(2, 1) ' copy this row to dupes
        gwksDupes.Cells(2, 2) = .Cells(2, 2)
        gwksDupes.Cells(2, 3) = .Cells(2, 3)
      ElseIf .Cells(2, 4) = .Cells(3, 4) Then
        gwksPartials.Cells(2, 1) = .Cells(2, 1) ' copy this row to partials
        gwksPartials.Cells(2, 2) = .Cells(2, 2)
        gwksPartials.Cells(2, 3) = .Cells(2, 3)
        gwksPartials.Cells(2, 4) = .Cells(2, 4)
      End If
      For lngRowNum = 3 To lngLastRow
        Call statusBarDisplay(lngRowNum, strStatus)
        If (.Cells(lngRowNum, 1) = .Cells(lngRowNum - 1, 1) And .Cells(lngRowNum, 3) = .Cells(lngRowNum - 1, 3)) Or _
            (.Cells(lngRowNum, 1) = .Cells(lngRowNum + 1, 1) And .Cells(lngRowNum, 3) = .Cells(lngRowNum + 1, 3)) Then
          Dim lngDupeRow As Long ' copy exact match row to dupes
          lngDupeRow = gwksDupes.Cells(gwksDupes.Rows.Count, 1).End(xlUp).Row + 1
          gwksDupes.Cells(lngDupeRow, 1) = .Cells(lngRowNum, 1)
          gwksDupes.Cells(lngDupeRow, 2) = .Cells(lngRowNum, 2)
          gwksDupes.Cells(lngDupeRow, 3) = .Cells(lngRowNum, 3)
        ElseIf .Cells(lngRowNum, 4) = .Cells(lngRowNum - 1, 4) Or .Cells(lngRowNum, 4) = .Cells(lngRowNum + 1, 4) Then
          Dim lngPartialRow As Long ' copy partial match row to dupes
          lngPartialRow = gwksPartials.Cells(gwksPartials.Rows.Count, 1).End(xlUp).Row + 1
          gwksPartials.Cells(lngPartialRow, 1) = .Cells(lngRowNum, 1)
          gwksPartials.Cells(lngPartialRow, 2) = .Cells(lngRowNum, 2)
          gwksPartials.Cells(lngPartialRow, 3) = .Cells(lngRowNum, 3)
          gwksPartials.Cells(lngPartialRow, 4) = .Cells(lngRowNum, 4)
        End If
      Next lngRowNum
    End With
  Else
    Debug.Print "Not enough data"
  End If
  Call UpdateStatusBar("")
End Sub




