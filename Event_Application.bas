Attribute VB_Name = "Event_Application"
Option Explicit
Option Base 1

'***************************************************************************
'Procedure: ReduceFunctionality
'Purpose: Limit Excel activity to boost VBA speed
'***************************************************************************
Sub ReduceFunctionality()
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  'Application.DisplayStatusBar = False
End Sub

'***************************************************************************
'Procedure: RestoreFunctionality
'Purpose: Return Excel to normal upon exit / end procedure
'***************************************************************************
Sub RestoreFunctionality()
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.DisplayStatusBar = True
End Sub

'***************************************************************************
'Procedure: getSearchPath
'Purpose: Set startPath to a selected folder, add "\" at end
'***************************************************************************
Sub getSearchPath(ByRef startPath As String, ByVal searchTitle As String)
  Dim intChoice As Integer
  Dim thisFileName As String
  With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = searchTitle
    .Show
    If .SelectedItems.Count = 0 Then
      startPath = ""
    Else
      startPath = .SelectedItems(1)
      If Right(startPath, 1) <> "\" Then startPath = startPath & "\"
    End If
  End With
End Sub

'***************************************************************************
'Procedure: UpdateStatusBar
'Purpose: Updates the status bar with a message
'Comments: Helps prevent system freeze with DoEvents
'***************************************************************************
Sub UpdateStatusBar(statusMessage As String)
  Application.StatusBar = statusMessage
  DoEvents
End Sub

'***************************************************************************
'Procedure: statusBarDisplay
'Purpose: Animates status bar as counter increases
'Comments: Outputs input string to be changed next iteration
'***************************************************************************
Function statusBarDisplay(lngCounter As Long, strMessage As String) As String
  If lngCounter Mod 1000 = 0 Then
    strMessage = Split(strMessage, ".")(0)
    Application.StatusBar = strMessage
  ElseIf lngCounter Mod 250 = 0 Then
    strMessage = strMessage & "."
    Application.StatusBar = strMessage
    DoEvents
  End If
  statusBarDisplay = strMessage
End Function

'***************************************************************************
'Procedure: displayPageFromTop
'Purpose: Brings to view of top left of page
'***************************************************************************
Sub displayPageFromTop(wksDestinationSheet As Worksheet)
  wksDestinationSheet.Activate
  wksDestinationSheet.Range("A1").Select
  ActiveWindow.ScrollRow = 1
End Sub
