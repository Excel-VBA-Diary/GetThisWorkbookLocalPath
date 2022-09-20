Attribute VB_Name = "ThisWorkbookLocalPath"
Option Explicit

Private Sub Test_ThisWorkbookLocalPath()
    Debug.Print ThisWorkbookLocalPath
End Sub

'-------------------------------------------------------------------------------
' Return the value of "Hide extensions for known file types" in Folder Options
' return value: 0=show,  1=do not show
'-------------------------------------------------------------------------------
Private Function HideFileExt() As Long
    With CreateObject("WScript.Shell")
        HideFileExt = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt")
    End With
End Function

'-------------------------------------------------------------------------------
' Return the value of "Show ...Recently Opened Items" setting under "Start" in Personal Settings.
' return value: 0=do not show,  1=show
'-------------------------------------------------------------------------------
Private Function Start_TrackDocs() As Long
    With CreateObject("WScript.Shell")
        Start_TrackDocs = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs")
    End With
End Function

'-------------------------------------------------------------------------------
' Return the path on the local drive of SharePoint files synced to OneDrive
' Addressing the issue of ThisWorkbook.Path returning URLs
'-------------------------------------------------------------------------------
Public Function ThisWorkbookLocalPath() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        ThisWorkbookLocalPath = ThisWorkbook.Path
        Exit Function
    End If
    
    If Start_TrackDocs = 0 Then
        MsgBox "'Store and display of recently opened items' id not enabled."
        ThisWorkbookLocalPath = vbNullString
        Exit Function
    End If
    
    Dim recentFileName As String
    If HideFileExt = 0 Then
        recentFileName = ThisWorkbook.Name & ".LNK"
    Else
        recentFileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".LNK"
    End If
    
    Dim wsh As Object, wsc As Object
    Set wsh = CreateObject("WScript.Shell")
    Set wsc = wsh.CreateShortcut(Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Windows\Recent\" & recentFileName)
    ThisWorkbookLocalPath = Replace(wsc.TargetPath, "\" & ThisWorkbook.Name, "")

End Function

