Attribute VB_Name = "ThisWorkbookLocalPath"
Option Explicit

Private Sub Test_ThisWorkbookLocalPath()
    Debug.Print ThisWorkbookLocalPath
End Sub

'-------------------------------------------------------------------------------
' フォルダーオプションの「登録されている拡張子は表示しない」の設定値を返す
' 戻り値：  0=表示する、1=表示しない
'-------------------------------------------------------------------------------
Private Function HideFileExt() As Long
    With CreateObject("WScript.Shell")
        HideFileExt = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt")
    End With
End Function

'-------------------------------------------------------------------------------
' 個人用設定の「スタート」の「～最近開いた項目を表示する」の設定値を返す
' 戻り値：  0=表示しない、1=表示する
'-------------------------------------------------------------------------------
Private Function Start_TrackDocs() As Long
    With CreateObject("WScript.Shell")
        Start_TrackDocs = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs")
    End With
End Function

'-------------------------------------------------------------------------------
' OneDriveに同期したSharePointファイルのローカルドライブ上のパスを返す
' ThisWorkbook.PathがURLを返す問題への対応
'-------------------------------------------------------------------------------
Public Function ThisWorkbookLocalPath() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        ThisWorkbookLocalPath = ThisWorkbook.Path
        Exit Function
    End If
    
    If Start_TrackDocs = 0 Then
        MsgBox "最近使った項目の表示が有効になっていません。"
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

