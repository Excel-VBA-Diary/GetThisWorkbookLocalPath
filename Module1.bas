Attribute VB_Name = "Module1"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
'最近開いた項目からローカルパスを取得する
'Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
'Get local path from recently opened items.
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath1() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath1 = ThisWorkbook.Path
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim recentFolderPath As String
    recentFolderPath = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Windows\Recent\"
    
    Dim baseName As String, recentFileName As String
    baseName = fso.GetBaseName(ThisWorkbook.Name)
    Select Case True
        Case fso.FileExists(recentFolderPath & ThisWorkbook.Name & ".LNK")
            recentFileName = ThisWorkbook.Name & ".LNK"
        Case fso.FileExists(recentFolderPath & baseName & ".LNK")
            recentFileName = baseName & ".LNK"
        Case Else
            Exit Function
    End Select
    
    Dim filePath As String
    filePath = CreateObject("WScript.Shell").CreateShortcut(recentFolderPath & recentFileName).TargetPath
    
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    If Not fso.FileExists(filePath) Then Exit Function
    GetThisWorkbookLocalPath1 = fso.GetParentFolderName(filePath)
        
End Function


'-------------------------------------------------------------------------------
' 個人用設定の「スタート」の「～最近開いた項目を表示する」の設定値を返す
' 戻り値：  Fase=表示しない、True=表示する
' Returns the value of the "Show Recently Opened Items" setting in "Start" of personal settings.
' Return value: Fase=Do not display, True=display
'-------------------------------------------------------------------------------
Public Function Is_Start_TrackDocs() As Boolean
    Dim errorNumber As Long
    With CreateObject("WScript.Shell")
        On Error Resume Next
        Is_Start_TrackDocs = CBool(.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs"))
        If Err.Number <> 0 Then Is_Start_TrackDocs = False
        On Error GoTo 0
    End With
End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath1()
    Dim i As Long, result As String
    For i = 1 To 10
        result = GetThisWorkbookLocalPath1()
        Debug.Print Time, i, result
    Next
End Sub


'-------------------------------------------------------------------------------
' 標準モジュールはここで終わり
'-------------------------------------------------------------------------------
