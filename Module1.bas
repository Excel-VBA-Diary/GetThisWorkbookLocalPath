Attribute VB_Name = "Module1"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
' 最近開いた項目からローカルパスを取得する
' Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
' Get local path from recently opened items.
'
' Arguments: Nothing
'
' Return Value:
'   Local Path of ThisWorkbook (String)
'   Return null string if fails conversion from URL path to local path.
'
' Usage:
'   Dim lp As String
'   lp = GetThisWorkbookLocalPath2
'
' Author: Excel VBA Diary (@excelvba_diary)
' Created: December 11, 2023
' Last Updated: January, 11, 2024
' Version: 1.002
' License: MIT
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath1() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath1 = ThisWorkbook.Path
        Exit Function
    End If
    
    '既に取得済みであれば、取得済みの値を返す
    'If it has already been retrieved, the retrieved value is returned.
    
    Static myLocalPathCache As String, lastUpdated As Date
    If myLocalPathCache <> "" And Now() - lastUpdated <= 30 / 86400 Then
        GetThisWorkbookLocalPath1 = myLocalPathCache
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim recentFolderPath As String
    recentFolderPath = Environ("USERPROFILE") & "\AppData\Roaming\Microsoft\Windows\Recent\"
    
    Dim baseName As String, lnkFilePath As String
    baseName = fso.GetBaseName(ThisWorkbook.Name)
    Select Case True
        Case fso.FileExists(recentFolderPath & ThisWorkbook.Name & ".LNK")
            lnkFilePath = recentFolderPath & ThisWorkbook.Name & ".LNK"
        Case fso.FileExists(recentFolderPath & baseName & ".LNK")
            lnkFilePath = recentFolderPath & baseName & ".LNK"
        Case Else
            ' No LNK file exists.
            Exit Function
    End Select
    
    Dim filePath As String
    filePath = CreateObject("WScript.Shell").CreateShortcut(lnkFilePath).TargetPath
    
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    If Not fso.FileExists(filePath) Then Exit Function
    myLocalPathCache = fso.GetParentFolderName(filePath)
    lastUpdated = Now()
    GetThisWorkbookLocalPath1 = myLocalPathCache
        
End Function


'-------------------------------------------------------------------------------
' 個人用設定の「スタート」の「～最近開いた項目を表示する」の設定値を返す
' 戻り値：  Fase=表示しない、True=表示する
' Returns the value of the "Show Recently Opened Items" setting in "Start" of personal settings.
' Return value: Fase=Do not display, True=display
'-------------------------------------------------------------------------------
Public Function Is_Start_TrackDocs() As Boolean
    Dim registryKey As String
    registryKey = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs"
    With CreateObject("WScript.Shell")
        On Error Resume Next
        Is_Start_TrackDocs = CBool(.regRead(registryKey))
        If Err.Number <> 0 Then Is_Start_TrackDocs = False
        On Error GoTo 0
    End With
End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code for GetThisWorkbookLocalPath1
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath1()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetThisWorkbookLocalPath1
End Sub


'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------
