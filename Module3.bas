Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
' PowerShellからThisWorkbook(自分自身)にキーストロークを送ってローカルパスを取得する
' Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
' Send keystrokes from PowerShell to ThisWorkbook (myself) to get local path.
'
' 参照設定で「Microsoft Forms 2.0 Object Library」をチェックする.
' ライブラリーがない場合はダミーのユーザーフォームを追加して削除すれば、
' 自動的に表示されチェックされる.
' Check the "Microsoft Forms 2.0 Object Library" in the References dialog box.
' If the library is not in the dialog box, add a dummy user form and remove it,
' then the library will automatically appear and be checked.
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
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath3() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath3 = ThisWorkbook.Path
        Exit Function
    End If
    
    '既に取得済みであれば、取得済みの値を返す
    'If it has already been retrieved, the retrieved value is returned.
    
    Static myLocalPathCache As String, lastUpdated As Date
    If myLocalPathCache <> "" And Now() - lastUpdated <= 30 / 86400 Then
        GetThisWorkbookLocalPath3 = myLocalPathCache
        Exit Function
    End If
    
    'クリップボードに空文字を設定する
    'Set an empty character in the clipboard
    
    Dim cb As New MSForms.DataObject
    cb.SetText ""
    cb.PutInClipboard
    cb.Clear

    'ThisWorkbookのウィンドウタイトルを取得し、ThisWorkbookを前面に表示する
    'Get the window title of ThisWorkbook and display ThisWorkbook window in front
    
    Dim myWindowTitle As String
    ThisWorkbook.Activate
    myWindowTitle = Application.Caption
    AppActivate myWindowTitle, True
    ActiveWindow.WindowState = xlNormal
    
    'PowerShellで自身に対してSendKeysを実行する
    'Run SendKeys to myself in PowerShell
    
    Dim wScript As String
    wScript = "PowerShell.exe -Command """ & _
              "Add-type -AssemblyName Microsoft.VisualBasic;" & _
              "Add-Type -AssemblyName System.Windows.Forms;" & _
              "[Microsoft.VisualBasic.Interaction]::AppActivate('" & myWindowTitle & "');" & _
              "Start-Sleep -Milliseconds 100;" & _
              "[System.Windows.Forms.SendKeys]::SendWait('%');" & _
              "Start-Sleep -Milliseconds 100;" & _
              "[System.Windows.Forms.SendKeys]::SendWait('FIL');" & _
              "[System.Windows.Forms.SendKeys]::SendWait('%');" & _
              "Start-Sleep -Milliseconds 100;" & _
              "[System.Windows.Forms.SendKeys]::SendWait('H{UP}{ENTER}');" & _
              "Start-Sleep -Milliseconds 100;"""
              
    Call CreateObject("Wscript.shell").Run(wScript, 0, True)

    'クリップボードからテキストを取得する
    'Windowsのバックグラウンド処理と競合する場合があるので最大3回までリトライを行なう
    'Get text from clipboard
    'Retry up to 3 times to avoid conflicts with Windows background processing
    
    Dim filePath As String, retryCount As Long, errNo As Long
    retryCount = 0
    Do
        On Error Resume Next
        cb.GetFromClipboard
        errNo = Err.Number
        On Error GoTo 0
        If errNo = 0 Then Exit Do
        retryCount = retryCount + 1
        If retryCount > 3 Then Exit Function
        Debug.Print Time, "Retry GetFromClipboard, retry count(s)="; retryCount
        Application.Wait [NOW()+"00:00:00.1"]
    Loop
    
    filePath = cb.GetText
    If filePath = "" Then Exit Function
        
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    myLocalPathCache = fso.GetParentFolderName(filePath)
    lastUpdated = Now()
    GetThisWorkbookLocalPath3 = myLocalPathCache

End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code for GetThisWorkbookLocalPath3
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath3()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetThisWorkbookLocalPath3
End Sub


'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------

