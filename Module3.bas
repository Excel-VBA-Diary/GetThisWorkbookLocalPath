Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
'PowerShellから自分自身にキーストロークを送ってローカルパスを取得する
'参照設定で「Microsoft Forms 2.0 Object Library」をチェックすること
'Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
'Send keystrokes from PowerShell to myself to get local path.
'Prerequisite: Check "Microsoft Forms 2.0 Object Library" in the References dialog box.
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath3() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath3 = ThisWorkbook.Path
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
        
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    GetThisWorkbookLocalPath3 = fso.GetParentFolderName(filePath)

End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath3()
    Dim i As Long, result As String
    For i = 1 To 10
        result = GetThisWorkbookLocalPath3()
        Debug.Print Time, i, result
    Next
End Sub


'-------------------------------------------------------------------------------
' 標準モジュールはここで終わり
'-------------------------------------------------------------------------------

