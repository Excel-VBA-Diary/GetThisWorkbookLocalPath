Attribute VB_Name = "Module2"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
'開いているエクスプローラからローカルパスを取得する
'Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
'Get local path from open explorer.
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath2() As String

    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath2 = ThisWorkbook.Path
        Exit Function
    End If
    
    '既に取得済みであれば、取得済みの値を返す
    'If it has already been retrieved, the retrieved value is returned.
    
    Static myLocalPath As String
    If myLocalPath <> "" Then
        GetThisWorkbookLocalPath2 = myLocalPath
        Exit Function
    End If
    
    Dim strPath As String, myLocationName As String, wObj As Object
    Select Case True
        Case LCase(ThisWorkbook.Path) Like "https://d.docs.live.net/????????????????"
            strPath = Environ("OneDrive")
        Case LCase(ThisWorkbook.Path) Like "https://*.sharepoint.com/personal/*microsoft_com/documents"
            strPath = Environ("OneDriveCommercial")
        Case Else
            myLocationName = Mid(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "/") + 1)
            For Each wObj In CreateObject("Shell.Application").Windows
                If LCase(wObj.FullName) Like "*explorer.exe" Then
                    If wObj.LocationName = myLocationName Then
                        strPath = DecodeURL(wObj.LocationURL)
                        strPath = Replace(strPath, "file:///", "")
                        strPath = Replace(strPath, "/", "\")
                        Exit For
                    End If
                End If
            Next
    End Select
    
    If strPath = "" Then Exit Function
                
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(strPath & "\" & ThisWorkbook.Name) Then Exit Function
    myLocalPath = strPath
    GetThisWorkbookLocalPath2 = myLocalPath
                
End Function


'-------------------------------------------------------------------------------
' エンコードされたURLをデコードする（ENCODEURL関数の逆変換）
' 参照設定で「Microsoft HTML Object Library」をチェックすること
' Decode encoded URL (reverse conversion of ENCODEURL function)
' Prerequisite: Check "Microsoft HTML Object Library" in the References dialog box.
'-------------------------------------------------------------------------------
Private Function DecodeURL(ByVal URL As String) As String
    If URL = "" Then Exit Function
    Dim htmlDoc As New MSHTML.HTMLDocument
    Dim span As MSHTML.HTMLSpanElement
    Set span = htmlDoc.createElement("span")
    span.setAttribute "id", "result"
    htmlDoc.appendChild span
    htmlDoc.parentWindow.execScript "document.getElementById('result').innerText = " & _
                                    "decodeURIComponent('" & URL & "');"
    DecodeURL = span.innerText
End Function


'-------------------------------------------------------------------------------
' エンコードされたURLをデコードする（ASCII文字のみ）
' Decode encoded URL (ASCII characters only)
'-------------------------------------------------------------------------------
Private Function DecodeURL_ASCII(ByVal URL As String) As String
    If URL = "" Then Exit Function
    Dim i As Long, v As Integer
    i = 1
    Do While i < Len(URL)
        i = InStr(i, URL, "%")
        If i = 0 Then Exit Do
        v = Val("&H" & Mid(URL, i + 1, 2))
        If v >= 32 And v <= 126 Then URL = Replace(URL, Mid(URL, i, 3), Chr(v))
        i = i + 1
    Loop
    DecodeURL_ASCII = URL
End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code for GetThisWorkbookLocalPath2
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath2()
    Dim i As Long, result As String
    For i = 1 To 10
        result = GetThisWorkbookLocalPath2()
        Debug.Print Time, i, result
    Next
End Sub


'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------
