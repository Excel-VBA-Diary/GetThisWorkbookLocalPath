Attribute VB_Name = "Module2"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive上のVBAでThisWorkbook.PathがURLを返す問題を解決する
' 開いているエクスプローラからローカルパスを取得する
' Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
' Get local path from open explorer.
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
' Last Updated: January, 14, 2024
' Version: 1.003
' License: MIT
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath2() As String

    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath2 = ThisWorkbook.Path
        Exit Function
    End If
    
    '既に取得済みであれば、取得済みの値を返す
    'If it has already been retrieved, the retrieved value is returned.
    
    Static myLocalPathCache As String, lastUpdated As Date
    If myLocalPathCache <> "" And Now() - lastUpdated <= 30 / 86400 Then
        GetThisWorkbookLocalPath2 = myLocalPathCache
        Exit Function
    End If
    
    Dim myLocalPath As String, urlFolderName As String, wObj As Object
    Dim tempArray As Variant, tempLocalPath As String, tempFolderName As String
    Select Case True
        Case LCase(ThisWorkbook.Path) Like "https://d.docs.live.net/????????????????"
            myLocalPath = Environ("OneDrive")
        Case LCase(ThisWorkbook.Path) Like "https://*-my.sharepoint.com/personal/*/documents"
            myLocalPath = Environ("OneDriveCommercial")
        Case Else
            urlFolderName = Mid(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "/") + 1)
            '日本語補正
            If LCase(urlFolderName) = "shared documents" Then urlFolderName = "ドキュメント"
            For Each wObj In CreateObject("Shell.Application").Windows
                If LCase(wObj.FullName) Like "*explorer.exe" Then
                    tempLocalPath = DecodeURL_ASCII(wObj.LocationURL)
                    tempLocalPath = Replace(tempLocalPath, "file:///", "")
                    tempLocalPath = Replace(tempLocalPath, "/", "\")
                    tempArray = Split(wObj.LocationName, " - ")
                    If UBound(tempArray) = 1 Then
                        If tempLocalPath Like Environ("OneDriveCommercial") & "*" Then
                            'OneDrive for Business (Cloud Icon)
                            tempFolderName = tempArray(0)
                        Else
                            'SharePoint sync folder (Building Icon)
                            tempFolderName = tempArray(1)
                        End If
                    Else
                        tempFolderName = wObj.LocationName
                    End If
                    If tempFolderName = urlFolderName Then
                        myLocalPath = tempLocalPath
                        Exit For
                    End If
                End If
            Next
    End Select
    
    If myLocalPath = "" Then Exit Function
                
    '実際にファイルが存在するか確認する
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(myLocalPath & "\" & ThisWorkbook.Name) Then Exit Function
    myLocalPathCache = myLocalPath
    lastUpdated = Now()
    GetThisWorkbookLocalPath2 = myLocalPathCache

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
' エンコードされたURLをデコードする（ENCODEURL関数の逆変換）
' Decode encoded URL (reverse conversion of ENCODEURL function)

' DecodeURL_ASCII関数の代わりにこの関数を使う場合は
' 参照設定で「Microsoft HTML Object Library」をチェックすること.
' If you use this function instead of the DecodeURL_ASCII function,
' Check the "Microsoft HTML Object Library" in the references dialog box.
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
' テストコード
' Test code for GetThisWorkbookLocalPath2
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath2()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetThisWorkbookLocalPath2
End Sub


'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------
