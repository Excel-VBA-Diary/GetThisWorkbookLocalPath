Attribute VB_Name = "Module2"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive���VBA��ThisWorkbook.Path��URL��Ԃ�������������
'�J���Ă���G�N�X�v���[�����烍�[�J���p�X���擾����
'Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
'Get local path from open explorer.
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath2() As String

    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath2 = ThisWorkbook.Path
        Exit Function
    End If
    
    '���Ɏ擾�ς݂ł���΁A�擾�ς݂̒l��Ԃ�
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
                
    '���ۂɃt�@�C�������݂��邩�m�F����
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(strPath & "\" & ThisWorkbook.Name) Then Exit Function
    myLocalPath = strPath
    GetThisWorkbookLocalPath2 = myLocalPath
                
End Function


'-------------------------------------------------------------------------------
' �G���R�[�h���ꂽURL���f�R�[�h����iENCODEURL�֐��̋t�ϊ��j
' �Q�Ɛݒ�ŁuMicrosoft HTML Object Library�v���`�F�b�N���邱��
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
' �G���R�[�h���ꂽURL���f�R�[�h����iASCII�����̂݁j
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
' �e�X�g�R�[�h
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
' ���̃��W���[���͂����ŏI���
' The script for this module ends here
'-------------------------------------------------------------------------------
