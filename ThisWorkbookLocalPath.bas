Attribute VB_Name = "ThisWorkbookLocalPath"
Option Explicit

Private Sub Test_ThisWorkbookLocalPath()
    Debug.Print ThisWorkbookLocalPath
End Sub

'-------------------------------------------------------------------------------
' �t�H���_�[�I�v�V�����́u�o�^����Ă���g���q�͕\�����Ȃ��v�̐ݒ�l��Ԃ�
' �߂�l�F  0=�\������A1=�\�����Ȃ�
'-------------------------------------------------------------------------------
Private Function HideFileExt() As Long
    With CreateObject("WScript.Shell")
        HideFileExt = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\HideFileExt")
    End With
End Function

'-------------------------------------------------------------------------------
' �l�p�ݒ�́u�X�^�[�g�v�́u�`�ŋߊJ�������ڂ�\������v�̐ݒ�l��Ԃ�
' �߂�l�F  0=�\�����Ȃ��A1=�\������
'-------------------------------------------------------------------------------
Private Function Start_TrackDocs() As Long
    With CreateObject("WScript.Shell")
        Start_TrackDocs = .RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs")
    End With
End Function

'-------------------------------------------------------------------------------
' OneDrive�ɓ�������SharePoint�t�@�C���̃��[�J���h���C�u��̃p�X��Ԃ�
' ThisWorkbook.Path��URL��Ԃ����ւ̑Ή�
'-------------------------------------------------------------------------------
Public Function ThisWorkbookLocalPath() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        ThisWorkbookLocalPath = ThisWorkbook.Path
        Exit Function
    End If
    
    If Start_TrackDocs = 0 Then
        MsgBox "�ŋߎg�������ڂ̕\�����L���ɂȂ��Ă��܂���B"
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

