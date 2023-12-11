Attribute VB_Name = "Module1"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive���VBA��ThisWorkbook.Path��URL��Ԃ�������������
'�ŋߊJ�������ڂ��烍�[�J���p�X���擾����
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
    
    '���ۂɃt�@�C�������݂��邩�m�F����
    'Verify that the file actually exists
    
    If Not fso.FileExists(filePath) Then Exit Function
    GetThisWorkbookLocalPath1 = fso.GetParentFolderName(filePath)
        
End Function


'-------------------------------------------------------------------------------
' �l�p�ݒ�́u�X�^�[�g�v�́u�`�ŋߊJ�������ڂ�\������v�̐ݒ�l��Ԃ�
' �߂�l�F  Fase=�\�����Ȃ��ATrue=�\������
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
' �e�X�g�R�[�h
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
' �W�����W���[���͂����ŏI���
'-------------------------------------------------------------------------------
