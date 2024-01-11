Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive���VBA��ThisWorkbook.Path��URL��Ԃ�������������
' PowerShell����ThisWorkbook(�������g)�ɃL�[�X�g���[�N�𑗂��ă��[�J���p�X���擾����
' Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
' Send keystrokes from PowerShell to ThisWorkbook (myself) to get local path.
'
' �Q�Ɛݒ�ŁuMicrosoft Forms 2.0 Object Library�v���`�F�b�N����.
' ���C�u�����[���Ȃ��ꍇ�̓_�~�[�̃��[�U�[�t�H�[����ǉ����č폜����΁A
' �����I�ɕ\������`�F�b�N�����.
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
    
    '���Ɏ擾�ς݂ł���΁A�擾�ς݂̒l��Ԃ�
    'If it has already been retrieved, the retrieved value is returned.
    
    Static myLocalPathCache As String, lastUpdated As Date
    If myLocalPathCache <> "" And Now() - lastUpdated <= 30 / 86400 Then
        GetThisWorkbookLocalPath3 = myLocalPathCache
        Exit Function
    End If
    
    '�N���b�v�{�[�h�ɋ󕶎���ݒ肷��
    'Set an empty character in the clipboard
    
    Dim cb As New MSForms.DataObject
    cb.SetText ""
    cb.PutInClipboard
    cb.Clear

    'ThisWorkbook�̃E�B���h�E�^�C�g�����擾���AThisWorkbook��O�ʂɕ\������
    'Get the window title of ThisWorkbook and display ThisWorkbook window in front
    
    Dim myWindowTitle As String
    ThisWorkbook.Activate
    myWindowTitle = Application.Caption
    AppActivate myWindowTitle, True
    ActiveWindow.WindowState = xlNormal
    
    'PowerShell�Ŏ��g�ɑ΂���SendKeys�����s����
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

    '�N���b�v�{�[�h����e�L�X�g���擾����
    'Windows�̃o�b�N�O���E���h�����Ƌ�������ꍇ������̂ōő�3��܂Ń��g���C���s�Ȃ�
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
        
    '���ۂɃt�@�C�������݂��邩�m�F����
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    myLocalPathCache = fso.GetParentFolderName(filePath)
    lastUpdated = Now()
    GetThisWorkbookLocalPath3 = myLocalPathCache

End Function


'-------------------------------------------------------------------------------
' �e�X�g�R�[�h
' Test code for GetThisWorkbookLocalPath3
'-------------------------------------------------------------------------------
Private Sub Test_GetThisWorkbookLocalPath3()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetThisWorkbookLocalPath3
End Sub


'-------------------------------------------------------------------------------
' ���̃��W���[���͂����ŏI���
' The script for this module ends here
'-------------------------------------------------------------------------------

