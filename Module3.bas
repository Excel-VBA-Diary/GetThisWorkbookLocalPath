Attribute VB_Name = "Module3"
Option Explicit

'-------------------------------------------------------------------------------
'OneDrive���VBA��ThisWorkbook.Path��URL��Ԃ�������������
'PowerShell���玩�����g�ɃL�[�X�g���[�N�𑗂��ă��[�J���p�X���擾����
'�Q�Ɛݒ�ŁuMicrosoft Forms 2.0 Object Library�v���`�F�b�N���邱��
'Resolve problem with ThisWorkbook.Path returning URL in VBA on OneDrive.
'Send keystrokes from PowerShell to myself to get local path.
'Prerequisite: Check "Microsoft Forms 2.0 Object Library" in the References dialog box.
'-------------------------------------------------------------------------------

Public Function GetThisWorkbookLocalPath3() As String
    
    If Not ThisWorkbook.Path Like "http*" Then
        GetThisWorkbookLocalPath3 = ThisWorkbook.Path
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
        
    '���ۂɃt�@�C�������݂��邩�m�F����
    'Verify that the file actually exists
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    GetThisWorkbookLocalPath3 = fso.GetParentFolderName(filePath)

End Function


'-------------------------------------------------------------------------------
' �e�X�g�R�[�h
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
' �W�����W���[���͂����ŏI���
'-------------------------------------------------------------------------------

