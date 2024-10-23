Option Explicit


' �萔��`
Const SRC_PATH = "..\src\"
Const BIN_PATH = "..\bin\"
Const FILE_EXT_XLSM = "xlsm"


' �I�u�W�F�N�g�ϐ�
Dim objFSO, objExcel
Dim scriptDir


' ���C�������̊J�n
Call Main


Sub Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objExcel = CreateObject("Excel.Application")


    ' �X�N���v�g�̏ꏊ���擾
    scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)


    On Error Resume Next
    If Err.Number <> 0 Then
        LogMessage "�������G���[: " & Err.Description
        Cleanup
        WScript.Quit 1
    End If
    On Error GoTo 0
End Sub


Sub Cleanup()
    If Not objExcel Is Nothing Then
        objExcel.Quit
        Set objExcel = Nothing
    End If
    Set objFSO = Nothing
End Sub


Function GetProjectPath(relativePath)
    GetProjectPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), relativePath)
End Function


Sub LogMessage(message)
    WScript.Echo Now & " " & message
End Sub


Sub Main()
    Initialize


    LogMessage "�C���|�[�g�������J�n���܂�"


    On Error Resume Next


    ' bin�t�H���_�̃p�X���擾
    Dim binPath
    binPath = GetProjectPath("bin")


    If Not objFSO.FolderExists(binPath) Then
        LogMessage "�G���[: bin�t�H���_��������܂���: " & binPath
        Cleanup
        WScript.Quit 1
    End If


    ' bin�t�H���_����Excel�t�@�C��������
    Dim file
    For Each file In objFSO.GetFolder(binPath).Files
        If LCase(objFSO.GetExtensionName(file.Name)) = FILE_EXT_XLSM Then
            LogMessage "�t�@�C�����o: " & file.Name
            Call ImportVBAModules(file)
        End If
    Next


    If Err.Number <> 0 Then
        LogMessage "�G���[����: " & Err.Description
    End If


    LogMessage "�C���|�[�g�������������܂���"
    Cleanup
End Sub


Sub ImportVBAModules(excelFile)
    Dim targetWorkbook, projectFolder, component, moduleFile


    On Error Resume Next


    LogMessage "�����J�n: " & excelFile.Name


    ' �v���W�F�N�g�t�H���_�̃p�X���擾
    projectFolder = GetProjectPath("src\" & objFSO.GetBaseName(excelFile.Name))


    ' �v���W�F�N�g�t�H���_�����݂��Ȃ��ꍇ�̓X�L�b�v
    If Not objFSO.FolderExists(projectFolder) Then
        LogMessage "�X�L�b�v: �\�[�X�t�H���_��������܂���: " & projectFolder
        Exit Sub
    End If


    ' Excel�t�@�C�����J��
    Set targetWorkbook = objExcel.Workbooks.Open(excelFile.Path)
    If HandleError("�t�@�C�����J���܂���: " & excelFile.Path) Then Exit Sub


    ' VBA�v���W�F�N�g�A�N�Z�X�m�F
    If Not CheckVBAAccess(targetWorkbook) Then
        targetWorkbook.Close False
        Exit Sub
    End If


    ' �v���W�F�N�g�ی�̉��������݂�
    On Error Resume Next
    targetWorkbook.VBProject.Protection = 0
    If Err.Number <> 0 Then
        LogMessage "�x��: VBA�v���W�F�N�g�̕ی�������ł��܂���"
        LogMessage "�v���W�F�N�g�̃p�X���[�h�ی���m�F���Ă�������"
        targetWorkbook.Close False
        Exit Sub
    End If
    On Error GoTo 0


    ' �����̃��W���[�����폜
    RemoveExistingModules targetWorkbook


    ' �V�������W���[�����C���|�[�g
    ImportNewModules targetWorkbook, projectFolder


    ' �t�@�C����ۑ����ĕ���
    On Error Resume Next
    targetWorkbook.Save
    If Err.Number <> 0 Then
        LogMessage "�G���[: �t�@�C���̕ۑ��Ɏ��s���܂���"
        LogMessage "�G���[ " & Err.Number & ": " & Err.Description
    End If


    targetWorkbook.Close
    LogMessage "��������: " & excelFile.Name
End Sub


Function HandleError(message)
    If Err.Number <> 0 Then
        LogMessage "�G���[: " & message
        LogMessage "�G���[ " & Err.Number & ": " & Err.Description
        HandleError = True
        Err.Clear
    Else
        HandleError = False
    End If
End Function


Function CheckVBAAccess(wb)
    Dim testAccess


    On Error Resume Next
    Set testAccess = wb.VBProject.VBComponents


    If Err.Number <> 0 Then
        LogMessage "�x��: VBA�v���W�F�N�g�ɃA�N�Z�X�ł��܂���"
        LogMessage "�ȉ���Excel�̃Z�L�����e�B�ݒ���m�F���Ă�������:"
        LogMessage "1. Excel > �I�v�V���� > �Z�L�����e�B�Z���^�["
        LogMessage "2. �Z�L�����e�B�Z���^�[�̐ݒ� > �}�N���̐ݒ�"
        LogMessage "3. VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������ ���I���ɂ���"
        CheckVBAAccess = False
    Else
        CheckVBAAccess = True
    End If


    On Error GoTo 0
End Function


Sub RemoveExistingModules(wb)
    Dim component


    LogMessage "�����̃��W���[�����폜��..."


    On Error Resume Next
    For Each component In wb.VBProject.VBComponents
        Select Case component.Type
            Case 1, 2, 3  ' Standard, Class, Form
                wb.VBProject.VBComponents.Remove component
                If Err.Number = 0 Then
                    LogMessage "�폜: " & component.Name
                Else
                    LogMessage "�x��: " & component.Name & " �̍폜�Ɏ��s"
                    Err.Clear
                End If
        End Select
    Next
    On Error GoTo 0
End Sub


Sub ImportNewModules(wb, folder)
    Dim moduleFile


    LogMessage "�V�������W���[�����C���|�[�g��..."


    For Each moduleFile In objFSO.GetFolder(folder).Files
        Dim extension
        extension = LCase(objFSO.GetExtensionName(moduleFile.Name))


        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            On Error Resume Next
            wb.VBProject.VBComponents.Import moduleFile.Path


            If Err.Number = 0 Then
                LogMessage "�C���|�[�g����: " & moduleFile.Name
            Else
                LogMessage "�x��: " & moduleFile.Name & " �̃C���|�[�g�Ɏ��s"
                LogMessage "�G���[ " & Err.Number & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End Sub
