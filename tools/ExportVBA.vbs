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


    ' src�t�H���_�̍쐬�i���݂��Ȃ��ꍇ�j
    Dim srcFullPath
    srcFullPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), "src")
    If Not objFSO.FolderExists(srcFullPath) Then
        objFSO.CreateFolder srcFullPath
        LogMessage "�쐬�����t�H���_: " & srcFullPath
    End If


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
    ' �X�N���v�g�̐e�t�H���_�itools�j�̐e�t�H���_����̑��΃p�X������
    GetProjectPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), relativePath)
End Function


Sub LogMessage(message)
    WScript.Echo Now & " " & message
End Sub


Sub Main()
    Initialize


    LogMessage "�������J�n���܂�"


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
            Call ExportVBAModules(file)
        End If
    Next


    If Err.Number <> 0 Then
        LogMessage "�G���[����: " & Err.Description
    End If
    On Error GoTo 0


    LogMessage "�������������܂���"
    Cleanup
End Sub


Sub ExportVBAModules(excelFile)
    Dim targetWorkbook, projectFolder, component


    On Error Resume Next


    LogMessage "�����J�n: " & excelFile.Name


    ' Excel�t�@�C�����J��
    Set targetWorkbook = objExcel.Workbooks.Open(excelFile.Path)
    If HandleError("�t�@�C�����J���܂���: " & excelFile.Path) Then Exit Sub


    ' src�t�H���_����Excel�t�@�C���Ɠ����̃t�H���_���쐬
    Dim baseName
    baseName = objFSO.GetBaseName(excelFile.Name)
    projectFolder = GetProjectPath("src\" & baseName)


    If Not objFSO.FolderExists(projectFolder) Then
        objFSO.CreateFolder projectFolder
        LogMessage "�쐬�����t�H���_: " & projectFolder
    End If


    ' VBA�v���W�F�N�g�A�N�Z�X�m�F
    If Not CheckVBAAccess(targetWorkbook) Then
        targetWorkbook.Close False
        Exit Sub
    End If


    ' ���W���[���̃G�N�X�|�[�g
    Call ExportModules(targetWorkbook, projectFolder)


    targetWorkbook.Close False
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


Sub ExportModules(wb, folder)
    Dim component, extension


    LogMessage "���W���[���̃G�N�X�|�[�g���J�n"


    For Each component In wb.VBProject.VBComponents
        Select Case component.Type
            Case 1: extension = "bas"  ' Standard Module
            Case 2: extension = "cls"  ' Class Module
            Case 3: extension = "frm"  ' UserForm
            Case Else: extension = ""
        End Select


        If extension <> "" Then
            On Error Resume Next
            Dim exportPath
            exportPath = objFSO.BuildPath(folder, component.Name & "." & extension)
            component.Export exportPath


            If Err.Number = 0 Then
                LogMessage "�G�N�X�|�[�g����: " & component.Name & "." & extension
            Else
                LogMessage "�x��: ���W���[�� " & component.Name & " �̃G�N�X�|�[�g�Ɏ��s"
                LogMessage "�G���[ " & Err.Number & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End Sub
