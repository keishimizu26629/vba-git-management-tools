Option Explicit


' 定数定義
Const SRC_PATH = "..\src\"
Const BIN_PATH = "..\bin\"
Const FILE_EXT_XLSM = "xlsm"


' オブジェクト変数
Dim objFSO, objExcel
Dim scriptDir


' メイン処理の開始
Call Main


Sub Initialize()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objExcel = CreateObject("Excel.Application")


    ' スクリプトの場所を取得
    scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)


    On Error Resume Next


    ' srcフォルダの作成（存在しない場合）
    Dim srcFullPath
    srcFullPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), "src")
    If Not objFSO.FolderExists(srcFullPath) Then
        objFSO.CreateFolder srcFullPath
        LogMessage "作成したフォルダ: " & srcFullPath
    End If


    If Err.Number <> 0 Then
        LogMessage "初期化エラー: " & Err.Description
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
    ' スクリプトの親フォルダ（tools）の親フォルダからの相対パスを解決
    GetProjectPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), relativePath)
End Function


Sub LogMessage(message)
    WScript.Echo Now & " " & message
End Sub


Sub Main()
    Initialize


    LogMessage "処理を開始します"


    On Error Resume Next


    ' binフォルダのパスを取得
    Dim binPath
    binPath = GetProjectPath("bin")


    If Not objFSO.FolderExists(binPath) Then
        LogMessage "エラー: binフォルダが見つかりません: " & binPath
        Cleanup
        WScript.Quit 1
    End If


    ' binフォルダ内のExcelファイルを処理
    Dim file
    For Each file In objFSO.GetFolder(binPath).Files
        If LCase(objFSO.GetExtensionName(file.Name)) = FILE_EXT_XLSM Then
            LogMessage "ファイル検出: " & file.Name
            Call ExportVBAModules(file)
        End If
    Next


    If Err.Number <> 0 Then
        LogMessage "エラー発生: " & Err.Description
    End If
    On Error GoTo 0


    LogMessage "処理が完了しました"
    Cleanup
End Sub


Sub ExportVBAModules(excelFile)
    Dim targetWorkbook, projectFolder, component


    On Error Resume Next


    LogMessage "処理開始: " & excelFile.Name


    ' Excelファイルを開く
    Set targetWorkbook = objExcel.Workbooks.Open(excelFile.Path)
    If HandleError("ファイルを開けません: " & excelFile.Path) Then Exit Sub


    ' srcフォルダ内にExcelファイルと同名のフォルダを作成
    Dim baseName
    baseName = objFSO.GetBaseName(excelFile.Name)
    projectFolder = GetProjectPath("src\" & baseName)


    If Not objFSO.FolderExists(projectFolder) Then
        objFSO.CreateFolder projectFolder
        LogMessage "作成したフォルダ: " & projectFolder
    End If


    ' VBAプロジェクトアクセス確認
    If Not CheckVBAAccess(targetWorkbook) Then
        targetWorkbook.Close False
        Exit Sub
    End If


    ' モジュールのエクスポート
    Call ExportModules(targetWorkbook, projectFolder)


    targetWorkbook.Close False
    LogMessage "処理完了: " & excelFile.Name
End Sub


Function HandleError(message)
    If Err.Number <> 0 Then
        LogMessage "エラー: " & message
        LogMessage "エラー " & Err.Number & ": " & Err.Description
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
        LogMessage "警告: VBAプロジェクトにアクセスできません"
        LogMessage "以下のExcelのセキュリティ設定を確認してください:"
        LogMessage "1. Excel > オプション > セキュリティセンター"
        LogMessage "2. セキュリティセンターの設定 > マクロの設定"
        LogMessage "3. VBAプロジェクトオブジェクトモデルへのアクセスを信頼する をオンにする"
        CheckVBAAccess = False
    Else
        CheckVBAAccess = True
    End If


    On Error GoTo 0
End Function


Sub ExportModules(wb, folder)
    Dim component, extension


    LogMessage "モジュールのエクスポートを開始"


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
                LogMessage "エクスポート完了: " & component.Name & "." & extension
            Else
                LogMessage "警告: モジュール " & component.Name & " のエクスポートに失敗"
                LogMessage "エラー " & Err.Number & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End Sub
