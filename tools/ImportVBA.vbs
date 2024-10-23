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
    GetProjectPath = objFSO.BuildPath(objFSO.GetParentFolderName(scriptDir), relativePath)
End Function


Sub LogMessage(message)
    WScript.Echo Now & " " & message
End Sub


Sub Main()
    Initialize


    LogMessage "インポート処理を開始します"


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
            Call ImportVBAModules(file)
        End If
    Next


    If Err.Number <> 0 Then
        LogMessage "エラー発生: " & Err.Description
    End If


    LogMessage "インポート処理が完了しました"
    Cleanup
End Sub


Sub ImportVBAModules(excelFile)
    Dim targetWorkbook, projectFolder, component, moduleFile


    On Error Resume Next


    LogMessage "処理開始: " & excelFile.Name


    ' プロジェクトフォルダのパスを取得
    projectFolder = GetProjectPath("src\" & objFSO.GetBaseName(excelFile.Name))


    ' プロジェクトフォルダが存在しない場合はスキップ
    If Not objFSO.FolderExists(projectFolder) Then
        LogMessage "スキップ: ソースフォルダが見つかりません: " & projectFolder
        Exit Sub
    End If


    ' Excelファイルを開く
    Set targetWorkbook = objExcel.Workbooks.Open(excelFile.Path)
    If HandleError("ファイルを開けません: " & excelFile.Path) Then Exit Sub


    ' VBAプロジェクトアクセス確認
    If Not CheckVBAAccess(targetWorkbook) Then
        targetWorkbook.Close False
        Exit Sub
    End If


    ' プロジェクト保護の解除を試みる
    On Error Resume Next
    targetWorkbook.VBProject.Protection = 0
    If Err.Number <> 0 Then
        LogMessage "警告: VBAプロジェクトの保護を解除できません"
        LogMessage "プロジェクトのパスワード保護を確認してください"
        targetWorkbook.Close False
        Exit Sub
    End If
    On Error GoTo 0


    ' 既存のモジュールを削除
    RemoveExistingModules targetWorkbook


    ' 新しいモジュールをインポート
    ImportNewModules targetWorkbook, projectFolder


    ' ファイルを保存して閉じる
    On Error Resume Next
    targetWorkbook.Save
    If Err.Number <> 0 Then
        LogMessage "エラー: ファイルの保存に失敗しました"
        LogMessage "エラー " & Err.Number & ": " & Err.Description
    End If


    targetWorkbook.Close
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


Sub RemoveExistingModules(wb)
    Dim component


    LogMessage "既存のモジュールを削除中..."


    On Error Resume Next
    For Each component In wb.VBProject.VBComponents
        Select Case component.Type
            Case 1, 2, 3  ' Standard, Class, Form
                wb.VBProject.VBComponents.Remove component
                If Err.Number = 0 Then
                    LogMessage "削除: " & component.Name
                Else
                    LogMessage "警告: " & component.Name & " の削除に失敗"
                    Err.Clear
                End If
        End Select
    Next
    On Error GoTo 0
End Sub


Sub ImportNewModules(wb, folder)
    Dim moduleFile


    LogMessage "新しいモジュールをインポート中..."


    For Each moduleFile In objFSO.GetFolder(folder).Files
        Dim extension
        extension = LCase(objFSO.GetExtensionName(moduleFile.Name))


        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            On Error Resume Next
            wb.VBProject.VBComponents.Import moduleFile.Path


            If Err.Number = 0 Then
                LogMessage "インポート完了: " & moduleFile.Name
            Else
                LogMessage "警告: " & moduleFile.Name & " のインポートに失敗"
                LogMessage "エラー " & Err.Number & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next
End Sub
