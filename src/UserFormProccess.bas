Attribute VB_Name = "UserFormProccess"
'@IgnoreModule ImplicitlyTypedConst, FunctionReturnValueAlwaysDiscarded
'@Folder "Module"
Option Explicit

Public Const LIST_NO_SELECTED = -1
'@entrypoint
Public Sub ToolBoot()
    Form_BackupTool.Show
End Sub

Public Sub List_Refresh(TargetList As MSForms.ListBox, BackupFolderPath As String)
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")

    Dim oFolder As Object
    If oFS.FolderExists(BackupFolderPath) = False Then
        MsgBox "設定されたフォルダが存在しません", vbOKOnly + vbCritical, "バックアップフォルダエラー"
        Exit Sub
    End If
    
    Set oFolder = oFS.GetFolder(BackupFolderPath)
    
    '2列リストにする
    TargetList.ColumnCount = 2
    
    '初期化
    TargetList.Clear
    
    Dim CurrentFolder As Object
    
    Dim oTs As Object
    Dim DescriptionFilePath As String
    
    For Each CurrentFolder In oFolder.subFolders
        TargetList.AddItem vbNullString
        TargetList.List(TargetList.ListCount - 1, 0) = CurrentFolder.Name
        
        DescriptionFilePath = oFS.BuildPath(CurrentFolder.Path, "BackupDescription.txt")
        If oFS.FileExists(DescriptionFilePath) = True Then
            Set oTs = oFS.OpenTextFile(DescriptionFilePath)
            TargetList.List(TargetList.ListCount - 1, 1) = oTs.ReadLine
            oTs.Close
        End If
    Next

End Sub
Public Sub BackupCheck()
    If CheckSettingFolderExists() = False Then
        MsgBox "CATIA起動状態で実行してください", vbOKOnly + vbCritical, "バックアップエラー"
    Else
        Call BackupCATSettings
        
    End If
        
End Sub

Public Sub CatiaBootCheck()

    If Form_BackupTool.List_BackupList.ListIndex = LIST_NO_SELECTED Then
            MsgBox "CATIAに適用したい設定バックアップが選択されていません" & vbNewLine & _
        "リストを選択して再度実行してください", vbOKOnly + vbCritical, "CATIAの終了"
    Else
        If CheckSettingFolderExists() = True Then
            MsgBox "CATIAが起動されています" & vbNewLine & "[OK]を押してこの画面を閉じた後、CATIAを終了してください", vbOKOnly + vbCritical, "CATIAの終了"
        End If
        Call ShowInfo("設定ファイルの準備をしています")
        Call CopyToTempFolder(Form_BackupTool.List_BackupList.List(Form_BackupTool.List_BackupList.ListIndex, 0))
        MsgBox "設定変更の準備ができました" & vbNewLine & "まだCATIAが起動している場合は終了してから[OK]を押してください", vbOKOnly + vbInformation, "設定変更前最終確認"
        Call ShowInfo(vbNullString)
        If CheckSettingFolderExists() = True Then
            MsgBox "CATIAが起動されているため処理を中断します", vbOKOnly + vbCritical, "CATIA起動エラー"
        Else
            Dim LimitSec As Integer
            LimitSec = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(COPY_WAIT_TIMEOUT_SECONDS_CELL).Value
            MsgBox "設定変更を開始します" & vbNewLine & "[OK]を押したのち、" & LimitSec & "秒以内にCATIAを起動してください", vbOKOnly + vbInformation, "設定変更前最終確認"
            Call CATIASettingApply
            Call HideInfo
            
        End If
    End If
    
End Sub
Public Sub ShowInfo(sInformation As String)
    With Form_BackupTool
        .Command_Backup.Enabled = False
        .Command_PreSetting.Enabled = False
        .List_BackupList.Enabled = False
        .Label_Info.Caption = sInformation
        DoEvents
    End With
End Sub

Public Sub HideInfo()
    With Form_BackupTool
        .Command_Backup.Enabled = True
        .Command_PreSetting.Enabled = True
        .List_BackupList.Enabled = True
        .Label_Info.Caption = vbNullString
        DoEvents
    End With
End Sub


