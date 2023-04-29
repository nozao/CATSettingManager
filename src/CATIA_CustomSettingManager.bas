Attribute VB_Name = "CATIA_CustomSettingManager"
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, ImplicitlyTypedConst
'@Folder "Module"
Option Explicit

'�V�[�g����`
Public Const SHEET_NAME_SETTING = "Settings"

'Settings�V�[�g��̊e��ݒ�l�i�[�Z����`
Public Const COPY_START_TRIGGER_FOLDER_PATH_CELL = "B2"
Public Const CATIA_USER_SETTING_FOLDER_PATH_CELL = "B3"
Public Const USER_SETTINGFILE_BACKUP_FOLDER_PATH_CELL = "B4"
Public Const COPY_WAIT_TIMEOUT_SECONDS_CELL = "B5"
Public Const PRE_COPY_LOCAL_FOLDER_PATH_CELL = "B6"

Enum ERR_NUMBER
    NORMAL_END = 0
    ERR_VOID_PATH
    ERR_TIMEOUT
End Enum

Sub CopyToTempFolder(ByVal sTargetBackupFolderName As String)
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    
    '�w�肳�ꂽ�o�b�N�A�b�v�t�H���_�p�X�𒊏o
    Dim sTargetBackupPath As String
    sTargetBackupPath = oFS.BuildPath(ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(USER_SETTINGFILE_BACKUP_FOLDER_PATH_CELL).Value, sTargetBackupFolderName & "\CATSettings")
    
    '�e���|�����t�H���_�p�X�𒊏o
    Dim sTempFolderPath As String
    sTempFolderPath = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(PRE_COPY_LOCAL_FOLDER_PATH_CELL).Value
    
    If oFS.FolderExists(sTempFolderPath) = False Then
        oFS.CreateFolder sTempFolderPath
    End If
    oFS.CopyFolder sTargetBackupPath, sTempFolderPath
    
    
End Sub

Sub BackupCATSettings()

    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    
    '�o�b�N�A�b�v�p�t�H���_�p�X�쐬
    Dim sNewFolderPath As String
    sNewFolderPath = Format(Now, "yyyymmdd_hhnn_") & "CATSettings"
    sNewFolderPath = oFS.BuildPath(ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(USER_SETTINGFILE_BACKUP_FOLDER_PATH_CELL).Value, sNewFolderPath) & "\"
    
    '�o�b�N�A�b�v�Ώۃt�H���_�p�X�쐬
    Dim sSettingsFolderPath As String
    sSettingsFolderPath = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(CATIA_USER_SETTING_FOLDER_PATH_CELL).Value
    
    '��̃o�b�N�A�b�v�p�t�H���_�����
    oFS.CreateFolder sNewFolderPath
    
    '�R�����g�t�@�C���쐬
    Dim oTs As Object
    Dim sResult As String
    Set oTs = oFS.OpenTextFile(oFS.BuildPath(sNewFolderPath, "BackupDescription.txt"), 2, True)
    sResult = InputBox("�o�b�N�A�b�v�R�����g����͂��Ă�������")
    oTs.WriteLine sResult
    oTs.Close
    
    Call ShowInfo("�o�b�N�A�b�v���s���ł�")
    DoEvents
    
    '�o�b�N�A�b�v���{
    oFS.CopyFolder sSettingsFolderPath, sNewFolderPath
    Set oFS = Nothing
    Call HideInfo
    MsgBox "�o�b�N�A�b�v���������܂���", vbOKOnly + vbInformation, "�o�b�N�A�b�v����"
    Call List_Refresh(Form_BackupTool.List_BackupList, ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(USER_SETTINGFILE_BACKUP_FOLDER_PATH_CELL).Value)
    
    
End Sub


Function CheckSettingFolderExists() As Boolean
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")

    CheckSettingFolderExists = oFS.FolderExists(ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(CATIA_USER_SETTING_FOLDER_PATH_CELL).Value)
    Set oFS = Nothing
    

End Function


Function TargetTriggerWait(CheckFolderPath As String, TimeoutSeconds As Integer)
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    
    Dim TimeoutLimit As Date
    TimeoutLimit = Now() + TimeValue("0:0:" & TimeoutSeconds)
    
    While TimeoutLimit > Now()
        If oFS.FolderExists(CheckFolderPath) = True Then
            TargetTriggerWait = NORMAL_END
            Exit Function
        End If
    Wend
    TargetTriggerWait = ERR_TIMEOUT
    

End Function

Public Sub CATIASettingApply()
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")

    Dim sBootTriggerFolderPath As String
    sBootTriggerFolderPath = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(COPY_START_TRIGGER_FOLDER_PATH_CELL).Value
    
    Dim iLimitBreak As Integer
    iLimitBreak = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(COPY_WAIT_TIMEOUT_SECONDS_CELL).Value
    
    Dim sSettingTargetFolderPath As String
    sSettingTargetFolderPath = oFS.GetParentFolderName(ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(CATIA_USER_SETTING_FOLDER_PATH_CELL).Value)
    
    Select Case TargetTriggerWait(sBootTriggerFolderPath, iLimitBreak)
    Case NORMAL_END
        oFS.CopyFolder ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(PRE_COPY_LOCAL_FOLDER_PATH_CELL).Value, sSettingTargetFolderPath & "\", True
        MsgBox "�ݒ肪�������܂���", vbOKOnly + vbInformation, "�ݒ�ύX��������"
    Case ERR_TIMEOUT
        MsgBox "�w�莞�ԓ���CATIA�̋N�����m�F�ł��܂���ł����B������x��蒼���Ă�������", vbOKOnly + vbCritical, "CATIA�N���G���["
    End Select
    
End Sub

