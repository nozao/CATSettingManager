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
        MsgBox "�ݒ肳�ꂽ�t�H���_�����݂��܂���", vbOKOnly + vbCritical, "�o�b�N�A�b�v�t�H���_�G���["
        Exit Sub
    End If
    
    Set oFolder = oFS.GetFolder(BackupFolderPath)
    
    '2�񃊃X�g�ɂ���
    TargetList.ColumnCount = 2
    
    '������
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
        MsgBox "CATIA�N����ԂŎ��s���Ă�������", vbOKOnly + vbCritical, "�o�b�N�A�b�v�G���["
    Else
        Call BackupCATSettings
        
    End If
        
End Sub

Public Sub CatiaBootCheck()

    If Form_BackupTool.List_BackupList.ListIndex = LIST_NO_SELECTED Then
            MsgBox "CATIA�ɓK�p�������ݒ�o�b�N�A�b�v���I������Ă��܂���" & vbNewLine & _
        "���X�g��I�����čēx���s���Ă�������", vbOKOnly + vbCritical, "CATIA�̏I��"
    Else
        If CheckSettingFolderExists() = True Then
            MsgBox "CATIA���N������Ă��܂�" & vbNewLine & "[OK]�������Ă��̉�ʂ������ACATIA���I�����Ă�������", vbOKOnly + vbCritical, "CATIA�̏I��"
        End If
        Call ShowInfo("�ݒ�t�@�C���̏��������Ă��܂�")
        Call CopyToTempFolder(Form_BackupTool.List_BackupList.List(Form_BackupTool.List_BackupList.ListIndex, 0))
        MsgBox "�ݒ�ύX�̏������ł��܂���" & vbNewLine & "�܂�CATIA���N�����Ă���ꍇ�͏I�����Ă���[OK]�������Ă�������", vbOKOnly + vbInformation, "�ݒ�ύX�O�ŏI�m�F"
        Call ShowInfo(vbNullString)
        If CheckSettingFolderExists() = True Then
            MsgBox "CATIA���N������Ă��邽�ߏ����𒆒f���܂�", vbOKOnly + vbCritical, "CATIA�N���G���["
        Else
            Dim LimitSec As Integer
            LimitSec = ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(COPY_WAIT_TIMEOUT_SECONDS_CELL).Value
            MsgBox "�ݒ�ύX���J�n���܂�" & vbNewLine & "[OK]���������̂��A" & LimitSec & "�b�ȓ���CATIA���N�����Ă�������", vbOKOnly + vbInformation, "�ݒ�ύX�O�ŏI�m�F"
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


