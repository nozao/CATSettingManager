VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_BackupTool 
   Caption         =   "CATSettingsManager"
   ClientHeight    =   2064
   ClientLeft      =   -24
   ClientTop       =   -120
   ClientWidth     =   2340
   OleObjectBlob   =   "Form_BackupTool.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Form_BackupTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Forms")
Option Explicit

Private Sub Command_Backup_Click()
    Call BackupCheck
End Sub

Private Sub Command_PreSetting_Click()
    Call CatiaBootCheck
End Sub

Private Sub UserForm_Activate()
    Call List_Refresh(Me.List_BackupList, ThisWorkbook.Sheets(SHEET_NAME_SETTING).Range(USER_SETTINGFILE_BACKUP_FOLDER_PATH_CELL).Value)
End Sub

