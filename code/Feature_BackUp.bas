Attribute VB_Name = "Feature_BackUp"
'@Folder("StowagePlan.feature.back_up")
Option Explicit
Option Private Module

Public Sub BackUpStowagePlan()
    CreateBackUpFolder
    Dim backUpFolderPath            As String
    Dim stowagePlanFileName         As String
    With Application
        backUpFolderPath = Environ$("UserProfile") & .PathSeparator & BACKUP_FOLDR_NAME
    End With
    
    With STOWAG_PLAN_BOOK
        stowagePlanFileName = backUpFolderPath & _
                    Application.PathSeparator & _
                    Format$(Now, "yyyymmdd_hhmmss") & _
                    "_" & _
                    CURRENT_VOY & _
                    "_" & _
                    CURRENT_PORT & _
                    "_" & _
                    .Name
        .SaveCopyAs stowagePlanFileName
    End With
End Sub


