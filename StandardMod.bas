Attribute VB_Name = "StandardMod"
Option Compare Database
Option Explicit
Public Enum RibbTogg
    ShowIt
    HideIt
End Enum
Sub test()
Dim vCtrl As New DopeVersionCtrl
    vCtrl.OpenRepoPath
End Sub
Public Function F_ProjName()
Dim vCtrl As New DopeVersionCtrl
    F_ProjName = vCtrl.ProjName
End Function
Public Sub S_ToggleRibbon(Direction As RibbTogg)
    Select Case Direction
        Case RibbTogg.HideIt
            DoCmd.ShowToolbar "Ribbon", acToolbarNo
        Case RibbTogg.ShowIt
            DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End Select
End Sub
Public Sub S_UpdateGit()
Dim Git As New DopeVBIDE
    With Git
        'form mods
        .ExportGitFldr vbext_ct_Document
        'class mods
        .ExportGitFldr vbext_ct_ClassModule
        'reg mods
        .ExportGitFldr vbext_ct_StdModule
    End With
End Sub
