Attribute VB_Name = "StandardMod"
Option Compare Database
Option Explicit
Sub test()
'Dim Arr As New DopeArray
'    Arr.AddNew Split(Environ("path"), ";")
    
    
    
Dim Vchk As New DopeVersionCtrl
    Vchk.OpenRepoPath
    
    
End Sub
Public Function F_ProjName()
Dim Vctrl As New DopeVersionCtrl
    F_ProjName = Vctrl.ProjName
End Function
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
