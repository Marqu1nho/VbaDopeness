Attribute VB_Name = "StandardMod"
Option Compare Database
Option Explicit

Public Function F_ProjName()
Dim Vctrl As New DopeVersionCtrl
    F_ProjName = Vctrl.ProjName
End Function
Public Sub S_UpdateGit()
Dim Obj As New DopeVBIDE
    Obj.ExportGitFldr vbext_ct_StdModule
End Sub
