VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Const ToolTbl As String = "tmp_ToolsMenu"
Private vCtrl As New DopeVersionCtrl
Private Enum RibbTogg
    ShowIt
    HideIt
End Enum

Private Sub btnClose_Click()
    Application.Quit
End Sub

Private Sub btnInstall_Click()
    Call pS_InstallNew
End Sub

Private Sub btnRefreshList_Click()
    Call pS_PrepList
End Sub

Private Sub Form_Load()
    'toggle ribbon
    Call sP_ToggleRibbon
    'assure this is correct version
    vCtrl.CheckVersion
    'set form name
    Me.Caption = vCtrl.ProjVersion
    'refresh list
    Call pS_PrepList
End Sub
Private Sub pS_PrepList()
Dim DB As DAO.Database

    Set DB = CurrentDb
    'DoCmd.SetWarnings False
    With DB
        .Execute "DELETE * FROM " & ToolTbl & ";"
        .Execute "A_UniqueTools"
    End With
    Call pS_RefreshLbx
    Call SysCmd(acSysCmdSetStatus, "Tool last refeshed " & Format(Now, "dd mmm yyyy, hh:mm am/pm"))
End Sub
Private Sub pS_RefreshLbx()
    Me.lbxToolsSlctn.Requery
End Sub
Private Sub pS_InstallNew()
    'assure an item is selected
    Select Case lbxToolsSlctn.ItemsSelected.Count
        Case 0
            'inform user
            MsgBox "Pick a tool first.", vbExclamation + vbSystemModal, vCtrl.ProjVersion
        Case Is > 0
            'install tool
            vCtrl.InstallNew lbxToolsSlctn.Value
            'close tool installer
            Application.Quit
    End Select
    'install version.
End Sub
Private Sub sP_ToggleRibbon(Optional Togg As RibbTogg = RibbTogg.HideIt)
    If vCtrl.IsBeta Then Exit Sub
    Select Case Togg
        Case RibbTogg.HideIt
            DoCmd.ShowToolbar "Ribbon", acToolbarNo
        Case RibbTogg.ShowIt
            DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End Select
End Sub
