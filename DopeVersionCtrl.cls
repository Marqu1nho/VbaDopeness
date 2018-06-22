VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DopeVersionCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pAppObj As Object
Private pProjectName As String
Private Const DblNl As String = vbNewLine & vbNewLine
Private Const TblNm As String = "Tool Version Control"
Private Const QryNm As String = "S_ThisTool"
Private Const ToolPrefix As String = "\BillingTools-"
Private Const ToolInstQA As String = "https://teslamotorsinc.sharepoint.com/sites/FinopsBilling/SitePages/Tool-Installer.aspx"
Public Enum Nm
    ProjNm
    PathNm
End Enum
Private Enum App
    Exl
    Acs
    Otl
    Wrd
End Enum
Public Enum CurrVrsn
    NumOnly
    FullNm
    Icon
End Enum
Public Enum VNum
    AsString
    AsInteger
    AsDouble
End Enum
Public Enum ProjDir
    ProjPath
    IconPath
End Enum

Private Sub Class_Initialize()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: initiate the application object
'--------------------------------------------------------------------------------------------------------------------------
    Set pAppObj = Application
End Sub
Public Sub CheckVersion()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Checks whether the version being used is correct or not. If it is the correct it will clean up any previous
'versions in the app data folder. If it is not the correct version it will delete the old icon from the desktop, download
'grab the new version from the SP Tool repository, add the new icon to their desktop, and close its self.
'--------------------------------------------------------------------------------------------------------------------------
Dim IsBeta As Boolean
Dim Msg As Variant
Dim CurrV As Double
Dim MyV As Double
Dim FSO As New Scripting.FileSystemObject
Dim Fldr As Scripting.Folder
Dim Fl As Scripting.File
Dim Nm As String
Const NoProdVrs As Integer = -1
    'exit if in beta mode
    If Me.IsBeta Then
        Msg = Me.ProjVersion & " is set to BETA state in the version control list and is not being reviewed." & DblNl & _
                 "Only tools in PROD state get version controlled - contact Billing Ops management for assistance."
        MsgBox Msg, vbSystemModal + vbExclamation, Me.ProjName
        Exit Sub
    End If
    'capture my version
    MyV = Me.ProjVersionNum
    'capture currversion
    CurrV = CurrVersion
    'if there is not a valid prod version currversion returns a -1
    If CurrV = NoProdVrs Then
        Msg = "There is not a PROD version of this tool in the control list." & DblNl & "Please review with BillOps Management."
        MsgBox Msg, vbSystemModal + vbCritical, Me.ProjName
        Application.Quit
    End If
    'test to see if this version is the current version
    Select Case MyV = CurrV
        Case False
            'advise user newest version will be captured
            Msg = "You are not using the current version (" & CurrVersion(NumOnly) & "). " & DblNl & _
                  "Press OK to download newest or CANCEL to close."
            Msg = MsgBox(Msg, vbOKCancel + vbSystemModal + vbDefaultButton1 + vbExclamation, Me.ProjVersion)
            Select Case Msg
                Case vbOK
                    'grab newest version
                    Call S_DownloadNewest
                    Application.Quit
                Case vbCancel
                    Application.Quit
            End Select
        Case True
            'clean up previous versions of this project in the app data folder
            Set Fldr = FSO.GetFolder(Me.ProjPaths)
            Nm = Me.ProjName
            For Each Fl In Fldr.Files
                Debug.Print Fl.Name
                If Fl.Path = CurrentProject.FullName Or Fl.Name Like "*.laccdb" Then GoTo NextFl
                If Fl.Name Like Nm & " v#*.*" Then Fl.Delete
NextFl:
            Next Fl
            'revert with positive message and exit
            MsgBox "You are using the current version of this tool - have fun!", vbInformation + vbSystemModal, Me.ProjVersion
    End Select
End Sub
Public Sub InstallNew(ToolName As String)
Dim ToolVers As Variant
Dim ToolExt As String
Dim ToolPath As String
Dim IconPath As String
Dim FlSys As New DopeFileSystem
Dim ToolArr As New DopeArray
Dim IconArr As New DopeArray
Dim Fldr As Variant
    'get newest version uversion list
    ToolVers = DMax("[Version Num]", TblNm, "[Tool Name]='" & ToolName & "' and STATE='PROD'")
    'get extension for this version
    ToolExt = DLookup("[File Extension]", TblNm, "[Tool Name]='" & ToolName & "' and [Version Num]=" & ToolVers & " AND STATE='PROD'")
    'adjust for single integer versions
    If Len(ToolVers) = 1 Then ToolVers = ToolVers & ".0"
    'combine the full version
    ToolVers = ToolName & " v" & ToolVers & ToolExt
    'construct paths for repo directories
    ToolArr.AddNew Array(F_ToolRepo(WithProjName:=False), ToolName, ToolVers)
    ToolPath = ToolArr.JoinedVals("\")
    IconArr.AddNew Array(F_ToolRepo(WithProjName:=False), ToolName, ToolName & ".ico")
    IconPath = IconArr.JoinedVals("\")
    'reset array
    Set ToolArr = Nothing
    'add folders to test
    ToolArr.AddNew Array(Environ("appdata") & ToolPrefix & ToolName, Environ("appdata") & ToolPrefix & ToolName & "\ProjIcon")
    For Each Fldr In ToolArr.ArrayVals
        FlSys.NewFolder Fldr
    Next Fldr
    'copy file to new location
    FlSys.CopyFile ToolPath, ToolArr.ArrayVals(0), True
    'copy icon to new location
    FlSys.CopyFile IconPath, ToolArr.ArrayVals(1), True
    'add shortcut
    Call S_AddShortcutToDesktop(ToolArr.ArrayVals(0) & "\" & ToolVers, ToolArr.ArrayVals(1) & "\" & ToolName & ".ico", ToolName)
    'close
End Sub
Public Property Get ProjName()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: returns project name which excludes any beta tags, the v#.# format as well as the extension
'--------------------------------------------------------------------------------------------------------------------------
Dim ErrMsg As String
Dim Nm As String
    'capture name of project
    Nm = F_GetFileName(ProjNm)
    'test version control pattern name
    Call S_TestVersionPattern(Nm)
    'strip extenion and beta tag if present
    Nm = F_CleanUpFileNm(Nm)
    'strip version piece of project name
    Nm = Replace(Nm, " v" & Me.ProjVersionNum(AsString), vbNullString)
    'return project name
    ProjName = Nm
End Property
Public Property Get ProjVersion()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: returns the project version which is a combination of the project name as well as as the v#.# format
'--------------------------------------------------------------------------------------------------------------------------
Dim Nm As String
    'capture full file name
    Nm = F_GetFileName(ProjNm)
    'remove extension
    Nm = Replace(Nm, F_GetFileExt(Nm), vbNullString)
    'return version name
    ProjVersion = Nm
End Property
Public Property Get ProjVersionNum(Optional Vtype As VNum = VNum.AsDouble)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: returns the project version number. By default it'll return as double. But it can alternatively be returned as
'as string or integer.
'--------------------------------------------------------------------------------------------------------------------------
Dim CutOff As Integer
Dim Nm As Variant
    'capture file name
    Nm = F_GetFileName(ProjNm)
    'test version control pattern name
    Call S_TestVersionPattern(Nm)
    'remove extension and beta tag if it exists
    Nm = F_CleanUpFileNm(Nm)
    'reverse string
    Nm = StrReverse(Nm)
    'look for first v letter on reversed string
    'we can do this because this string passed s_testversionpattern sub
    CutOff = InStr(Nm, "v") - 1
    'capture version number piece and rever to normal
    Nm = StrReverse(Left(Nm, CutOff))
    'convert value per the vtype passed in
    Select Case Vtype
        Case VNum.AsDouble
            Nm = CDbl(Nm)
        Case VNum.AsInteger
            Nm = CInt(Nm)
        Case VNum.AsString
            Nm = CStr(Nm)
    End Select
    'return value
    ProjVersionNum = Nm
End Property
Public Sub OpenProjPath()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Opens to the project folder in the local appdata folder. It uses the DopeFileSystem Object.
'--------------------------------------------------------------------------------------------------------------------------
Dim FlSys As New DopeFileSystem
    FlSys.NewFileWndw Me.ProjPaths
End Sub
Public Sub OpenRepoPath()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Opens the file repository SharePoint library that the user should have synced on their machine.
'--------------------------------------------------------------------------------------------------------------------------
Dim FlSys As New DopeFileSystem
    FlSys.NewFileWndw F_ToolRepo
End Sub
Public Property Get ProjExtension()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Returns the file extension
'--------------------------------------------------------------------------------------------------------------------------
Dim Nm As String
    'return extension piece of project name
    ProjExtension = F_GetFileExt(F_GetFileName(ProjNm))
End Property
Public Sub SetProjLocalFiles()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Set up the project path in the appdata\roaming\ folder. It will create a project folder called
'BillingTools-ProjectName as well as a sub folder below that called ProjIcon where the icon is housed.
'--------------------------------------------------------------------------------------------------------------------------
Dim PathArr As Variant
Dim Itm As Variant
Dim FSO As New Scripting.FileSystemObject
    'set items in patharray
    PathArr = Array(F_ProjPath(ProjPath), F_ProjPath(IconPath))
    'test if each path exists, create if it doesn't
    For Each Itm In PathArr
        If Not FSO.FolderExists(Itm) Then FSO.CreateFolder Itm
    Next Itm
End Sub
Private Function F_ToolRepo(Optional WithProjName As Boolean = True)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Returns the path of the tool respository. Note that this assumes the user has synced the local one drive storage
'location. By default it'll return the the name of this project. If false is passed in, it'll only return the repo path.
'--------------------------------------------------------------------------------------------------------------------------
Dim ErrMsg As String
Dim Arr As New DopeArray
Dim Msg As String
Dim FSO As New Scripting.FileSystemObject
Const ChrmPath As String = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

    'add first set of items to array object
    Arr.AddNew Array(Environ("userprofile"), "Tesla", "Finops Billing - Tool Version Repository")
    'assure path exists
    If Not FSO.FolderExists(Arr.JoinedVals("\")) Then
        Msg = "Looks like you need to sync to the tool files first." & DblNl & _
              "Press OK for instructions to sync the FinOps OneDrive." & DblNl & _
              "This app will close once you press OK."
        MsgBox Msg, vbSystemModal + vbExclamation, Me.ProjName
        VBA.Shell ChrmPath & " -url " & ToolInstQA
        Application.Quit
    End If
    'add project name if necessary
    If WithProjName Then Arr.AddNew Me.ProjName
    'return joined string
    F_ToolRepo = Arr.JoinedVals("\")
    
End Function
Public Property Get ProjPaths(Optional Path As ProjDir = ProjDir.ProjPath)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: set the local project and icon sub folder paths. Then based on the path passed in it will return the project
'path or the icon sub path.
'--------------------------------------------------------------------------------------------------------------------------
    'make sure directories are created
    Me.SetProjLocalFiles
    'return path per path variable passed in
    ProjPaths = F_ProjPath(Path)
End Property
Public Property Get IsBeta()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Retruns TRUE/FALSE if the state for the version being used is set to BETA or PROD in the version control list.
'--------------------------------------------------------------------------------------------------------------------------
    IsBeta = DLookup("State", QryNm, "[Version Num]=" & Me.ProjVersionNum & " AND [File Extension]='" & Me.ProjExtension & "'") = "BETA"
End Property
Public Property Get CurrVersion(Optional Typ As CurrVrsn = CurrVrsn.NumOnly)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Grabs the current PROD version from the vesion control list. By default it will only grab the value from the
'version num field. But alternatively the full path of the the current version or icon on the SP version repository
'can be returned.
'--------------------------------------------------------------------------------------------------------------------------
Dim Arr As New DopeArray
Dim VNum As Variant
Dim Ext As String
    VNum = Nz(DMax("[Version Num]", QryNm, "STATE='PROD'"), -1)
    Ext = Nz(DLookup("[File Extension]", QryNm, "[Version Num]=" & VNum))
    Select Case Typ
        Case CurrVrsn.NumOnly
            'nothing - value is just returned at end of property
        Case CurrVrsn.FullNm
            'adjust for full intergers so the string will match
            If Len(VNum) = 1 Then VNum = VNum & ".0"
            Arr.AddNew Array(F_ToolRepo(True), Me.ProjName & " v" & VNum & Ext)
            VNum = Arr.JoinedVals("\")
        Case CurrVrsn.Icon
            Arr.AddNew Array(F_ToolRepo(True), Me.ProjName & ".ico")
            VNum = Arr.JoinedVals("\")
    End Select
    'return value
    CurrVersion = VNum
End Property
Private Function F_GetFileName(NmType As Nm)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: returns the name of the project based on if access or excel is being used. note this is not the full formed path.
'--------------------------------------------------------------------------------------------------------------------------
Dim Proj As Object
    'exit if the app object has yet to be set
    If pAppObj Is Nothing Then
        F_GetFileName = Null
        Exit Function
    End If
    'capture the project type
    Select Case True
        Case pAppObj.Name Like "*Access*"
            Set Proj = Access.Application.CurrentProject
        Case pAppObj.Name Like "*Excel*"
            Set Proj = Excel.Application.ThisWorkbook
    End Select
    'return the name per the name type passed in
    Select Case NmType
        Case Nm.ProjNm
            F_GetFileName = Proj.Name
        Case Nm.PathNm
            F_GetFileName = Proj.FullName
    End Select
End Function
Private Sub S_AddShortcutToDesktop(ByVal Path As String, ByVal Icon As String, Optional InstallNm As Variant = Null)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: Called from S_DownloadNewest. Grabs the icon from the repo folder and replaces the item in the app data icon
'folder. it will replace it every time there is a new version which will allow us to be able to change the icon and have
'it update to all the users icons.
'--------------------------------------------------------------------------------------------------------------------------
Dim FlSys As New DopeFileSystem
Dim Dsktp As String
Dim Flnm As String
Dim ShrCt As IWshRuntimeLibrary.WshShortcut
Dim Shll As New IWshRuntimeLibrary.WshShell
Dim FSO As New Scripting.FileSystemObject
Dim Fldr As Scripting.Folder
Dim Fl As Scripting.File
Dim Msg As String
Dim Prj As String
Const IcnFldr As String = "\Desktop\Billing Ops Tools\"
    'capture desktop path
    Dsktp = Environ("userprofile") & IcnFldr
    'create it if it doesn't already exists
    FlSys.NewFolder Dsktp
    'capture desktop path and project name
    Select Case IsNull(InstallNm)
        Case True
            Prj = Me.ProjName
        Case False
            Prj = InstallNm
    End Select
    Set Fldr = FSO.GetFolder(Dsktp)
    'loop through shortcuts on billing ops tools folders and get rid of any for this project
    For Each Fl In Fldr.Files
        If Fl.Name Like Prj & "*.lnk" Then Kill Fl
    Next Fl
    'capture path file name
    Flnm = FlSys.GetFileName(Path)
    'replace extension with .lnk extension
    Flnm = Replace(Flnm, FlSys.GetFileExt(Flnm), ".lnk")
    Flnm = Dsktp & Flnm
    Set ShrCt = Shll.CreateShortcut(Flnm)
    'create shortcut
    With ShrCt
        .TargetPath = Path
        .WindowStyle = vbNormalFocus
        .Description = "Shortcut to " & Path
        .WorkingDirectory = Environ("userprofile")
        .IconLocation = Icon
        .Save
    End With
    'notify user
    Select Case IsNull(InstallNm)
        Case True
            Msg = "The latest version of the " & Me.ProjName & " tool is now on your " & _
                  """Billing Ops Tools"" folder on your desktop with below name:" & DblNl & _
                  F_StripFileName(Path)
        Case False
            Msg = "The latest version of the " & InstallNm & " tool is now on your " & _
                  """Billing Ops Tools"" folder on your desktop with below name:" & DblNl & _
                  F_StripFileName(Path)
    End Select
    MsgBox Msg, vbInformation + vbSystemModal, Me.ProjName
    'open billing tools folder
    FlSys.NewFileWndw Dsktp
End Sub
Private Sub S_DownloadNewest()
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: downloads the newest version from the version repositiory. this is driven by the version control list.
'--------------------------------------------------------------------------------------------------------------------------
Dim RepoFile As Variant
Dim FlSys As New DopeFileSystem
Dim FromPath As String
Dim ToPath As String
Dim IcoFile As String
    'capture paths
    FromPath = CurrVersion(FullNm)
    ToPath = Me.ProjPaths(ProjPath) & FlSys.GetFileName(FromPath)
    IcoFile = Me.ProjPaths(IconPath) & Me.ProjName & ".ico"
    'copy file from repo to app data folder
    FlSys.CopyFile CurrVersion(FullNm), Me.ProjPaths(ProjPath), True
    'copy icon file from repo to app data icon folder
    FlSys.CopyFile CurrVersion(Icon), Me.ProjPaths(IconPath), True
    'add shortcut to desktop
    Call S_AddShortcutToDesktop(ToPath, IcoFile)
    
    

    
End Sub
Private Function F_ProjPath(Typ As ProjDir)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: returns either the project path or the icon path in the icon local app data folder.
'--------------------------------------------------------------------------------------------------------------------------
Dim Proj As String
Dim Ico As String
    'set project path
    Proj = Environ("appdata") & ToolPrefix & Me.ProjName & "\"
    'set icon path
    Ico = Proj & "ProjIcon\"
    Select Case Typ
        Case ProjDir.ProjPath
            F_ProjPath = Proj
        Case ProjDir.IconPath
            F_ProjPath = Ico
    End Select
End Function
Private Sub S_TestVersionPattern(ByVal Nm As String)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: check to assure the project is in correct pattern which consists of "some project name" [space] "v#.#.extension"
'if it is not in the correct version it will raise an error.
'--------------------------------------------------------------------------------------------------------------------------
Const VPattern As String = "* v#*.#*.*"
Dim ErrMsg As String
    If Not Nm Like VPattern Then
        ErrMsg = Nm & " file isn't formatted for version control." & DblNl & _
                "File name would need to end in "" v#.#.xlsm"" or "" v#.#.accdb/e"""
        Err.Raise vbObjectError, "DopeVersionCtrl.ProjName", ErrMsg
    End If
End Sub
Private Function F_CleanUpFileNm(ByVal Nm As String)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: retuns the project name without extension and beta tag.
'--------------------------------------------------------------------------------------------------------------------------
Const BetaPattern As String = "*BETA v#*.#*"
Const DupPattern As String = "* (#)"
    'remove extension
    Nm = Replace(Nm, F_GetFileExt(Nm), vbNullString)
    'remove beta tag if present
    If Nm Like BetaPattern Then Nm = Replace(Nm, "BETA v", " v")
    'remove excess parantheses in case multiple versions are opened
    If Nm Like DupPattern Then Nm = Left(Nm, Len(Nm) - 4)
    'return cleaned up value
    F_CleanUpFileNm = Nm
End Function
Private Function F_GetFileExt(ByVal Nm As String)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: retuns file extesnion by looking at the string after the last period. if a period is not found then null
'is returned.
'--------------------------------------------------------------------------------------------------------------------------
Dim CutOff As Integer
    'reverse name
    Nm = StrReverse(Nm)
    'find cutoff of first period
    CutOff = InStr(Nm, ".")
    'if period is not found then cutoff will be so we'll raise an error
    Select Case CutOff
        Case 0
            F_GetFileExt = Null
        Case Is > 0
            F_GetFileExt = StrReverse(Left(Nm, CutOff))
    End Select
End Function
Private Function F_StripFileName(ByVal Nm As String)
'--------------------------------------------------------------------------------------------------------------------------
'PURPOSE: retruns file name of a full formed path by looking for the last \ character. If a \ character is not found it is
'it returns the same string passed in.
'--------------------------------------------------------------------------------------------------------------------------
Dim CutOff As Integer
    'reverse string
    Nm = StrReverse(Nm)
    'look for first "\"
    CutOff = InStr(Nm, "\")
    'test to see if the back slash was found
    Select Case CutOff
        Case 0
            'return the original name as it was passed in
            F_StripFileName = StrReverse(Nm)
        Case Is > 0
            'return just the file name
            F_StripFileName = StrReverse(Left(Nm, CutOff - 1))
    End Select
End Function
