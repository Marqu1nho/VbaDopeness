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
Private AppObj As Object
Private Const DblNl As String = vbNewLine & vbNewLine
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
Private Enum CurrVrsn
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
Private Function F_ToolRepo(Optional WithProjName As Boolean = True)
Dim ErrMsg As String
Dim Arr As New DopeArray
    'add first set of items to array object
    Arr.AddNew Array(Environ("userprofile"), "Tesla", "Finops Billing - Tool Version Repository")
    'add project name if necessary
    If WithProjName Then Arr.AddNew Me.ProjName
    'return joined string
    F_ToolRepo = Arr.JoinedVals("\")
    
End Function
Private Sub Class_Initialize()
    Set AppObj = Application
End Sub
Public Property Get ProjName()
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
Dim Nm As String
    'capture full file name
    Nm = F_GetFileName(ProjNm)
    'remove extension
    Nm = Replace(Nm, F_GetFileExt(Nm), vbNullString)
    'return version name
    ProjVersion = Nm
End Property
Public Property Get ProjVersionNum(Optional Vtype As VNum = VNum.AsDouble)
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
Dim FlSys As New DopeFileSystem
    FlSys.NewFileWndw Me.ProjPaths
End Sub
Public Sub OpenRepoPath()
Dim FlSys As New DopeFileSystem
    FlSys.NewFileWndw F_ToolRepo
End Sub
Public Property Get ProjExtension()
Dim Nm As String
    'return extension piece of project name
    ProjExtension = F_GetFileExt(F_GetFileName(ProjNm))
End Property
Public Sub SetProjLocalFiles()
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
Private Sub S_AddShortcutToDesktop(ByVal Path As String, ByVal Icon As String)
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
    Prj = Me.ProjName
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
    Msg = "Version " & CurrVersion(NumOnly) & " of the " & Me.ProjName & _
          " tool is now in the ""Billing Ops Tools"" folder on your desktop."
    MsgBox Msg, vbInformation + vbSystemModal, Me.ProjName
    'open billing tools folder
    FlSys.NewFileWndw Dsktp
End Sub
Public Property Get ProjPaths(Optional Path As ProjDir = ProjDir.ProjPath)
    'make sure directories are created
    Me.SetProjLocalFiles
    'return path per path variable passed in
    ProjPaths = F_ProjPath(Path)
End Property
Public Sub ReinstallCurrent()
    
    'see what correct version is
    Call S_DownloadNewest
    'download it again
    
    'quit
End Sub
Public Property Get IsBeta()
    IsBeta = DLookup("State", "S_ThisTool", "[Version Num]=" & Me.ProjVersionNum) = "BETA"
End Property
Private Property Get CurrVersion(Optional Typ As CurrVrsn = CurrVrsn.NumOnly)
Dim Arr As New DopeArray
Dim VNum As Variant
Dim Ext As String
    VNum = Nz(DMax("[Version Num]", "S_ThisTool", "STATE='PROD'"), -1)
    Ext = Nz(DLookup("[File Extension]", "S_ThisTool", "[Version Num]=" & VNum))
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

Public Sub CheckVersion()
Dim IsBeta As Boolean
Dim Msg As Variant
Dim CurrV As Double
Dim MyV As Double
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
            'revert with positive message and exit
            MsgBox "You are using the current version of this tool - have fun!", vbInformation + vbSystemModal, Me.ProjName
    End Select
    
    
End Sub
Private Sub S_DownloadNewest()
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
Dim Proj As String
Dim Ico As String
    'set project path
    Proj = Environ("appdata") & "\BillingTools-" & Me.ProjName & "\"
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
Const VPattern As String = "* v#*.#*.*"
Dim ErrMsg As String
    If Not Nm Like VPattern Then
        ErrMsg = Nm & " file isn't formatted for version control." & DblNl & _
                "File name would need to end in "" v#.#.xlsm"" or "" v#.#.accdb/e"""
        Err.Raise vbObjectError, "DopeVersionCtrl.ProjName", ErrMsg
    End If
End Sub

Private Function F_CleanUpFileNm(ByVal Nm As String)
Const BetaPattern As String = "*BETA v#*.#*"
    'remove extension
    Nm = Replace(Nm, F_GetFileExt(Nm), vbNullString)
    'remove beta tag if present
    If Nm Like BetaPattern Then Nm = Replace(Nm, "BETA", vbNullString)
    'return cleaned up value
    F_CleanUpFileNm = Nm
End Function
Private Function F_GetFileExt(ByVal Nm As String)
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
Private Function F_GetFileName(NmType As Nm)
Dim Proj As Object
    'exit if the app object has yet to be set
    If AppObj Is Nothing Then
        F_GetFileName = Null
        Exit Function
    End If
    'capture the project type
    Select Case True
        Case AppObj.Name Like "*Access*"
            Set Proj = Access.Application.CurrentProject
        Case AppObj.Name Like "*Excel*"
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
