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
Public Enum Vnum
    AsString
    AsInteger
    AsDouble
End Enum
Public Enum ProjDir
    ProjPath
    IconPath
End Enum

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
    ProjVersion = F_GetFileName(ProjNm)
End Property
Public Property Get ProjVersionNum(Optional Vtype As Vnum = Vnum.AsString)
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
        Case Vnum.AsDouble
            Nm = CDbl(Nm)
        Case Vnum.AsInteger
            Nm = CInt(Nm)
        Case Vnum.AsString
            Nm = CStr(Nm)
    End Select
    'return value
    ProjVersionNum = Nm
End Property
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
Public Property Get ProjPaths(Optional Path As ProjDir = ProjDir.ProjPath)
    'make sure directories are created
    Me.SetProjLocalFiles
    'return path per path variable passed in
    ProjPaths = F_ProjPath(Path)
End Property
Public Sub ReinstallCurrent()
    Debug.Print "Coming soon - need to figure out where we'll keep the version files."
    'see what correct version is
    
    'download it again
    
    'quit
End Sub
Public Sub CheckVersion()
    Debug.Print "Coming soon - need to figure out where we'll keep the version list."
    's_downloadnewest get called from here
End Sub
Private Sub S_DownloadNewest()
    Debug.Print "Coming soon - need to figure out where we'll keep the version files."
End Sub
Private Function F_ProjPath(Typ As ProjDir)
Dim Proj As String
Dim Ico As String
    'set project path
    Proj = Environ("appdata") & "\App-" & Me.ProjName & "\"
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
