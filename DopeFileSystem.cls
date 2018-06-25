VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DopeFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Enum FlDialog
    FldrPicker = msoFileDialogFolderPicker
    FilesPicker = msoFileDialogFilePicker
End Enum
Public Function NewFolder(Path As Variant, Optional OpenFldr As Boolean = False)
Dim FSO As New Scripting.FileSystemObject
Dim ParentFld As String
Dim ErrMsg As String
Dim CutOff As String
    'remove last backslash if necessary
    If Path Like "*\" Then Path = Left(Path, Len(Path) - 1)
    'test if folder exists
    If Not FSO.FolderExists(Path) Then
        'extract parent folder
        ParentFld = StrReverse(Path)
        CutOff = InStr(ParentFld, "\")
        ParentFld = Left(StrReverse(ParentFld), Len(ParentFld) - CutOff)
        'only create final folder if parent folder exists
        Select Case FSO.FolderExists(ParentFld)
            Case True
                FSO.CreateFolder Path
            Case False
                ErrMsg = "Not all folders in your path exist." & vbNewLine & "Make sure all folders exists before creating last folder in path."
                Err.Raise vbObjectError, "DopeFileSystem.NewFolder", ErrMsg
        End Select
    End If
    If OpenFldr Then Me.NewFileWndw Path
End Function
Public Sub NewFileWndw(Path As Variant)
Dim FSO As New Scripting.FileSystemObject
    If Not FSO.FolderExists(Path) Then
        Err.Raise vbObjectError, "DopeFileSystem.NewFileWndw", "Bad path."
        Exit Sub
    End If
    Shell "explorer.exe " & Path, vbNormalFocus
End Sub
Public Function GetFileName(ByVal Path As String)
Dim CutOff As Integer
    'in case a folder name is passed, get rid of final "\"
    If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
    'reverse name
    Path = StrReverse(Path)
    'find cutoff of first period
    CutOff = InStr(Path, "\") - 1
    'if period is not found then cutoff will be so we'll raise an error
    Select Case CutOff
        Case -1
            GetFileName = StrReverse(Path)
        Case Is > -1
            GetFileName = StrReverse(Left(Path, CutOff))
    End Select
End Function
Public Sub CopyFile(FromPath As String, ToPath As String, Optional ReplaceTo As Boolean = False)
Dim FSO As New Scripting.FileSystemObject
Dim ErrMsg As String
    'assure that ToPath exists
    If Not FSO.FileExists(FromPath) Then
        ErrMsg = "FromPath file doesn't exist."
        Err.Raise vbObjectError, "DopeFileSystem.CopyFile", ErrMsg
    End If
    'fix to path
    If Not ToPath Like "*\" Then ToPath = ToPath & "\"
    'move file
    FSO.CopyFile FromPath, ToPath, overwritefiles:=ReplaceTo
End Sub
Public Function GetFileExt(ByVal Path As String)
Dim CutOff As Integer
    'reverse name
    Path = StrReverse(Path)
    'find cutoff of first period
    CutOff = InStr(Path, ".")
    'if period is not found then cutoff will be so we'll raise an error
    Select Case CutOff
        Case 0
            GetFileExt = Null
        Case Is > 0
            GetFileExt = StrReverse(Left(Path, CutOff))
    End Select
End Function
Public Function GetPath(GetType As FlDialog, Optional MultFiles As Boolean = False, Optional FilterExt As String = "*.*", _
                        Optional filterNm As String = "All Files", Optional InitialDir As String = vbNullString, _
                        Optional BtnNm As String = vbNullString)
Dim FSO As New Scripting.FileSystemObject
Dim FilePicker As FileDialog
Dim Slctn As Variant
Dim AllSlctn As Variant
Dim Arr As New DopeArray
    'set button type if not passed in
    If BtnNm = vbNullString Then
        Select Case GetType
            Case FlDialog.FldrPicker
                BtnNm = "Select Folder(s)"
            Case FlDialog.FilesPicker
                BtnNm = "Select File(s)"
        End Select
    End If
    'test that initial dir passed in is valid
    If Not InitialDir = vbNullString Then
        'if not valid path then set to blank
        Debug.Print "Setting InitialDir to vbnullstring as not valid directory.."
        If Not FSO.FolderExists(InitialDir) Then InitialDir = vbNullString
    End If
    'set file picker object based on get type passed in
    Set FilePicker = FileDialog(GetType)
    'set configuration
    With FilePicker
        .AllowMultiSelect = MultFiles
        .Title = BtnNm
        .ButtonName = BtnNm
        .InitialFileName = InitialDir
        '.Filters.Add filterNm, FilterExt
        .Show
    End With
    'capture each file selected
    For Each Slctn In FilePicker.SelectedItems
        Arr.AddNew Slctn
    Next
    'return value based on how many files were selected
    Select Case Arr.Count
        Case 1
            GetPath = Arr.JoinedVals
        Case Is > 1
            GetPath = Arr.ArrayVals
        Case Else
            GetPath = Null
    End Select
End Function
