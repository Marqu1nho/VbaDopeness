VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DopeVBIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Private Proj As VBIDE.VBProject
Private Compnts As VBIDE.VBComponents
Private CdMd As VBIDE.CodeModule
Private Const pMeName As String = "DopeVBIDE"
Public Sub ExportGitFldr(Typ As vbext_ComponentType)
Dim Comp As VBIDE.VBComponent
Static Path As String
Dim FlSys As New DopeFileSystem
Dim Exp As Variant

    'capture path
    If Path = vbNullString Then Path = FlSys.GetPath(FldrPicker, False)
    For Each Comp In Compnts
        'determine whether this component is in scope to
        Exp = (Comp.Type = Typ)
        'export if in scope
        If Exp Then Comp.Export Path & "\" & Comp.Name & F_CompExt(Comp)
    Next Comp
End Sub
Private Function F_CompExt(Comp As VBIDE.VBComponent)
        'return extension based on component type
        Select Case Comp.Type
            Case vbext_ComponentType.vbext_ct_ClassModule, vbext_ComponentType.vbext_ct_Document
                F_CompExt = ".cls"
            Case vbext_ComponentType.vbext_ct_StdModule
                F_CompExt = ".bas"
            Case Else
                F_CompExt = ".txt"
        End Select
End Function
Public Sub ImportExistingClasses()
Dim FSO As New Scripting.FileSystemObject
Dim Fldr As Scripting.Folder
Dim Fl As Scripting.File
Dim Comp As VBIDE.VBComponent
Dim Nm As String
Dim Ext As String
Dim FlSys As New DopeFileSystem
Static Path As String
    'get path
    If Path = vbNullString Then Path = FlSys.GetPath(FldrPicker, False)
    'capture folder
    Set Fldr = FSO.GetFolder(Path)
    'loop through each folder
    For Each Fl In Fldr.Files
        'skip if not a class file
        If Not Fl.Name Like "*.cls" Then GoTo NextFl
        'skip the DopeVBIDE class
        If Fl.Name Like "DopeVBIDE*" Then GoTo NextFl
        'capture name w/ extension
        Nm = Fl.Name
        'loop through to see if we already have that class in
        For Each Comp In Proj.VBComponents
            If Nm = Comp.Name Then
                'remove component
                Call S_RemoveComponent(Comp)
                'import new component
                Proj.VBComponents.Import Fl.Path
                'exit loop as we've found a match
                Exit For
            End If
        Next Comp
        
NextFl:
    Next Fl
    

End Sub
Public Sub RemoveNonVBIDEClasses()
Dim Comp As VBIDE.VBComponent
Dim Typ As VBIDE.vbext_ComponentType
    For Each Comp In Compnts
        Select Case True
            Case Not Comp.Type = vbext_ct_ClassModule
                Debug.Print "NOT CLASS", Comp.Name
            Case Comp.Type = vbext_ct_ClassModule
                If Comp.Name = pMeName Then GoTo NextComp
                Debug.Print "IS CLASS", Comp.Name
                Call S_RemoveComponent(Comp)
        End Select
NextComp:
    Next Comp
End Sub
Private Sub S_RemoveComponent(Comp As VBComponent)
    Proj.VBComponents.Remove Comp
End Sub
Private Sub Class_Initialize()
    'prepare global variables
    Set Proj = Application.VBE.ActiveVBProject
    Set Compnts = Proj.VBComponents
End Sub
