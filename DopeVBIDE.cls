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
Public Sub ExportGitFldr(Typ As vbext_ComponentType)
Dim Comp As VBIDE.VBComponent
Dim Path As Variant
Dim FlSys As New DopeFileSystem
Dim Exp As Variant

    'capture path
    Path = FlSys.GetPath(FldrPicker, False)
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
Private Sub Class_Initialize()
    'prepare global variables
    Set Proj = Application.VBE.ActiveVBProject
    Set Compnts = Proj.VBComponents
End Sub
