VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DopeLogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private SavePath As String
Private ProjName As String
Private FSO As New Scripting.FileSystemObject
Private TxtStream As Scripting.TextStream
Private PrjoNm As String
Public Enum LogTyp
    L_Norm
    L_Warn
    L_Err
End Enum
Public Sub LogInfo(Typ As LogTyp, RtnNm As String, LogMsg As String)
Dim FullLine As String
Dim Arr As New DopeArray
Dim TimeStamp As String
Const ErrBeg As String = "Error------------------------------------------------"
Const ErrEnd As String = "ErrorEnd---------------------------------------------"
    'reformat input variables for readability
    LogMsg = Replace(LogMsg, vbNewLine, vbTab & vbNewLine)
    LogMsg = ">>" & LogMsg
    RtnNm = "Location:" & RtnNm
    'set timestamp
    TimeStamp = Format(Now, "hh:mm:ss am/pm")
    'concatenate line
    Arr.AddNew Array(F_LogTyp(Typ), TimeStamp, RtnNm, LogMsg)
    FullLine = Join(Array(F_LogTyp(Typ), TimeStamp, RtnNm, LogMsg), "|")
    'add error line if error
    If Typ = LogTyp.L_Err Then TxtStream.WriteLine ErrBeg
    'write line to file
    TxtStream.WriteLine FullLine
    'close error
    If Typ = LogTyp.L_Err Then TxtStream.WriteLine ErrEnd
End Sub

Private Sub Class_Initialize()
Dim TimeStamp As String
    'capture project name
    ProjName = CurrentProject.Name
    'set timestamp
    TimeStamp = Format(Now, "ddmmmyyyy")
    'set save path
    SavePath = Environ("appdata") & "\" & ProjName & "_" & TimeStamp & ".txt"
    'create file if it doesn't already exist
    If Not FSO.FileExists(SavePath) Then FSO.CreateTextFile SavePath
    'set text stream variable
    Set TxtStream = FSO.GetFile(SavePath).OpenAsTextStream(ForAppending)
    'add line for logging initiated
    Call Me.LogInfo(L_Norm, "DopeLogging.Class_Initialize", "Started Logging..")
End Sub

Private Function F_LogTyp(Typ As LogTyp)
'purpose: return the string log type
    Select Case Typ
        Case LogTyp.L_Norm
            F_LogTyp = "NORML"
        Case LogTyp.L_Warn
            F_LogTyp = "WARNG"
        Case LogTyp.L_Err
            F_LogTyp = "ERROR"
    End Select

End Function
