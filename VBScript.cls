VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VBS_FILE As String = "skvbscript"

Private Type scriptFile
   fullPath As String
   Folder As String
   FileName As String
End Type

#If Development Then
   'for early binding
   Private fso As Scripting.FileSystemObject
   Private shell As wshShell
#Else
   'for late binding
   Private fso As Object
   Private shell As Object
#End If

Private this As scriptFile

Public Property Get Folder() As String
   Folder = this.Folder
End Property

Public Property Get FileName() As String
   FileName = this.FileName
End Property

Public Property Get fullPath() As String
   fullPath = this.fullPath
End Property

Private Sub Class_Initialize()
   Dim fileNum As Long
   On Error GoTo FailExit
   #If Development Then
      'for early binding
      Set fso = New Scripting.FileSystemObject
      Set shell = New wshShell
   #Else
      'for late binding
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set shell = CreateObject("Wscript.Shell")
   #End If
   
   this.fullPath = Environ("temp") & "\" & VBS_FILE & ".vbs"
   On Error Resume Next
   If Dir(this.fullPath) <> "" Then Kill this.fullPath
   On Error GoTo 0
   
   On Error GoTo FailExit
   If Dir(this.fullPath) = "" Then
      fileNum = FreeFile
      Open this.fullPath For Output As #fileNum
      Close #fileNum
   End If
   
   With fso
      If .FileExists(this.fullPath) Then
         this.FileName = .GetFileName(this.fullPath)
         this.Folder = .GetParentFolderName(this.fullPath)
      Else: GoTo FailExit
      End If
   End With
   Exit Sub
FailExit:
   Debug.Print Err.Number, Err.description, Err.source
   End
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   fso.DeleteFile this.fullPath
   On Error GoTo 0
   Set fso = Nothing
   Set shell = Nothing
End Sub

Public Sub ReopenFile(ByVal path As String)
   Dim script As String
   script = ReopenScript
   Call WriteScript(script)
   DoEvents
   If Not shell Is Nothing Then
      If path <> vbNullString Then
        shell.Run Chr(34) & this.fullPath & Chr(34) & " " & path
      End If
   End If
End Sub

Private Sub WriteScript(ByVal script As String)
   Dim output() As String
   Dim i As Long
   Dim fileNum As Long
   On Error GoTo FailExit
   If script = vbNullString Then Exit Sub
   fileNum = FreeFile
   output = Split(script, "|")
   Open this.fullPath For Output As #fileNum
   For i = LBound(output) To UBound(output)
      Print #fileNum, output(i)
   Next i
   Close #fileNum
   Exit Sub
FailExit:
   Debug.Print Err.Number, Err.source, Err.description
End Sub

'TODO: testing
Public Sub MoveFile(ByVal source As String, ByVal destination As String)
   Dim file As String
   Dim name As String
   Dim script As String
   Dim shell As Object
   name = VS_MOVE_FILE
   script = MoveToScript
   file = ScriptPath(name, script)
   Set shell = CreateObject("Wscript.Shell")
   DoEvents
   shell.Run Chr(34) & file & Chr(34) & " " & source & " " & destination
   Set shell = Nothing
End Sub

'TODO: testing
Public Sub CopyFile(ByVal source As String, ByVal destination As String)
   Dim file As String
   Dim name As String
   Dim script As String
   Dim shell As Object
   name = VS_COPY_FILE
   script = CopyScript
   file = ScriptPath(name, script)
   Set shell = CreateObject("Wscript.Shell")
   DoEvents
   shell.Run Chr(34) & file & Chr(34) & " " & source & " " & destination
   Set shell = Nothing
End Sub

Private Function ReopenScript() As String
   Dim vs As String
   vs = "Option Explicit" & _
             "|Call ReopenFile(WScript.Arguments(0))"
   vs = vs & "|Sub ReopenFile(File)" & _
             "|  Dim fso, xlApp, objStream" & _
             "|  Dim wait" & _
             "|  On Error Resume Next"
   vs = vs & "|  Set fso = CreateObject(""Scripting.FileSystemObject"")"
   vs = vs & "|  If Not fso.FileExists(File) Then Exit Sub" & _
             "|  wait = 300" & _
             "|  Do" & _
             "|    Set objStream = fso.OpenTextFile(File, 8, False)" & _
             "|    If Err.Number = 0 Then" & _
             "|      objStream.Close"
   vs = vs & "|      Set xlApp = GetObject(, ""Excel.Application"")"
   vs = vs & "|      If Err.Number <> 0 Then" & _
             "|      WScript.Echo ""err <> 0"""
   vs = vs & "|        Set xlApp = CreateObject(""Excel.Application"")"
   vs = vs & "|      End If" & _
             "|      'xlApp.Application.DisplayAlerts = False" & _
             "|      xlApp.Workbooks.Open File" & _
             "|      'xlApp.Application.Visible = True" & _
             "|      Set objStream = Nothing" & _
             "|      Set xlApp = Nothing" & _
             "|      Set fso = Nothing" & _
             "|      Exit Sub" & _
             "|    Else" & _
             "|      WScript.Sleep wait" & _
             "|      Err.Clear" & _
             "|    End If" & _
             "|  Loop" & _
             "|  Set fso = Nothing" & _
             "|  Set xlApp = Nothing" & _
             "|End Sub"
   ReopenScript = vs
End Function

Private Function CopyScript() As String
   Dim vs As String
   vs = "Option Explicit" & _
        "|Call CopyFile(WScript.Arguments(0), WScript.Arguments(1))"
   vs = vs & "|Sub CopyFile(Source, Destination)" & _
        "|  Dim wasReadOnly" & _
        "|  Dim fso" & _
        "|  Dim fileName" & _
        "|  On Error Resume Next" & _
        "|  If Right(Destination, 1) <> " & Chr(34) & "\" & Chr(34) & " Then Destination = Destination & " & Chr(34) & "\" & Chr(34)
   vs = vs & "|  Set fso = CreateObject(""Scripting.FileSystemObject"")"
   vs = vs & "|  wasReadOnly = False" & _
        "|  If fso.FileExists(Source) = False Then Exit Sub" & _
        "|  fileName = Right(Source, InStrRev(Source, " & Chr(34) & "\" & Chr(34) & ") + 1)" & _
        "|  fileName = Destination & fileName" & _
        "|  If fso.FileExists(fileName) Then" & _
        "|    If fso.GetFile(fileName).Attributes And 1 Then" & _
        "|       fso.GetFile(fileName).Attributes = fso.GetFile(fileName).Attributes - 1" & _
        "|       wasReadOnly = True" & _
        "|    End If" & _
        "|    fso.DeleteFile fileName, True" & _
        "|  End If" & _
        "|  fso.CopyFile Source, Destination, True" & _
        "|  If wasReadOnly Then" & _
        "|    fso.GetFile(fileName).Attributes = fso.GetFile(fileName).Attributes + 1" & _
        "|  End If" & _
        "|  Set fso = Nothing" & _
        "|End Sub"
   CopyScript = vs
End Function

Private Function MoveToScript() As String
   Dim vs As String
   vs = "Option Explicit" & _
        "|Call MoveFile(WScript.Arguments(0), WScript.Arguments(1))"
   vs = vs & "|Sub MoveFile(fromSource, toDestination)" & _
        "|  Dim fso" & _
        "|  Dim fileName" & _
        "|  On Error Resume Next" & _
        "|  If Right(toDestination, 1) <> " & Chr(34) & "\" & Chr(34) & " Then toDestination = toDestination & " & Chr(34) & "\" & Chr(34) & _
        "|  fileName = Right(fromSource, InStrRev(fromSource, " & Chr(34) & "\" & Chr(34) & ") + 1)" & _
        "|  fileName = toDestination & fileName"
   vs = vs & "|  Set fso = CreateObject(""Scripting.FileSystemObject"")"
   vs = vs & "|  If fso.FileExists(fromSource) = False Then Exit Sub" & _
        "|  If fso.FileExists(fileName) Then" & _
        "|    fso.DeleteFile(fileName), True" & _
        "|  End If" & _
        "|  fso.MoveFile fromSource, toDestination" & _
        "|  Set fso = Nothing" & _
        "|End Sub"
   MoveToScript = vs
End Function


