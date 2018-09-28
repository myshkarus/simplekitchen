Attribute VB_Name = "VBScript"
Option Explicit
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Private tempFolder As String

Sub test()
  Call CopyFile("D:\1\2.xlsx", "D:\2")
  Call CopyFile("D:\1\1.xlsx", "D:\2")
  Call MoveFile("D:\1\3.xlsx", "D:\2")
End Sub

Public Sub MoveFile(ByVal source As String, ByVal destination As String)
  Dim file As String
  Dim name As String
  Dim script As String
  Dim wshShell As Object

  script = MoveToScript: name = "vsmove"
  DoEvents
  Set wshShell = MyShell()
  If Not wshShell Is Nothing Then
    DoEvents
    file = FilePath(name, script)
    wshShell.Run Chr(34) & file & Chr(34) & " " & source & " " & destination
  End If
  'это в какой-нибудь другой процедуре, которая исполняется позднее:
  Sleep 200
  Kill file
  Set wshShell = Nothing
End Sub

Public Sub CopyFile(ByVal source As String, ByVal destination As String)
  Dim file As String
  Dim name As String
  Dim script As String
  Dim wshShell As Object

  script = CopyScript: name = "vscopy"
  DoEvents
  Set wshShell = MyShell()
  If Not wshShell Is Nothing Then
    DoEvents
    file = FilePath(name, script)
    wshShell.Run Chr(34) & file & Chr(34) & " " & source & " " & destination
  End If
  'это в какой-нибудь другой процедуре, которая исполняется позднее:
  Sleep 200
  Kill file
  Set wshShell = Nothing
End Sub

Private Function MyShell() As Object
  Dim wsh As Object
  Set wsh = CreateObject("Wscript.Shell")
  tempFolder = wsh.SpecialFolders("Templates")
  'tempFolder = ThisWorkbook.Path & "\1\"
  If Not wsh Is Nothing Then Set MyShell = wsh
  Set wsh = Nothing
End Function

Private Function FilePath(ByVal name As String, ByVal myScript As String) As String
  Dim intFileNum As Integer
  Dim sFileName As String
  If Len(name) = 0 Then Exit Function
  sFileName = tempFolder & "\" & name & ".vbs"
  intFileNum = FreeFile
  Open sFileName For Output As intFileNum
  Print #intFileNum, myScript
  Close intFileNum
  FilePath = sFileName
End Function

Private Function CopyScript() As String
  Dim s As String
  s = s & "Option Explicit" & vbCrLf
  s = s & "Call CopyFile(WScript.Arguments(0),WScript.Arguments(1))" & vbCrLf
  s = s & "Sub CopyFile(Source, Destination)" & vbCrLf
  s = s & "Dim wasReadOnly" & vbCrLf
  s = s & "Dim fso" & vbCrLf
  s = s & "Dim fileName" & vbCrLf
  s = s & "On Error Resume Next" & vbCrLf
  s = s & "If Right(Destination, 1) <> " & Chr(34) & "\" & Chr(34) & " Then Destination = Destination & " & Chr(34) & "\" & Chr(34) & vbCrLf
  s = s & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
  s = s & "wasReadOnly = False" & vbCrLf
  s = s & "If fso.FileExists(Source) = False Then Exit Sub" & vbCrLf
  s = s & "fileName = Right(Source, InStrRev(Source, " & Chr(34) & "\" & Chr(34) & ") + 1)" & vbCrLf
  s = s & "fileName = Destination & fileName" & vbCrLf
  's = s & "WScript.Echo fileName" & vbCrLf
  s = s & "If fso.FileExists(fileName) Then" & vbCrLf
  s = s & "  If fso.GetFile(fileName).Attributes And 1 Then" & vbCrLf
  s = s & "     fso.GetFile(fileName).Attributes = fso.GetFile(fileName).Attributes - 1" & vbCrLf
  s = s & "     wasReadOnly = True" & vbCrLf
  s = s & "  End If" & vbCrLf
  s = s & "  fso.DeleteFile fileName, True" & vbCrLf
  s = s & "End if" & vbCrLf
  s = s & "fso.CopyFile Source, Destination, True" & vbCrLf
  s = s & "If wasReadOnly Then" & vbCrLf
  s = s & "  fso.GetFile(fileName).Attributes = fso.GetFile(fileName).Attributes + 1" & vbCrLf
  s = s & "End if" & vbCrLf
  s = s & "Set fso = Nothing" & vbCrLf
  s = s & "End Sub" & vbCrLf
  CopyScript = s
End Function

Private Function MoveToScript() As String
  Dim s As String
  s = s & "Option Explicit" & vbCrLf
  s = s & "Call MoveFile(WScript.Arguments(0),WScript.Arguments(1))" & vbCrLf
  s = s & "Sub MoveFile(fromSource, toDestination)" & vbCrLf
  s = s & "Dim fso" & vbCrLf
  s = s & "Dim fileName" & vbCrLf
  s = s & "On Error Resume Next" & vbCrLf
  s = s & "If Right(toDestination, 1) <> " & Chr(34) & "\" & Chr(34) & " Then toDestination = toDestination & " & Chr(34) & "\" & Chr(34) & vbCrLf
  s = s & "fileName = Right(fromSource, InStrRev(fromSource, " & Chr(34) & "\" & Chr(34) & ") + 1)" & vbCrLf
  s = s & "fileName = toDestination & fileName" & vbCrLf
  's = s & "WScript.Echo fileName" & vbCrLf
  s = s & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
  s = s & "If fso.FileExists(fromSource) = False Then Exit Sub" & vbCrLf
  s = s & "If fso.FileExists(fileName) Then" & vbCrLf
  s = s & "  fso.DeleteFile(fileName), True" & vbCrLf
  s = s & "End if" & vbCrLf
  s = s & "fso.MoveFile fromSource, toDestination" & vbCrLf
  s = s & "Set fso = Nothing" & vbCrLf
  s = s & "End Sub" & vbCrLf
  MoveToScript = s
End Function



