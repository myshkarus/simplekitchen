VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------
'@hqinside.com
'================================================================================================
'SysInfo [Static] Class, ver.1.2 24/03/2018: Attribute VB_PredeclaredId = True
'Возврат системных параметров
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'ТЕКУЩИЙ СТАТУС: тестирование
'================================================================================================

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

#If Win64 Then
  #If VBA7 Then
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As LongPtr
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As LongPtr
  #Else
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As LongPtr) As LongPtr
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As LongPtr
  #End If
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
  #Else
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
  #End If
#End If

Public Enum TargetProvider
  AceOLEDB
  JetOLEDB                                       'Microsoft Jet 4.0 OLE DB Provider
  SQLOLEDB                                       'Microsoft OLE DB Provider for SQL Server
  SQLLiteOLEDB                                   'SQLLite OLE DB 4.0 Provider
End Enum


'Sub DisplayMonitorInfo()
'Dim w As Long, h As Long
'    w = GetSystemMetrics32(0) ' width in points
'    h = GetSystemMetrics32(1) ' height in points
'    MsgBox Format(w, "#,##0") & " x " & Format(h, "#,##0"), _
'    vbInformation, "Monitor Size (width x height)"
'End Sub

#If Win64 Then
  Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

  Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
  Dim TimerId As LongPtr
  Dim hdc As LongPtr
  Dim mhWndForm As LongPtr
#Else
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
  Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim TimerId As Long
  Dim hdc As Long
  Dim mhWndForm As Long
#End If
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90


Dim TimerSeconds As Single, tim As Boolean, Counter As Long
Private Const SM_CXMAXIMIZED = 61
Private Const SM_CYMAXIMIZED = 62
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private DefaultHighMaximizedWindow As Long
Private g_screenWidth As Long

Public Function PointPerPixelX() As Double
  hdc = GetDC(0)
  PointPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
  ReleaseDC 0, hdc
End Function

Public Function PointPerPixelY() As Double
  hdc = GetDC(0)
  PointPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
  ReleaseDC 0, hdc
End Function

Public Sub SetScreenDimension()
  DefaultHighMaximizedWindow = GetSystemMetrics32(SM_CYMAXIMIZED)
  DefaultHighMaximizedWindow = DefaultHighMaximizedWindow * PointPerPixelY
  g_screenWidth = GetSystemMetrics32(SM_CXMAXIMIZED)
  g_screenWidth = g_screenWidth * PointPerPixelX
  DefaultHighMaximizedWindow = DefaultHighMaximizedWindow - 20
  g_screenWidth = g_screenWidth - 20
  If DefaultHighMaximizedWindow < 300 Then DefaultHighMaximizedWindow = 300
  If g_screenWidth < 300 Then g_screenWidth = 300
End Sub

Public Property Get ScreenHeight() As Long
ScreenHeight = CLng(GetSystemMetrics(SM_CYSCREEN))
End Property

Public Property Get ScreenWidth() As Long
ScreenWidth = CLng(GetSystemMetrics(SM_CXSCREEN))
End Property

Public Function provider(Optional useProvider As TargetProvider = AceOLEDB) As String
  Select Case useProvider
  Case TargetProvider.SQLOLEDB: provider = "SQLOLEDB.1"
  Case TargetProvider.SQLLiteOLEDB:              'сделать позднее
  Case Else:
    If val(ExcelVersion) < 12 Then
      provider = "Microsoft.Jet.OLEDB.4.0"       '.mdb Access 97 - Access 2003
    Else
      provider = "Microsoft.ACE.OLEDB.12.0"      '.mdb, .accdb Access 2007+
    End If
  End Select
End Function

Public Function TimeToMillisecond() As String
  Dim tSystem As SYSTEMTIME
  Dim sRet As String * 12
  On Error Resume Next
  GetSystemTime tSystem
  sRet = Strings.Format(Hour(Now), "00") & ":" & Strings.Format(Minute(Now), "00") & ":" & Strings.Format(Second(Now), "00") & ":" & Strings.Format(tSystem.wMilliseconds, "000")
  TimeToMillisecond = sRet
End Function

Public Function AddInName() As String
  Dim project As VBIDE.VBProject
  For Each project In Application.VBE.VBProjects
  If project Is ThisWorkbook.VBProject Then
      AddInName = project.name
    End If
  Next project
  'not return correct value (refer to personal.xlsm)
  'AddInName = Application.VBE.ActiveVBProject.name
End Function

Public Function AddInVersion() As String
  'not return correct value (refer to personal.xlsm)
  'AddInVersion = "v." & Application.VBE.ActiveVBProject.description
  Dim project As VBIDE.VBProject
  For Each project In Application.VBE.VBProjects
    If project Is ThisWorkbook.VBProject Then
      AddInVersion = "v." & project.description
    End If
  Next project
End Function

Public Function ExcelOwner() As String
  ExcelOwner = Application.UserName
End Function

Public Function OSVersion() As String
  OSVersion = Application.OperatingSystem
End Function

Public Function ExcelVersion() As String
  ExcelVersion = Application.Version
End Function

Public Function Build() As String
  Dim ver As String
  Dim chkref As Object
  Dim minor As String, minorup As String
  For Each chkref In ThisWorkbook.VBProject.References
    If chkref.name = "Excel" Then
      ver = RetrieveDllVersion(chkref.FullPath)
      'major = RetrievePart(ver, 0)
      'majorup = RetrievePart(ver, 1)
      minor = RetrievePart(ver, 2, ".")
      minorup = RetrievePart(ver, 3, ".")
    End If
  Next
  Build = minor & "." & minorup
End Function

Public Function ComputerName() As String
  Dim cName As String
  cName = GetMyComputerName()
  If Len(cName) <> 0 Then ComputerName = cName
End Function

Public Function Mac() As String
  Dim mAddress As String
  mAddress = GetMACAddress()
  If Len(mAddress) <> 0 Then Mac = mAddress
End Function

Public Function LocalIP() As String
  Dim lIP As String
  lIP = GetLocalIP()
  If Len(lIP) <> 0 Then LocalIP = lIP
End Function

Public Function PublicIP() As String
  Dim pIP As String
  pIP = GetPublicIP()
  If Len(pIP) <> 0 Then PublicIP = pIP
End Function

Public Function VBA() As String
  Dim bit64 As Boolean
  bit64 = IsOffice64bit
  If bit64 Then
    VBA = "64-bit"
  Else: VBA = "32-bit"
  End If
End Function

Public Function UserDownloadFolder() As String
  UserDownloadFolder = GetDownloadFolder
End Function

Private Function GetDownloadFolder() As String
  'Special folders: AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStorage,
  'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs, Recent
  'SendTo, StartMenu, Storage, Templates
  Dim tempPath As String
  'Dim tempPath1 As String, tempPath2 As String, tempPath3 As String
  Dim wshShell As Object
  Dim SpecialPath As String
  Dim folderNames As String
  Dim Folders() As String
  Dim i As Integer
  On Error GoTo FailExit: Const Proc As String = "GetDownloadFolder"
  folderNames = "Downloads,Загрузки,Завантаження"
  Folders = Split(folderNames, ",")
  Set wshShell = CreateObject("WScript.Shell")
  SpecialPath = wshShell.specialfolders("MyDocuments")
  'test 1: Windows XP???
  For i = 0 To UBound(Folders)
    tempPath = SpecialPath & "\" & Folders(i)
    If DirExist(tempPath) Then
      GetDownloadFolder = tempPath
      i = UBound(Folders) + 1                    'зачем єто?
      Exit For
    End If
  Next i
  'test 2: Windows 7++
  For i = 0 To UBound(Folders)
    tempPath = Left(SpecialPath, InStrRev(SpecialPath, "\") - 1) & "\" & Folders(i)
    If DirExist(tempPath) Then
      GetDownloadFolder = tempPath
      i = UBound(Folders) + 1
      Exit For
    End If
  Next i
  Set wshShell = Nothing
  Exit Function
FailExit:
  Set wshShell = Nothing
  'CatchLogger ErrorLevel, FunctionLevelError, , , Proc, className
End Function

Private Function IsOffice64bit() As Boolean
  Dim bIs64Bit As Boolean
  bIs64Bit = False
  #If Win64 Then
    bIs64Bit = True
  #End If
  IsOffice64bit = bIs64Bit
End Function

Private Function GetPublicIP() As String
  Dim HttpRequest As Object
  On Error Resume Next
  'Create the XMLHttpRequest object.
  Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
  'Check if the object was created.
  If Err.Number <> 0 Then
    'Return error message.
    GetPublicIP = "n/a"
    'Release the object and exit.
    Set HttpRequest = Nothing
    Exit Function
  End If
  On Error GoTo 0
  'Create the request - no special parameters required.
  HttpRequest.Open "GET", "http://myip.dnsomatic.com", False
  'Send the request to the site.
  HttpRequest.Send
  'Return the result of the request (the IP string).
  GetPublicIP = HttpRequest.ResponseText
  Set HttpRequest = Nothing
End Function

Private Function GetLocalIP() As String
  Dim strComputer     As String
  Dim objWMIService   As Object
  Dim colItems        As Object
  Dim objItem         As Object
  Dim myIPAddress     As String
  On Error GoTo FailExit: Const Proc As String = "GetLocalIP"
  'Set the computer.
  strComputer = "."
  'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  'A select query is used to get a collection of IP addresses from the network adapters that have the property IPEnabled equal to true.
  Set colItems = objWMIService.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
  'Loop through all the objects of the collection and return the first non-empty IP.
  For Each objItem In colItems
    If Not IsNull(objItem.IPAddress) Then myIPAddress = Trim(objItem.IPAddress(0))
    Exit For
  Next
  'Return the IP string.
  GetLocalIP = myIPAddress
  Set objWMIService = Nothing
  Set colItems = Nothing
  Set objItem = Nothing
  Exit Function
FailExit:
  Set objWMIService = Nothing
  Set colItems = Nothing
  Set objItem = Nothing
  'CatchLogger ErrorLevel, FunctionLevelError, , , Proc, className
End Function

Private Function GetMACAddress() As String
  Dim strComputer     As String
  Dim objWMIService   As Object
  Dim colItems        As Object
  Dim objItem         As Object
  Dim myMACAddress    As String
  On Error GoTo FailExit: Const Proc As String = "GetMACAddress"
  'Set the computer.
  strComputer = "."
  'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
  Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
  'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
  For Each objItem In colItems
    If Not IsNull(objItem.IPAddress) Then myMACAddress = objItem.macAddress
    Exit For
  Next
  'Return the IP string.
  GetMACAddress = myMACAddress
  Set objWMIService = Nothing
  Set colItems = Nothing
  Set objItem = Nothing
  Exit Function
FailExit:
  Set objWMIService = Nothing
  Set colItems = Nothing
  Set objItem = Nothing
  'CatchLogger ErrorLevel, FunctionLevelError, , , Proc, className
End Function

Private Function GetMyComputerName() As String
  Dim stBuff As String * 255
  #If Win64 Then
    Dim lAPIResult As LongPtr
  #Else
    Dim lAPIResult As Long
  #End If
  
  Dim lBuffLen As Long
  On Error GoTo FailExit: Const Proc As String = "GetMyComputerName"
  lBuffLen = 255
  lAPIResult = GetComputerName(stBuff, lBuffLen)
  If lBuffLen > 0 Then GetMyComputerName = Left(stBuff, lBuffLen)
  Exit Function
FailExit:
  'CatchLogger ErrorLevel, FunctionLevelError, , , Proc, className
End Function

Private Function RetrieveDllVersion(ByVal dll As String) As String
  Dim fso As Object                              'Scripting.FileSystemObject
  On Error GoTo FailExit: Const Proc As String = "RetrieveDllVersion"
  Set fso = CreateObject("Scripting.FileSystemObject")
  RetrieveDllVersion = fso.GetFileVersion(dll)
  Set fso = Nothing
  Exit Function
FailExit:
  Set fso = Nothing
  'CatchLogger ErrorLevel, FunctionLevelError, , , Proc, className
End Function

Private Function className() As String
  className = TypeName(Me)
End Function

'Public Function OfficeLanguage() As UILanguage
'  Dim ID As Long
'  On Error GoTo FailExit
'  ID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'  Select Case ID
'  Case 2057: OfficeLanguage = UILanguage.eng     'wdEnglishUK  United Kingdom English language.
'  Case 1033: OfficeLanguage = UILanguage.eng     'wdEnglishUS  United States English language.
'  Case 1058: OfficeLanguage = UILanguage.ukr     'wdUkrainian  Ukrainian language.
'  Case 1049: OfficeLanguage = UILanguage.rus     'wdRussian    Russian language.
'  Case Else: GoTo FailExit
'  End Select
'  Exit Function
'FailExit:
'
'End Function

Public Function OfficeDefaultSeparator() As String
  OfficeDefaultSeparator = Application.International(xlListSeparator)
End Function

