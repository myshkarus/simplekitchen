Attribute VB_Name = "Utility"
Option Private Module
Option Explicit

Const dhcDelimiters As String = " ,.!:;<>?"

'���������: ��������� � ��������

#If Win64 Then
  #If VBA7 Then
    Private Declare PtrSafe Function IsCharAlphaNumericA Lib "user32" (ByVal bytChar As Byte) As LongPtr
    Private Declare PtrSafe Function IsCharAlphaNumericW Lib "user32" (ByVal intChar As Integer) As LongPtr
    Private Declare PtrSafe Function IsCharAlphaA Lib "user32" (ByVal byChar As Byte) As LongPtr
    Private Declare PtrSafe Function IsCharAlphaW Lib "user32" (ByVal intChar As Integer) As LongPtr
    Private Declare PtrSafe Function GetCPInfo Lib "kernel32" (ByVal CodePage As LongPtr, lpCPInfo As CPINFO) As LongPtr
  #Else
    Private Declare Function IsCharAlphaNumericA Lib "user32" (ByVal bytChar As Byte) As LongPtr
    Private Declare Function IsCharAlphaNumericW Lib "user32" (ByVal intChar As Integer) As LongPtr
    Private Declare Function IsCharAlphaA Lib "user32" (ByVal byChar As Byte) As LongPtr
    Private Declare Function IsCharAlphaW Lib "user32" (ByVal intChar As Integer) As LongPtr
    Private Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As LongPtr, lpCPInfo As CPINFO) As LongPtr
  #End If
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function IsCharAlphaNumericA Lib "user32" (ByVal bytChar As Byte) As Long
    Private Declare PtrSafe Function IsCharAlphaA Lib "user32" (ByVal bytChar As Byte) As Long
    Private Declare PtrSafe Function IsCharAlphaNumericW Lib "user32" (ByVal intChar As Integer) As Long
    Private Declare PtrSafe Function IsCharAlphaW Lib "user32" (ByVal intChar As Integer) As Long
    Private Declare PtrSafe Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
  #Else
    Private Declare Function IsCharAlphaNumericA Lib "user32" (ByVal bytChar As Byte) As Long
    Private Declare Function IsCharAlphaA Lib "user32" (ByVal bytChar As Byte) As Long
    Private Declare Function IsCharAlphaNumericW Lib "user32" (ByVal intChar As Integer) As Long
    Private Declare Function IsCharAlphaW Lib "user32" (ByVal intChar As Integer) As Long
    Private Declare Function GetCPInfo Lib "kernel32" (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
  #End If
#End If

#If Win64 Then
  #If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As LongPtr) As LongPtr
    Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal flags As Long) As LongPtr
  #Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As LongPtr) As LongPtr
    Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal flags As Long) As LongPtr
  #End If
#Else
  #If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
    Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
  #Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
    Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
  #End If
#End If

'names of preinstalled WorkSheets:
Public Const LANG_SETTINGS As String = "dictionary"
Public Const MAIN_SETTINGS As String = "main"
Public Const ERRORLIST As String = "errorlist"

'Public Const DEFAULT_LANGUAGE As Integer = 1

Public Const VINDENT As Integer = 10: Public Const VSHORTINDENT As Integer = 7: Public Const HINDENT As Integer = 5
Public Const DEFAULTHEIGHT As Integer = 18: Public Const BUTTONSHIFT As Integer = 1
Public Const DEFAULTFONT As String = "Tahoma": Public Const DEFAULTTEXTFONT As String = "Tahoma": Public Const DEFAULTFONTSIZE As Integer = 9
'Const VSHORTINDENT = 4: Const DEFAULTFONTSIZE = 8: Const VSHORTINDENT = 6: Const DEFAULTHEIGHT = 19

Public Enum UILanguage
  ukr = 1
  rus = 2
  eng = 3
End Enum

Public Enum EntryAttribute
  EntryPrivate = 0             ' don't display in user inteface
  EntryReadOnly = 1            ' read-only EntryUser
  EntryReadWrite = 2           ' read-write EntryUser
  EntryReadProtectedWrite = 3  ' read-write with password
End Enum

Public Enum LogDescription '5 digit number
  ClassConstructorFinished = 50000  '��������� ������ ����������� ������
  ClassConstructorError = 50001     '���������� �������� � ������������� ������
  ClassTerminated = 50002           '��������� ������ ������
  KeyRecorded = 50003               '������� ������ ����� ��������:='" & key & "'" & " � ���� '" & IEntry_Source & "'; eAttribute:=" & currentAttribute
  UnexistedCellError = 50004        '������� ��������� ������c������� ������
  NewAttributeRecorded = 50005      '��� ������ ���������� ����� �������: " & EntryType
  RecordNewValue = 50006            '����������� ����� �������� ��� ������
  DeferredRecordSucceded = 50007    '����������� ���������� ������ ��������
  InterfaceLanguageError = 50008    '������ ��� ��������� ����� ����������
  FormError = 50009                 '� ������ ����� ��������� ������
  ParentFormError = 50010           '������ ��� ��������� ������������ �����
  FormInitializationError = 50011   '���������� �������� � �������������� �����
  FormInitialization = 50012        '������������� �����...
  FormActivationError = 50013       '��� ��������� ����� ��������� ������
  FormTuningError = 50014           '��� �������� ����� ��������� ������
  FormTuningCompleted = 50015       '��������� ��������� �����
  FormTermination = 50016           '��������� ������ �����
  FormTerminationError = 50017      '���������� �������� ��� ���������� ������ �����
  FormCompletingError = 50018       '������ ���������� �����
  ControlFocusError = 50019         '���� ��� ��������� ������ �� ������� �����
  PasswordConfirmationError = 50020 '������ ��� ������������� ������
  OldPasswordError = 50021          '������� ������ ������ ������
  AllFieldsToBeCompleted = 50022    '��������� ��� ����������� ����!
  RecordRemoved = 50023             '������� ������
  RecordsCounted = 50024            '���������� ���������� ������� �
  RecordsCleared = 50025            '������� ������� � ��������
  KeyExistError = 50026             '������ ����������. ���� ��� ���������� key:=
  KeyNewRegistered = 50027          '��������������� ����� ���� key:=
  RecordFound = 50028               '������� ������ � ������: key=
  RecordNotFound = 50029            '�� ������� ������ � ������: key=
  OutOfRangeError = 50030           'Subscript out of range
  RecordReturnedByKey = 50031       '���������� ������ �� �����: key=
  RecordReturnedByValue = 50032     '���������� ������ �� ��������: value=
  RecordPointerDefined = 50033      '��������� ��������� �� ��� �������
  RecordPointerNotDefined = 50034   '�� ������� ���������� ��������� �� ��� �������
  FreeLoginError = 50035            '������ � �������� ������� ��������� �������
  NewUserAdded = 50036              '�������� ����� ������������
  GeneralOptionsRestoredError = 50037 '������ ��� ������ ����� ��������
  GeneralOptionsRestoredInfo = 50038 '�������� ����� ���������
  LoginListCreationError = 50039    '������ ��� ���������� ������ �������
  DefaultUserOptionsError = 50040   '������ ��� ���������� ���������� ��� ������ ������������
  DefaultUserOptionsInfo = 50041    '��������� ���������� ���������� ��� ������ ������������
  LanguageRestoredInfo = 50042      '��������� �������������� �������������� �������
  LanguageRestoredError = 50043     '������ ��� �������������� �������������� �������
End Enum


' Gather character width information.
Private Const CP_ACP = 0  '  default to ANSI code page
Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.
Private Type CPINFO
    MaxCharSize As Long                    '  max length (Byte) of a char
    DefaultChar(MAX_DEFAULTCHAR) As Byte   '  default character
    LeadByte(MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type


' Maximum length for Soundex string.
Private Const dhcLen = 4
'--------------------------

Public Log As ILog '����� ������

Public Sub AddNameForRibbonPointer(ByVal name As String, ByVal address As String)
  ThisWorkbook.Names.Add name:=name, RefersTo:="=" & address, visible:=False
End Sub

Public Sub RestoreRibbon()
  If myRibbon Is Nothing Then
    #If VBA7 Then
      Dim lPointer As LongPtr
      lPointer = CLngPtr([RibbonPointer])
    #Else
      Dim lPointer As Long
      lPointer = CLng([RibbonPointer])
    #End If
  CopyMemory myRibbon, lPointer, LenB(lPointer)
  End If
End Sub

'
'Dim numLayouts As Long
'Dim i As Long
'Dim layouts() As LongPtr
'
'numLayouts = GetKeyboardLayoutList(0, ByVal 0&)
'ReDim layouts(numLayouts - 1)
'GetKeyboardLayoutList numLayouts, layouts(0)
'
'Dim msg As String
'msg = "Loaded keyboard layouts: " & vbCrLf & vbCrLf
'
'For i = 0 To numLayouts - 1
'   msg = msg & Hex(layouts(i)) & vbCrLf
'Next
'
'MsgBox msg
Public Sub LayoutsList()
''���������� �������
'ActivateKeyboardLayout 68748313, 1 '�������
'ActivateKeyboardLayout 67699721, 2 '����������
'ActivateKeyboardLayout 69338146, 3 '����������
'����������������� �������
ActivateKeyboardLayout &H4190419, 1 '�������
'ActivateKeyboardLayout &H4090409, 2 '����������
'ActivateKeyboardLayout &H4220422, 3 '����������
End Sub



Public Sub CatchLogger(ByVal logger As LogLevel, Optional ByVal err As ErrorType, _
                        Optional ByVal description As LogDescription, Optional parameters As String, _
                        Optional ByVal Source As String, Optional ByVal className As String)
  Dim logMessage As String
  'Dim logClassName As String
  On Error Resume Next
  If Log Is Nothing Then
    Set Log = New logger: Log.Register
  End If
  'parameters - ���������
  'logClassName = className
  If logger = ErrorLevel Then
    logMessage = Log.StandardLogString(logger, err, description, Source, className)
  Else:
  '�������� ��������������, ����� ��������� ������������������:
    Exit Sub '��� ����� ������
    ' logMessage = Log.StandardLogString(logger, , description, Source, className)
  End If
  Log.Append logger, logMessage
End Sub

Public Function ReplaceChar(ByVal text As String, ByVal char As String, ByVal pos As Integer) As String
  Dim newText As String
  newText = vbNullString
  If pos <= 0 Then Exit Function
  If pos > Len(text) Then pos = Len(text)
  newText = Left(text, pos - 1) & char & Right(text, Len(text) - pos)
  ReplaceChar = newText
End Function

Public Function AddChar(ByVal text As String, ByVal char As String, ByVal pos As Integer) As String
  Dim newText As String
  newText = vbNullString
  If pos < 0 Then Exit Function
  If pos > Len(text) Then pos = Len(text)
  Select Case pos
  Case 0
    newText = char & text
  Case Len(text)
    newText = text & char
  Case Else
    newText = Left(text, pos) & char & Right(text, Len(text) - pos)
  End Select
  AddChar = newText
End Function

Public Function ExcludeChar(ByVal text As String, ByVal pos As Integer) As String
  Dim newText As String
  newText = vbNullString
  If Len(text) = 0 Or pos <= 0 Then Exit Function
  If pos > Len(text) Then Exit Function
  Select Case pos
  Case 1
    newText = Right(text, Len(text) - 1)
  Case Len(text)
    newText = Left(text, Len(text) - 1)
  Case Else
    newText = Left(text, pos - 1) & Right(text, Len(text) - pos)
  End Select
  ExcludeChar = newText
End Function





Public Sub ClearLogger()
  Dim File As FileWriter
  Set File = New FileWriter
    File.OpenFile "D:\log.txt", True
    File.CloseFile
  Set File = Nothing
End Sub

Public Function RangeExists(ByVal r As String) As Boolean
    Dim test As Range
    On Error Resume Next
    Set test = Range(r)
    RangeExists = err.Number = 0
    Set test = Nothing
End Function

Public Function CreateClass(className As String) As Object
  Dim modified As String
  modified = "C" & className
  Set CreateClass = Application.Run(modified)
End Function

Public Function CMap() As EntryMap
  Set CMap = New EntryMap
  'CSetting.Initialize
End Function

Public Function CSetting() As EntryUser
  Set CSetting = New EntryUser
  'CSetting.Initialize
End Function

Public Sub SayHello()
    MsgBox "Hello " & Application.UserName
End Sub

Public Function WorksheetExist(ByVal wSheet As String, Optional wBook As String) As Boolean
  Dim sheet As Object
  On Error GoTo ErrHandler

  If wBook = vbNullString Then wBook = ThisWorkbook.name
  Set sheet = Workbooks(wBook).Worksheets(wSheet)
  WorksheetExist = True
  If Not sheet Is Nothing Then Set sheet = Nothing
  Exit Function
ErrHandler:
  WorksheetExist = False
  If Not sheet Is Nothing Then
    Set sheet = Nothing
  End If
End Function

Public Function SheetReady(ByVal sheetName As String, ByVal TableHeader As String, Optional ByVal BookName As String) As Boolean
  Dim mySheet As Worksheet
  Dim myBook As Workbook
  Dim ready As Boolean
  Dim customHeader() As String
  Dim index As Integer
  On Error GoTo FailExit
  If sheetName = vbNullString Then Exit Function
  '� sheetName �� ������ ���� ����������� ������ + ���� ����������
  If Len(BookName) > 0 Then
    Set myBook = Workbooks(BookName)
  Else: Set myBook = ThisWorkbook
  End If
  If Not WorksheetExist(sheetName, myBook.name) Then
    myBook.Worksheets.Add().name = sheetName
  End If
  'for worksheet it needs to realize:
'    With source
'      .AutoFilter.ShowAllData: .Cells.EntireColumn.Hidden = False: .Cells.EntireRow.Hidden = False
'    End With
  
  If Len(TableHeader) > 0 Then
    customHeader = Split(TableHeader, ";")
    If Not myBook.Worksheets(sheetName) Is Nothing Then
      Set mySheet = myBook.Worksheets(sheetName)
      For index = 0 To UBound(customHeader)
        If mySheet.Cells(1, 1).Offset(0, index).value <> customHeader(index) Then
          mySheet.Cells(1, 1).Offset(0, index).value = customHeader(index)
        End If
        ready = True
      Next index
    Else: GoTo FailExit
    End If
  End If
CleanExit:
  If Not mySheet Is Nothing Then
    SheetReady = ready
    Set mySheet = Nothing
    Set myBook = Nothing
  End If
  Exit Function
FailExit:
  SheetReady = False
  err.Raise vbObjectError + 106, "������ 'Utility' (������� SheetReady)", _
     "������ �������� �������."
End Function

' WARNING: this provides only very basic security and should
' not be used to protect sensitive data.
Public Function ValidPassword(ByVal Password As String) As Boolean
'  Dim oSetting As EntryUser
'  Dim bValid As Boolean
'  bValid = False
'  Set oSetting = New EntryUser
'  If oSetting.GetSetting("Password") Then
'    If oSetting.value = sPassword Then
'      bValid = True
'    Else
'      bValid = False
'    End If
'  Else
'    bValid = False
'  End If
'  Set oSetting = Nothing
'  ValidPassword = bValid
ValidPassword = True
End Function

Public Function RetrievePart(ByVal strArray As Variant, ByVal pos As Integer, Optional ByVal separator As String) As String
  On Error Resume Next
  If Len(separator) = 0 Then
    RetrievePart = Split(strArray, "|")(pos)
  Else:
    RetrievePart = Split(strArray, separator)(pos)
  End If
End Function

Public Function OnlyLiterals(ByVal name As String) As String
  Dim sText As String, str As String, i As Integer
  sText = Trim(name)
  For i = 1 To Len(sText)
    If Not IsNumeric(Mid(sText, i, 1)) Then
      str = str & Mid(sText, i, 1)
    End If
  Next i
  OnlyLiterals = str
End Function

Public Function OnlyDigits(ByVal name As String) As String
  Dim sText As String, str As String, i As Integer
  sText = Trim(name)
  For i = Len(sText) To 1 Step -1
    If IsNumeric(Mid(sText, i, 1)) Then
      str = Mid(sText, i, 1) & str
    End If
  Next i
  OnlyDitits = str
End Function


'===================


Public Function dhExtractString(ByVal strIn As String, _
 ByVal intPiece As Integer, _
 Optional ByVal strDelimiter As String = dhcDelimiters) As String
    
    ' Pull tokens out of a delimited list.  strIn is the
    ' list, and intPiece tells which chunk to pull out.
    '
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strIn:
    '       String in which to search.
    '   intPiece:
    '       Integer indicating the particular chunk to retrieve.
    '       If this value is larger than the number of available
    '       tokens, the function returns "".
    '   strDelimiter (optional):
    '       String containing one or more characters to be used as
    '       token delimiters.
    '       If the delimiter's not found, the function returns "".
    ' Out:
    '   Return Value:
    '       The requested token from strIn. See the examples.
    ' Examples:

    '   dhExtractString("This,is,a,test", 1, ",") == "This"
    '   dhExtractString("This,is,a,test", 2, ",") == "is"
    '   dhExtractString("This,is,a,test", 4, ",") == "test"
    '   dhExtractString("This,is,a,test", 5, ",") == ""
    '   dhExtractString("This is a test", 2, " ") == "is"
    
    ' Note: if delimiter isn't found, output is the whole string.
    '   dhExtractString("Hello", 1, " ") = "Hello"
    
    ' You might think this function would be faster
    ' using the built-in Split function, but it's not. The code
    ' might be simpler, but it always takes a bit longer to run.
    ' This code stops as soon as it's pulled off the piece
    ' it wants, but Split breaks apart the entire input string.
    
    ' Doubled delimiters contain an empty token between them.
    '   dhExtractString("Hello", 1, "l") == "He"
    '   dhExtractString("Hello", 2, "l") == ""
    '   dhExtractString("Hello", 3, "l") == "o"
    '
    '   dhExtractString("This:is;a?test", 1, ":;? ") == "This"
    
    ' Requires:
    '   dhTranslate
    
    ' Used by:
    '   dhExtractCollection
    '   dhFirstWord
    '   dhLastWord
    
    Dim lngPos As Long
    Dim lngPos1 As Long
    Dim lngLastPos As Long
    Dim intLoop As Integer

    lngPos = 0
    lngLastPos = 0
    intLoop = intPiece
    
    ' If there's more than one delimiter, EntryMap them
    ' all to the first one.
    If Len(strDelimiter) > 1 Then
        strIn = dhTranslate(strIn, strDelimiter, _
         Left$(strDelimiter, 1))
    End If
    
    Do While intLoop > 0
        lngLastPos = lngPos
        lngPos1 = InStr(lngPos + 1, strIn, Left$(strDelimiter, 1))
        If lngPos1 > 0 Then
            lngPos = lngPos1
            intLoop = intLoop - 1
        Else
            lngPos = Len(strIn) + 1
            Exit Do
        End If
    Loop
    ' If the string wasn't found, and this wasn't
    ' the first pass through (intLoop would equal intPiece
    ' in that case) and intLoop > 1, then you've run
    ' out of chunks before you've found the chunk you
    ' want. That is, the chunk number was too large.
    ' Return "" in that case.
    If (lngPos1 = 0) And (intLoop <> intPiece) And (intLoop > 1) Then
        dhExtractString = vbNullString
    Else
        dhExtractString = Mid$(strIn, lngLastPos + 1, _
         lngPos - lngLastPos - 1)
    End If
End Function

Public Function dhExtractCollection(ByVal strText As String, _
 Optional ByVal strDelimiter As String = dhcDelimiters) As Collection
 
    ' Return a collection containing all the tokens contained
    ' in a String, using the supplied delimiters.
        
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   strDelimiter (optional, default = dhcDelimiters)
    '       String composed of characters that act
    '       as delimiters. If unspecified, use the
    '       delimiters in dhcDelimiters.
    ' Out:
    '   Return Value:
    '       Collection filled with the tokens from the
    '       input String, extracted using the supplied
    '       delimiters
    ' Example:
    '   dhExtractCollection("This is a test") returns a collection
    '       that contains the items: "this", "is", "a", "test"
    ' Requires:
    '   dhTranslate
    '   dhExtractString
 
    Dim colWords As Collection
    Dim lngI As Long
    Dim strTemp As String
    Dim strChar As String * 1
    Dim astrItems() As String
    
    Set colWords = New Collection
    
    ' If there's more than one delimiter, EntryMap them
    ' all to the first one.
    If Len(strDelimiter) = 0 Then
        colWords.Add strText
    Else
        strChar = Left$(strDelimiter, 1)
        If Len(strDelimiter) > 1 Then
            strText = dhTranslate(strText, strDelimiter, strChar)
        End If
            
        astrItems = Split(strText, strChar)
    
        ' Loop through all the tokens, adding them to the
        ' output collection.
        For lngI = LBound(astrItems) To UBound(astrItems)
            colWords.Add astrItems(lngI)
        Next lngI
    End If
    
    ' Return the output collection.
    Set dhExtractCollection = colWords
End Function

Public Function dhCountIn(strText As String, strFind As String, _
 Optional lngCompare As VbCompareMethod = vbBinaryCompare) As Long
    
    ' Determine the number of times strFind appears in strText
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   strFind:
    '       Text to find within strText
    '   lngCompare (Optional, default is vbCompareBinary):
    '       Indicates how the search should compare values:
    '           vbBinaryCompare: case-sensitive
    '           vbTextCompare: case-insensitive
    '           vbDatabaseCompare (Doesn't work here)
    '           Any LocaleID value: compare as if in the selected locale.
    
    ' Out:
    '   Return Value:
    '       The number of times strFind appears in
    '       strText, respecting the lngCompare flag.
    ' Example:
    '   dhCountIn("This is a test", "is") returns 2
    
    ' Used by:
    '   dhExtractCollection
    '   dhCountWords
    '   dhCountTokens
    
    Dim lngCount As Long
    Dim lngPos As Long
    
    ' If there's nothing to find, there surely can't be any
    ' found, so return 0.
    If Len(strFind) > 0 Then
        lngPos = 1
        Do
            lngPos = InStr(lngPos, strText, strFind, lngCompare)
            If lngPos > 0 Then
                lngCount = lngCount + 1
                lngPos = lngPos + Len(strFind)
            End If
        Loop While lngPos > 0
    Else
        lngCount = 0
    End If
    dhCountIn = lngCount
End Function

Public Function TrimAll(ByVal strText As String, _
 Optional fRemoveTabs As Boolean = True) As String
    
    ' Remove leading and trailing white space, and
    ' reduce any amount of internal white space (including tab
    ' characters) to a single space.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   fRemoveTabs (Optional, default True):
    '       Should the code remove tabs, too?
    ' Out:
    '   Return Value:
    '       Input text, with leading and trailing white space removed
    ' Example:
    '   dhTrimAll("   this   is a     test ") returns "this is a test"
    ' Requires:
    '   dhTranslate
    ' Used by:
    '   dhCountWords
    
    Dim strTemp As String
    Dim strOut As String
    Dim lngI As Long
    Dim strCh As String * 1
    
    ' Trim off white space from the front and back.
    ' If requested, first convert all tabs into spaces,
    ' or RTrim and LTrim will miss them.
    If fRemoveTabs Then
        strText = Translate(strText, vbTab, " ")
    End If
    strTemp = Trim$(strText)
    For lngI = 1 To Len(strTemp)
        ' Look at each character, in turn.
        strCh = Mid$(strTemp, lngI, 1)
        
        ' If this character a space, and the previous
        ' added character was a space? If not, add it on.
        If Not (strCh = " " And Right$(strOut, 1) = " ") Then
            strOut = strOut & strCh
        End If
    Next lngI
    TrimAll = strOut
End Function

Public Function dhCountWords(ByVal strText As String) As Long
    
    ' Return the number of words in a string.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    ' Out:
    '   Return Value:
    '       The number of words, separated by spaces, in strText
    ' Example:
    '   dhCountWords("Hi there, my name is Cleo, what's yours?") returns 8
    
    ' Requires:
    '   dhTrimAll
    '   dhTranslate
    '   dhCountIn
    '   dhcDelimiters
    
    ' Used by:
    '   dhLastWord
    
    If Len(strText) = 0 Then
        dhCountWords = 0
    Else
        ' Get rid of any extraneous stuff, including delimiters and
        ' spaces. First convert delimiters to spaces, and then
        ' remove all extraneous spaces.
        strText = dhTrimAll(dhTranslate(strText, dhcDelimiters, " "))
        ' If there are three spaces, there are
        ' four words, right?
        dhCountWords = dhCountIn(strText, " ") + 1
    End If
End Function

Public Function dhCountTokens(ByVal strText As String, _
 ByVal strDelimiter As String, _
 Optional lngCompare As VbCompareMethod = vbBinaryCompare) As Long
 
    ' Return the number of tokens, given a set of delimiters, in a string
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Text to be analyzed
    '   strDelimiter:
    '       One or more delimiter characters, in a string
    ' Out:
    '   Return Value:
    '       Number of delimiters + 1, which should be the number of tokens
    '       Two delimiters in a row returns an empty token between them.
    ' Example:
    '   dhCountTokens("This is a test", " ") returns 4
    '   dhCountTokens("This:is!a test", ": !") returns 4
    '   dhCountTokens("This:!:is:!:a:!:test", ": !") returns 10
    '       They are:
    '           This, "", "", is, "", "", a, "", "", test
    ' Requires:
    '   dhTranslate
    '   dhCountIn
    
    Dim strChar As String * 1
    
    ' If there's no search text, there can't be any tokens.
    If Len(strText) = 0 Then
        dhCountTokens = 0
    ElseIf Len(strDelimiter) = 0 Then
        ' If there's no delimiters, the output
        ' is the entire input.
        dhCountTokens = 1
    Else
        strChar = Left$(strDelimiter, 1)
        
        ' Flatten all the delimiters to just the first one in
        ' the list.
        If Len(strDelimiter) > 1 Then
            strText = Translate(strText, strDelimiter, _
             strChar, lngCompare)
        End If
        ' Count the tokens. Actually, count
        ' delimiters, and add one.
        dhCountTokens = dhCountIn(strText, strChar) + 1
    End If
End Function

Public Function dhFirstWord( _
 ByVal strText As String, _
 Optional ByRef strRest As String = "") As String
    
    ' Retrieve the first word of a string
    ' Fill strRest with the rest of the string
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       The input text
    ' Out:
    '   strRest (optional):
    '       If supplied, filled in with the rest of the string
    '       including any separating spaces or other
    '       delimiters
    '   Return Value:
    '       The first word in strText
    ' Example:
    '   dhFirstWord("This is a test", strRest) returns "This"
    '       and places " is a test" into strRest.
    ' Requires:
    '   dhExtractString
    
    Dim strTemp As String
    
    ' This is easy!
    ' Get the first word.
    strTemp = dhExtractString(strText, 1)
    
    ' Extract everything after the first word,
    ' and put that into strRest.
    strRest = Mid$(strText, Len(strTemp) + 1)
    
    ' Return the first word.
    dhFirstWord = strTemp
End Function

Public Function dhLastWord( _
 ByVal strText As String, _
 Optional ByRef strRest As String = "") As String
    
    ' Retrieve the last word of a string
    ' Fill strRest with the rest of the string
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       The input text
    ' Out:
    '   strRest (optional):
    '       If supplied, filled in with the rest of the string
    '       including any separating spaces or other
    '       delimiters
    '   Return Value:
    '       The last word in strText
    ' Example:
    '   dhLastWord("This is a test", strRest) returns "test"
    '       and places "This is a " into strRest.
    ' Requires:
    '   dhTrimAll
    '   dhTranslate
    
    Dim strTemp As String
    Dim astrItems() As String
    
    ' This is not quite so easy.
    ' Get rid of any extraneous stuff, including delimiters and
    ' spaces. First convert delimiters to spaces, and then
    ' remove all extraneous spaces.
    strText = dhTrimAll(dhTranslate(strText, dhcDelimiters, " "))
    astrItems = Split(strText)
    strTemp = astrItems(UBound(astrItems))
    
    ' Extract everything before the last word,
    ' and put that into strRest.
    strRest = Left$(strText, Len(strText) - Len(strTemp))
    dhLastWord = strTemp
End Function

Public Function Translate( _
 ByVal strIn As String, _
 ByVal strMapIn As String, _
 ByVal strMapOut As String, _
 Optional lngCompare As VbCompareMethod = vbBinaryCompare) As String
    
    ' Take a list of characters in strMapIn, match them
    ' one-to-one in strMapOut, and perform a character
    ' replacement in strIn.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strIn:
    '       String in which to perform replacements
    '   strMapIn:
    '       EntryMap of characters to find
    '   strMapOut:
    '       EntryMap of characters to replace.  If the length
    '       of this string is shorter than that of strMapIn,
    '       use the final character in the string for all
    '       subsequent matches.
    '       If strMapOut is empty, just delete all the characters
    '       in strMapIn.
    '       If strMapOut is shorter than strMapIn, rightfill strMapOut
    '       with its final character. That is:
    '           dhTranslate(someString, "ABCDE", "X")
    '       is equivalent to
    '           dhTranslate(someString, "ABCDE", "XXXXX")
    '       That makes it simple to replace a bunch of characters with
    '       a single character.
    '   lngCompare (Optional, default is vbCompareBinary):
    '       Indicates how the search should compare values:
    '           vbBinaryCompare: case-sensitive
    '           vbTextCompare: case-insensitive
    '           vbDatabaseCompare (Doesn't work here)
    '           Any LocaleID value: compare as if in the selected locale.
    ' Out:
    '   Return Value:
    '       strIn, with appropriate replacements
    ' Example:
    '   dhTranslate("This is a test", "aeiou", "AEIOU") returns
    '     "ThIs Is A tEst"
    '   dhTranslate(someString, _
    '    "���������������������������������������������������������", _
    '    "AAAAAAAaaaaaaaEEEEeeeeIIIIiiiiNnOOOOOoooooOoUUUUuuuuYyysD")
    '     returns someString with 8-bit characters flattened
    '
    ' Used by:
    '   dhExtractString
    '   dhExtractCollection
    '   dhTrimAll
    '   dhCountWords
    '   dhCountTokens
    
    Dim lngI As Long
    Dim lngPos As Long
    Dim strChar As String * 1
    Dim strOut As String
    
    ' If there's no list of characters
    ' to replace, there's no point going on
    ' with the work in this function.
    If Len(strMapIn) > 0 Then
        ' Right-fill the strMapOut set.
        If Len(strMapOut) > 0 Then
            strMapOut = Left$(strMapOut & String(Len(strMapIn), _
             Right$(strMapOut, 1)), Len(strMapIn))
        End If
        
        For lngI = 1 To Len(strIn)
            strChar = Mid$(strIn, lngI, 1)
            lngPos = InStr(1, strMapIn, strChar, lngCompare)
            If lngPos > 0 Then
                ' If strMapOut is empty, this doesn't fail,
                ' because Mid handles empty strings gracefully.
                strOut = strOut & Mid$(strMapOut, lngPos, 1)
            Else
                strOut = strOut & strChar
            End If
        Next lngI
    End If
    Translate = strOut
End Function

'Public Function dhProperLookup( _
' ByVal strIn As String, _
' Optional blnForceToLower As Boolean = True, _
' Optional rst As ADODB.Recordset = Nothing, _
' Optional strField As String = "") As Variant
'
'    ' Convert a word to Proper case, using optional
'    ' lookup table for word spellings.
'
'    ' Suggested by code posted to Compuserve's MSACCESS forum
'    ' by Emmanuel Soheyli (75333,1003)
'
'    ' In:
'    '   strIn:
'    '       Input string to be converted, word by word.
'    '   blnForceToLower (Optional, default = True):
'    '       convert all letters except the first to lower case?
'    '   rst (Optional, default = Nothing):
'    '       ADODB recordset in which to look for word matches
'    '   strField (Optional, default = ""):
'    '       Field in rst in which to search. If you supply rst, you
'    '       must supply a field name in strField.
'    ' Out:
'    '   Return Value:
'    '       The "properized" string.
'    ' Example:
'    '   See TestProper test function.
'
'    ' Requires:
'    '   dhFixWord
'    '   dhIsCharAlphaNumeric
'
'    Dim strOut As String
'    Dim strWord As String
'    Dim lngI As Long
'    Dim strC As String * 1
'
'    On Error GoTo HandleErr
'
'    strOut = vbNullString
'    strWord = vbNullString
'
'    If blnForceToLower Then
'        strIn = LCase$(strIn)
'    End If
'
'    For lngI = 1 To Len(strIn)
'        strC = Mid$(strIn, lngI, 1)
'        If dhIsCharAlphaNumeric(strC) Or strC = "'" Then
'            strWord = strWord & strC
'        Else
'            strOut = strOut & dhFixWord(strWord, rst, strField) & strC
'            ' Reset strWord for the next word.
'            strWord = vbNullString
'        End If
'NextChar:
'    Next lngI
'
'    ' Process the final word.
'    strOut = strOut & dhFixWord(strWord, rst, strField)
'
'ExitHere:
'    dhProperLookup = strOut
'    Exit Function
'
'HandleErr:
'    ' If there's an error, just go on to the next character.
'    ' This may mean the output word is missing characters,
'    ' of course. If this bothers you, just change the Resume
'    ' statement to resume at "ExitHere."
'    Select Case err
'        Case Else
'            ' MsgBox "Error: " & Err.Description & " (" & Err.Number & ")"
'    End Select
'    Resume NextChar
'
'End Function
'
'Private Function dhFixWord( _
' ByVal strWord As String, _
' Optional rst As ADODB.Recordset = Nothing, _
' Optional strField As String = "") As String
'
'    ' Used by:
'    '   dhProperLookup
'
'    ' "Properize" a single word
'    Dim strOut As String
'
'    On Error GoTo HandleErr
'
'    If Len(strWord) > 0 Then
'        ' Many things can go wrong. Just assume you want the
'        ' standard properized version unless you hear otherwise.
'        strOut = UCase(Left$(strWord, 1)) & Mid$(strWord, 2)
'        ' Did you pass in a recordset? If so, lookup
'        ' the value now.
'        If Not rst Is Nothing Then
'            If Len(strField) > 0 Then
'                rst.MoveFirst
'                rst.Find strField & " = " & _
'                 "'" & Replace(strWord, "'", "''") & "'"
'                If Not rst.EOF Then
'                    strOut = rst(strField)
'                End If
'            End If
'        End If
'    End If
'
'ExitHere:
'    dhFixWord = strOut
'    Exit Function
'
'HandleErr:
'    ' If anything goes wrong, anything, just get out.
'    Select Case err.Number
'        Case Else
'            ' MsgBox "Error: " & Err.Description & " (" & Err.Number & ")"
'    End Select
'    Resume ExitHere
'End Function
'
Public Function dhIsCharAlpha(strText As String) As Boolean
    ' Is the first character of strText an alphabetic character?
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Text to check. Only first character will be examined.
    ' Out:
    '   Return Value:
    '       True if first character of strText is alphabetic in
    '       the current locale.
    ' Example:
    '   If dhIsCharAlpha(strSomeValue) Then
    '       ' you know the first character is alphabetic.
    '   End If
    
    ' Requires:
    '   dhIsCharsetWide
    ' Used by:
    '   dhSoundex
    '   dhIsCharNumeric
    
    If dhIsCharsetWide() Then
        dhIsCharAlpha = CBool(IsCharAlphaW(AscW(strText)))
    Else
        dhIsCharAlpha = CBool(IsCharAlphaA(Asc(strText)))
    End If
End Function

Public Function dhIsCharAlphaNumeric(strText As String) As Boolean
    ' Is the first character of strText an alphanumeric character?
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Text to check. Only first character will be examined.
    ' Out:
    '   Return Value:
    '       True if first character of strText is alphanumeric in
    '       the current locale.
    ' Example:
    '   If dhIsCharAlphaNumeric(strSomeValue) Then
    '       ' you know the first character is alphanumeric.
    '   End If
    
    ' Requires:
    '   dhIsCharsetWide
    '
    ' Used by:
    '   dhIsCharNumeric
    
    If dhIsCharsetWide() Then
        dhIsCharAlphaNumeric = CBool(IsCharAlphaNumericW(AscW(strText)))
    Else
        dhIsCharAlphaNumeric = CBool(IsCharAlphaNumericA(Asc(strText)))
    End If
End Function

Public Function dhIsCharNumeric(strText As String) As Boolean
    ' Is the first character of strText a numeric character?
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Text to check. Only first character will be examined.
    ' Out:
    '   Return Value:
    '       True if first character of strText is numeric.
    ' Example:
    '   If dhIsCharNumeric(strSomeValue) Then
    '       ' you know the first character is numeric
    '   End If
    
    ' Requires:
    '   dhIsCharsetWide

    dhIsCharNumeric = dhIsCharAlphaNumeric(strText) _
     And Not dhIsCharAlpha(strText)
End Function

Public Function dhIsCharNumeric1(strText As String) As Boolean
    ' Is the first character of strText a numeric character?
    ' Doesn't use any API calls.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Text to check. Only first character will be examined.
    ' Out:
    '   Return Value:
    '       True if first character of strText is numeric.
    ' Example:
    '   If dhIsCharNumeric1(strSomeValue) Then
    '       ' you know the first character is numeric
    '   End If
    
    dhIsCharNumeric1 = (strText Like "[0-9]*")
End Function

Public Function dhIsCharsetWide() As Boolean
    ' Get the maximum character width of the
    ' operating system font.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   Nothing
    ' Out:
    '   Return Value:
    '       True if the current character set is either
    '       DBCS or Unicode, False otherwise.
    ' Example:
    '   See dhIsCharAlpha
    
    Dim tSystemFontInfo As CPINFO
    
    Call GetCPInfo(CP_ACP, tSystemFontInfo)
    dhIsCharsetWide = (tSystemFontInfo.MaxCharSize > 1)
End Function

Public Function dhTrimNull(ByVal strValue As String) As String
    ' Find the first vbNullChar in a string, and return
    ' everything prior to that character.
    ' Useful when combined with the Windows API function calls.
    
    ' Note: No matter what you've read, what you've seen,
    ' or whose code you've borrowed, the Trim function will
    ' not accomplish the goal of this function, and this
    ' is a goal you'll need fulfilled often.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strValue:
    '       Input text, possibly containing a null character
    '       (chr$(0), or vbNullChar)
    ' Out:
    '   Return Value:
    '       strValue trimmed on the right, at the location
    '       of the null character, if there was one.
    
    Dim lngPos As Long
    
    lngPos = InStr(strValue, vbNullChar)
    Select Case lngPos
        Case 0
            ' Not found at all, so just
            ' return the original value.
            dhTrimNull = strValue
        Case 1
            ' Found at the first position, so return
            ' an empty string.
            dhTrimNull = vbNullString
        Case Is > 1
            ' Found in the string, so return the portion
            ' up to the null character.
            dhTrimNull = Left$(strValue, lngPos - 1)
    End Select
End Function

Public Function dhTokenReplace(ByVal strIn As String, _
 ParamArray varItems() As Variant) As String
    
    ' Replace %1, %2, %3, etc., with the values passed in varItems.
    ' Using numbered, replaceable parameters is necessary in order
    ' to allow you to place text strings to be translated into a table,
    ' and make the appropriate replacements at run-time.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strIn:
    '       String containing text and replaceable parameters, in the form
    '       %1, %2, etc.
    '   varItems():
    '       Array of items to place into strIn.
    ' Out:
    '   Return Value:
    '       Input text with the replacements made.
    ' Example:
    '   dhTokenReplace("This %1 a %2 of %3 this works.", "is", "test", "how")
    '       returns "This is a test of how this works"
    
    ' WARNING: If you pass an array, rather than a delimited list, this
    ' code won't work correctly. Make sure you call this as shown in
    ' the example.
    
    On Error GoTo HandleErr
    
    Dim lngPos As Long
    Dim strReplace As String
    Dim intI As Integer
    
    For intI = UBound(varItems) To LBound(varItems) Step -1
        strReplace = "%" & (intI + 1)
        lngPos = InStr(1, strIn, strReplace)
        If lngPos > 0 Then
            strIn = Left$(strIn, lngPos - 1) & _
             varItems(intI) & Mid$(strIn, lngPos + Len(strReplace))
        End If
    Next intI
    
ExitHere:
    dhTokenReplace = strIn
    Exit Function
    
HandleErr:
    ' If any error occurs, just return the
    ' string as it currently exists.
    Select Case err.Number
        Case Else
            ' MsgBox "Error: " & Err.Description & _
            '  " (" & Err.Number & ")"
    End Select
    Resume ExitHere
End Function

Public Function dhSoundex(ByVal strIn As String) As String
    
    ' Create a Soundex lookup string for the
    ' input text.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strIn:
    '       The text to encode
    ' Out:
    '   Return value:
    '       strIn, converted to Soundex format.
    
    ' Requires:
    '   dhcLen
    '   dhIsCharAlpha
    '   dhPadRight
    
    Dim strOut As String
    Dim intI As Integer
    Dim intPrev As Integer
    Dim strChar As String * 1
    Dim intChar As Integer
    Dim blnPrevSeparator As Boolean
    
    strOut = ""
    strIn = UCase(strIn)
    blnPrevSeparator = True
    
    strOut = Left$(strIn, 1)
    For intI = 2 To Len(strIn)
        ' If the output string is full, quit now.
        If Len(strOut) >= dhcLen Then
            Exit For
        End If
        ' Get each character, in turn. If the
        ' character's a letter, handle it.
        strChar = Mid$(strIn, intI, 1)
        If dhIsCharAlpha(strChar) Then
            ' Convert the character to its code.
            intChar = CharCode(strChar)
                    
            ' If the character's not empty, and if it's not
            ' the same as the previous character, tack it
            ' onto the end of the string.
            If (intChar > 0) Then
                If blnPrevSeparator Or (intChar <> intPrev) Then
                    strOut = strOut & intChar
                    intPrev = intChar
                End If
            End If
            blnPrevSeparator = (intChar = 0)
        End If
    Next intI
    ' Return the string, right padded with 0's.
    dhSoundex = dhPadRight(strOut, dhcLen, "0")
End Function

Private Function CharCode(strChar As String) As Integer
    Select Case strChar
        Case "A", "E", "I", "O", "U", "Y"
            CharCode = 0
        Case "C", "G", "J", "K", "Q", "S", "X", "Z"
            CharCode = 2
        Case "D", "T"
            CharCode = 3
        Case "M", "N"
            CharCode = 5
        Case "B", "F", "P", "V"
            CharCode = 1
        Case "L"
            CharCode = 4
        Case "R"
            CharCode = 6
        Case Else
            CharCode = -1
    End Select
End Function

Public Function dhSoundsLike(ByVal strItem1 As String, _
 ByVal strItem2 As String, _
 Optional blnIsSoundex As Boolean = False) As Integer
 
    ' Return a number between 0 and 4 (4 being the best) indicating
    ' the similarity between the Soundex representation for
    ' two strings.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
        
    ' Requires:
    '   dhSoundex
    '   dhcLen
    
    ' In:
    '   strItem1 , strItem2:
    '       Strings to compare
    '   blnIsSoundex (Optional, default False):
    '       Are the strings already in Soundex format?
    ' Out:
    '   Return Value:
    '       Integer between 0 (not similar) and dhcLen (very similar) indicating
    '       the similarity in the Soundex representation of the two strings.
    ' Note:
    '   This code is extremely low-tech. Don't laugh! It just compares
    '   the two Soundex strings until it doesn't find a match, and returns
    '   the position where the two diverged.
    '
    '   Remember, two Soundex strings are completely different if the
    '   original words start with different characters. That is, this
    '   function always returns 0 unless the two words begin with the
    '   same character.
    
    Dim intI As Integer
    
    If Not blnIsSoundex Then
        strItem1 = dhSoundex(strItem1)
        strItem2 = dhSoundex(strItem2)
    End If
    For intI = 1 To dhcLen
        If Mid$(strItem1, intI, 1) <> Mid$(strItem2, intI, 1) Then
            Exit For
        End If
    Next intI
    dhSoundsLike = (intI - 1)
End Function

Public Function dhOrdinal(lngItem As Long) As String
    ' Given an integer, return a string
    ' representing the ordinal value.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   lngItem:
    '       Long value to be converted to ordinal
    ' Out:
    '   Return Value:
    '       String containing ordinal value
    ' Example:
    '   dhOrdinal(34) returns "34th"
    '   dhOrdinal(1) returns "1st"
    
    Dim intDigit As Integer
    Dim strOut As String * 2
    Dim intTemp As Integer
    
    ' All teens use "th"
    intTemp = lngItem Mod 100
    If intTemp >= 11 And intTemp <= 19 Then
        strOut = "th"
    Else
        ' Get that final digit
        intDigit = lngItem Mod 10
        Select Case intDigit
            Case 1
                strOut = "st"
            Case 2
                strOut = "nd"
            Case 3
                strOut = "rd"
            Case Else
                strOut = "th"
        End Select
    End If
    dhOrdinal = lngItem & strOut
End Function

Public Function dhRoman(ByVal intValue As Integer) As String
    
    ' Convert a decimal number between 1 and 3999
    ' into a Roman number.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   intValue:
    '       A value between 1 and 3999 to be converted
    '       to roman numerals.
    ' Out:
    '   Return Value:
    '       The roman numeral representation of the integer
    ' Example:
    '   dhRoman(123) returns "CXXIII"
    
    Dim varDigits As Variant
    Dim lngPos As Integer
    Dim intDigit As Integer
    Dim strTemp As String
    
    ' Build up the array of roman digits
    varDigits = Array("I", "V", "X", "L", "C", "D", "M")
    lngPos = LBound(varDigits)
    strTemp = ""
    Do While intValue > 0
        intDigit = intValue Mod 10
        intValue = intValue \ 10
        Select Case intDigit
            Case 1
                strTemp = varDigits(lngPos) & strTemp
            Case 2
                strTemp = varDigits(lngPos) & _
                 varDigits(lngPos) & strTemp
            Case 3
                strTemp = varDigits(lngPos) & _
                 varDigits(lngPos) & varDigits(lngPos) & strTemp
            Case 4
                strTemp = varDigits(lngPos) & _
                 varDigits(lngPos + 1) & strTemp
            Case 5
                strTemp = varDigits(lngPos + 1) & strTemp
            Case 6
                strTemp = varDigits(lngPos + 1) & _
                 varDigits(lngPos) & strTemp
            Case 7
                strTemp = varDigits(lngPos + 1) & _
                 varDigits(lngPos) & varDigits(lngPos) & strTemp
            Case 8
                strTemp = varDigits(lngPos + 1) & _
                 varDigits(lngPos) & varDigits(lngPos) & _
                 varDigits(lngPos) & strTemp
            Case 9
                strTemp = varDigits(lngPos) & _
                 varDigits(lngPos + 2) & strTemp
        End Select
        lngPos = lngPos + 2
    Loop
    dhRoman = strTemp
End Function

Public Function dhPadLeft(strText As String, intWidth As Integer, _
 Optional strPad As String = " ") As String
 
    ' Pad strText on the left, so the whole output is
    ' at least intWidth characters.
    ' If strText is longer than intWidth, just return strText.
    ' If strPad is wider than one character, this code only takes
    '  the first character got padding.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   intWidth:
    '       Minimum width of the output. If
    '       Len(strText) < intWidth, then the
    '       output will be exactly intWidth characters
    '       wide. The code will not truncate strText,
    '       no matter what.
    '   strPad (Optional, default is " "):
    '       string whose first character will
    '       be used to pad the output.
    ' Out:
    '   Return Value:
    '       strText, possibly padded on the left with
    '       the first character of strPad.
    ' Example:
    '   dhPadLeft("Name", 10, ".") returns
    '       "......Name"
    '   dhPadLeft("Name", 10) returns
    '       "      Name"
    
    If Len(strText) > intWidth Then
        dhPadLeft = strText
    Else
        dhPadLeft = Right$(String(intWidth, strPad) & _
         strText, intWidth)
    End If
End Function

Public Function dhPadRight(strText As String, intWidth As Integer, _
 Optional strPad As String = " ") As String
  
    ' Pad strText on the right, so the whole output is
    ' at least intWidth characters.
    ' If strText is longer than intWidth, just return strText.
    ' If strPad is wider than one character, this code only takes
    '  the first character got padding.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   intWidth:
    '       Minimum width of the output. If
    '       Len(strText) < intWidth, then the
    '       output will be exactly intWidth characters
    '       wide. The code will not truncate strText,
    '       no matter what.
    '   strPad (Optional, default is " "):
    '       string whose first character will
    '       be used to pad the output.
    ' Out:
    '   Return Value:
    '       strText, possibly padded on the right with
    '       the first character of strPad.
    ' Example:
    '   dhPadRight("Name", 10, ".") returns
    '       "Name......"
    '
    
    If Len(strText) > intWidth Then
        dhPadRight = strText
    Else
        dhPadRight = Left$(strText & _
         String(intWidth, strPad), intWidth)
    End If
    
End Function

Public Function dhXORText(strText As String, strPWD As String) _
 As String
    ' Encrypt or decrypt a string using the XOR operator.
    
    ' From "Visual Basic EntryLanguage Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 2000; Sybex, Inc. All rights reserved.
    
    ' In:
    '   strText:
    '       Input text
    '   strPWD:
    '       Password to be used for encryption.
    ' Out:
    '   Return Value:
    '       The encrypted/decrypted string.
    
    Dim abytText() As Byte
    Dim abytPWD() As Byte
    Dim intPWDPos As Integer
    Dim intPWDLen As Integer
    Dim intChar As Integer
    
    abytText = strText
    abytPWD = strPWD
    intPWDLen = LenB(strPWD)
    For intChar = 0 To LenB(strText) - 1
        ' Get the next number between 0 and intPWDLen - 1
        intPWDPos = (intChar Mod intPWDLen)
        abytText(intChar) = abytText(intChar) Xor _
         abytPWD(intPWDPos)
    Next intChar
    dhXORText = abytText
End Function




'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\








