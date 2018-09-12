Attribute VB_Name = "Helper"
Option Explicit

Public Function RangeAtPosition(ByVal wsh As Worksheet, ByVal left As Single, ByVal top As Single) As Range
  Dim suggRng As Range
  Dim rng As Range, usedRng As Range
  Dim x As Long, y As Long
  Dim row As Long, column As Long
  On Error GoTo FailExit
  
  If Not wsh Is Nothing Then
    Set usedRng = wsh.UsedRange
    For Each rng In usedRng.Resize(usedRng.rows.count + 3, 1)
      If top >= rng.top And top < rng.Offset(1, 0).top Then
        row = rng.row
        Exit For
      End If
    Next rng
    For Each rng In usedRng.Resize(1, usedRng.columns.count + 3)
      If left >= rng.left And left < rng.Offset(0, 1).left Then
        column = rng.column
        Exit For
      End If
      Set rng = rng.Offset(0, 1)
    Next rng
    Set suggRng = wsh.Cells(row, column)
    If Not suggRng Is Nothing Then
      'Debug.Print suggRng.address
      Set RangeAtPosition = suggRng
    End If
    Exit Function
  End If
FailExit:
Set RangeAtPosition = Nothing
End Function

Public Sub ABC_Click(Optional ByVal litera As String)
  Dim buttonName As String, subName As String
  Dim offButton As Boolean
  subName = vbNullString
  buttonName = ActiveSheet.Shapes(Application.Caller).Name
  
  If left$(buttonName, 3) <> "All" Then
    ActiveSheet.Shapes("All_on").Visible = msoFalse
    ActiveSheet.Shapes("All_off").Visible = msoTrue
  Else
    Call ResetAlphabet
  End If
  Select Case ButtonType(buttonName)
  Case bType.typeOff
    subName = left$(buttonName, InStr(1, buttonName, "_off") - 1)
    ActiveSheet.Shapes(buttonName).Visible = msoFalse
    ActiveSheet.Shapes(subName & "_on").Visible = msoTrue
  Case bType.typeOn
    subName = left$(buttonName, InStr(1, buttonName, "_on") - 1)
    ActiveSheet.Shapes(subName & "_off").Visible = msoTrue
    ActiveSheet.Shapes(buttonName).Visible = msoFalse
  End Select
  
  If litera <> vbNullString Then
  'Debug.Print litera
  End If
End Sub

Public Sub ResetAlphabet()
  Dim sh As Shape
  For Each sh In ActiveSheet.Shapes
    Select Case ButtonType(sh.Name)
    Case bType.typeOff: sh.Visible = msoTrue
    Case bType.typeOn:  sh.Visible = msoFalse
    Case Else:
    End Select
  Next sh
End Sub

Public Function ButtonType(ByVal Name As String) As bType
  If Right$(Name, 3) = "_on" Then
    ButtonType = typeOn
  ElseIf Right$(Name, 4) = "_off" Then
    ButtonType = typeOff
  End If
End Function

Private Function FileName(text As String) As String
  Debug.Print text
  ID = CStr(Right(Trim(text), 8))
  FileName = ID
End Function

'Public Function searchRange(wsh As Worksheet, strValue As String, _
'             Optional startRange As Range, Optional LookRange As lookAt = WholeSheet) As Range
'Dim rangeToSearch As Range
'On Error GoTo errorHandler
'If (wsh Is Nothing) Or Len(strValue) = 0 Then Exit Function
'If startRange Is Nothing Then Set startRange = wsh.Cells(1, 1)
'
'Select Case LookRange
'  Case lookAt.ColumnEntire: Set rangeToSearch = startRange.EntireColumn
'  Case lookAt.RowEntire: Set rangeToSearch = startRange.EntireRow
'  Case Else: Set rangeToSearch = wsh.Cells
'End Select
'
'With wsh.Range(rangeToSearch.Address)
'  Set searchRange = .Find(What:=strValue, After:=startRange, LookIn:=xlValues, lookAt:=xlWhole)
'End With
'Exit Function
'errorHandler:
'  Set rangeToSearch = Nothing
'End Function

Public Function TargetRange(wsh As Worksheet, ByVal startrng As Range) As Range
  Dim sRow As Long, sCol As Long
  Dim row As Long, col As Long
  Dim rng As Range
  If (wsh Is Nothing) Or (startrng Is Nothing) Then Exit Function


  sRow = startrng.row: sCol = startrng.column
  row = wsh.Range(Cells(65536, sCol).address).End(xlUp).row: col = wsh.Range(Cells(sRow, 255).address).End(xlToLeft).column

  If row > 1 And wsh.Cells(row, sCol) = "" Then
    Do While (wsh.Cells(row, sCol) = "")
      If row > 1 Then row = row - 1
    Loop
  End If

  If col > 1 And wsh.Cells(sRow, col) = "" Then
    Do While (wsh.Cells(sRow, col) = "")
      If col > 1 Then col = col - 1
    Loop
  End If

  Set rng = wsh.Range(Cells(sRow, sCol).address, Cells(row, col).address)
  If Not rng Is Nothing Then
    Set TargetRange = rng
    Set rng = Nothing
  End If
End Function

Public Sub ClearFilter(ByVal wsName As String, Optional wbName As String)
  Dim wsh As Worksheet
  On Error GoTo ErrHandler
    
  If wbName = vbNullString Then wbName = ThisWorkbook.Name
  Set wsh = Workbooks(wbName).Worksheets(wsName)
   
  If Not wsh Is Nothing Then
    If wsh.AutoFilterMode Then wsh.AutoFilter.ShowAllData
    wsh.Cells.EntireColumn.Hidden = False: wsh.Cells.EntireRow.Hidden = False
  End If
  Exit Sub
ErrHandler:
  Debug.Print "Error in ClearFilter function: " & Err.Source, Err.Number, Err.Description
End Sub

Public Sub MakeSheetVisible(ByVal status As Boolean, ByVal InitStatus As Integer, ByVal wsName As String, Optional wbName As String)
  Dim wsh As Worksheet
  Dim wshState As Integer

  If wbName = vbNullString Then
    Set wsh = ThisWorkbook.Worksheets(wsName)
  Else:
  Set wsh = Workbooks(wbName).Worksheets(wsName)
End If

If status Then
  If (wsh.Visible = xlSheetHidden Or wsh.Visible = xlSheetVeryHidden) Then wsh.Visible = xlSheetVisible
Else:
wsh.Visible = InitStatus
End If
  
If Not wsh Is Nothing Then Set wsh = Nothing
End Sub

Public Function GetFileName(ByVal strFullPath As String) As String
  Dim strFind As String
  Dim iCount As Integer
    
  Do Until left(strFind, 1) = "\"
    iCount = iCount + 1
    strFind = Right(strFullPath, iCount)
    If iCount = Len(strFullPath) Then Exit Do
  Loop

  GetFileName = Right(strFind, Len(strFind) - 1)
End Function

Public Function DirExist(ByVal sPath As String) As Boolean
  Dim Exist As String
  On Error GoTo ErrHandler
  Exist = Dir(sPath, vbDirectory)
  If Exist = vbNullString Then
    DirExist = False
  Else: DirExist = True
  End If
  Exit Function
ErrHandler:
  'Application.DisplayStatusBar = True
  'Application.StatusBar = "ÓÊÀÆÈÒÅ ÏÐÀÂÈËÜÍÛÉ ÏÓÒÜ Ê ÔÀÉËÓ!!!"
  'DirExist = True
End Function

Public Function GetDirectory(ByVal path As String)
  If Len(path) > 0 Then
    GetDirectory = left(path, InStrRev(path, "\"))
  End If
End Function

'Public Function SheetExist(ByVal wsName As String, Optional wbName As String) As Boolean
'Dim objSheet As Object
'On Error GoTo errHandler
'
'If wbName = vbNullString Then wbName = ThisWorkbook.name
'Set objSheet = Workbooks(wbName).Sheets(wsName)
'SheetExist = True
'If Not objSheet Is Nothing Then Set objSheet = Nothing
'Exit Function
'errHandler:
'  SheetExist = False
'  If Not objSheet Is Nothing Then
'    Set objSheet = Nothing
'  End If
'End Function

Public Function WorkbookIsOpen(ByVal wbName As String) As Boolean
  Dim b As Object
  For Each b In Workbooks
    If StrComp(b.Name, wbName, vbTextCompare) = 0 Then
      WorkbookIsOpen = True
      Exit Function
    End If
  Next
  WorkbookIsOpen = False
End Function

Public Function RangeExists(ByVal r As String) As Boolean
  Dim test As Range
  On Error Resume Next
  Set test = Range(r)
  RangeExists = Err.Number = 0
  Set test = Nothing
End Function

Public Function Hash(str As String) As String
  Dim bytes() As Byte, i&, lo&, hi&
  lo = &H9DC5&
  hi = &H11C&
  bytes = str
  For i = 0 To UBound(bytes) Step 2
    lo = 31& * ((bytes(i) + bytes(i + 1) * 256&) Xor (lo And 65535))
    hi = 31& * hi + lo \ 65536 And 65535
  Next
  lo = (lo And 65535) + (hi And 32767) * 65536 Or (&H80000000 And -(hi And 32768))
  Hash = Hex(lo)
End Function


