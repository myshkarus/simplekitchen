Attribute VB_Name = "ADOTest"
Option Explicit
Private dic As Scripting.Dictionary
Public Const Settings = "Settings"
Public Const DBPath = "DatabasePath"
'это заменить позднее:
Public Const path As String = "D:\_ ”’Ќя проект\cookinghouse.accdb"

Public Sub test()
Dim d As ADODictionary

Dim result As Object
  Dim data As New ADOWrapper
  'Dim sqlQuery As String: sqlQuery = "SELECT FullName, UnitOfMeasure, Note FROM tblUnitOfMeasure WHERE HighPriority=TRUE"
  Dim sqlQuery As String: sqlQuery = "SELECT * FROM tblSystemDictionary"
Set d = New ADODictionary
Set result = data.GetRecordset(sqlQuery)
Set d.rs = result
Set dic = d.mydictionary

Dim keys As Variant
Dim i As Integer

  keys = dic.keys
  For i = 0 To UBound(keys)
    Debug.Print dic.keys(i), dic.item(dic.keys(i))
  Next i
 
End Sub


Public Sub SetupSettings()
  Dim wsh As Worksheet
  On Error GoTo ErrHandler
  If Not WorksheetExist(Settings) Then
    ThisWorkbook.Worksheets.Add().Name = Settings
  End If
  Set wsh = ThisWorkbook.Worksheets(Settings)
  If Not RangeExists(DBPath) Then
    ThisWorkbook.Names.Add Name:=DBPath, RefersTo:=wsh.Range("A1")
    wsh.Range(DBPath) = path
  End If
  Exit Sub
ErrHandler:
End Sub

Public Function DatabasePath() As String
  Dim path As String
  path = ThisWorkbook.Names(DBPath).RefersToRange(1, 1) 'Range(DBPath).Value
  If Len(path) = 0 Then GoTo ErrHandler
  DatabasePath = path
  Exit Function
ErrHandler:
End Function

Public Sub TestADO()
  On Error GoTo errorHandler
  Dim i As Long
  Dim data As New ADOWrapper
  'Dim sqlQuery As String: sqlQuery = "SELECT FullName, UnitOfMeasure, Note FROM tblUnitOfMeasure WHERE HighPriority=TRUE"
  Dim sqlQuery As String: sqlQuery = "SELECT * FROM tblUnitOfMeasure WHERE HighPriority=TRUE"
  
  Dim result As Object
  Dim wsh As Worksheet
  
  Set wsh = ThisWorkbook.Worksheets("result")
  Set result = data.GetRecordset(sqlQuery)
    
  With wsh
    For i = 0 To result.Fields.count - 1         'r.Fields.Count - 1
      .Cells(1, i + 1) = result.Fields(i).Name
    Next i
    .Range("A2").CopyFromRecordset result
  End With
  

   ' Do your things with data here
  '  Dim i As Long, j As Long
  '  For i = 1 To result.RecordCount
  '    For j = 1 To result.Fields.Count
  '      Debug.Print result.Fields(j - 1)
  '    Next j
  '    result.MoveNext
  '  Next i
  'Call data.ListTablesADO
  Set result = Nothing
  Set data = Nothing
  Exit Sub
errorHandler:
  Debug.Print Err.Source & ", " & Err.Description
End Sub

