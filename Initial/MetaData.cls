VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TInfo
  sysSheet As Worksheet
  navRange As Range
End Type

Private Enum lookAt
  ColumnEntire
  RowEntire
  WholeSheet
End Enum

Private this As TInfo

Implements IMetaData

Private Sub Class_Initialize()
  Dim ID As String: ID = "ID"
  On Error GoTo FailExit
  Set this.sysSheet = ThisWorkbook.Worksheets(TableList)
  If Not RangeExists(ID) Then
    ThisWorkbook.Names.Add Name:=ID, RefersTo:=FindRange("Table", this.sysSheet.Range("A1"))
  End If
  Set this.navRange = ThisWorkbook.Names(ID).RefersToRange
  Exit Sub
FailExit:
  MsgBox "Не найден лист 'TableList' с метаданными!" & vbCrLf & "Обратитесь к разработчику", vbCritical Or vbOKOnly, "Системная ошибка"
  End
End Sub

Private Sub Class_Terminate()
  With this
    Set .sysSheet = Nothing
    Set .navRange = Nothing
  End With
End Sub

Private Function IMetaData_Update(ByVal sheetID As table, ByVal opt As Options, ByVal value As Variant) As Variant
  Dim tempRng As Range
  On Error GoTo FailExit
  Set tempRng = FindRange(sheetID, , ColumnEntire).Offset(, opt)
  tempRng.value = value
  IMetaData_Update = value
  Set tempRng = Nothing
  Exit Function
FailExit:
End Function

Private Property Get IMetaData_Parameter(ByVal sheetID As table, ByVal opt As Options) As Variant
  IMetaData_Parameter = FindRange(sheetID, , ColumnEntire).Offset(, opt).value
End Property

Private Function IMetaData_Create(ByVal sheetID As table, Optional ByVal value As String) As Worksheet
  Dim tempName As String
  Dim wsh As Worksheet
  On Error GoTo FailExit
  tempName = FindRange(sheetID, , ColumnEntire).Offset(, Options.sheetName).value
  If tempName <> vbNullString Then
    If Not SheetExist(tempName) Then
      On Error GoTo 0
      Set wsh = ThisWorkbook.Worksheets.Add
      wsh.Name = tempName
      With wsh.CustomProperties
        .Add Name:=tempName, value:=value
      End With
    Else:
    Set wsh = ThisWorkbook.Worksheets(tempName)
  End If
End If
If Not wsh Is Nothing Then Set IMetaData_Create = wsh
Set wsh = Nothing
Exit Function
FailExit:
End Function

Private Function FindRange(ByVal strSearch As String, Optional ByVal startRange As Range, Optional ByVal rangeToLookAt As lookAt = WholeSheet) As Range
  Dim rng As Range
  If Len(strSearch) <> 0 Then
    If startRange Is Nothing Then
      Set startRange = this.navRange
    End If
    Select Case rangeToLookAt
    Case ColumnEntire
      Set rng = startRange.EntireColumn
    Case RowEntire
      Set rng = startRange.EntireRow
    Case Else
      Set rng = this.sysSheet.Cells
    End Select
    Set FindRange = rng.Find(strSearch, startRange)
  End If
  Set rng = Nothing
End Function

Private Function SheetExist(ByVal sheetName As String, Optional workbookName As String) As Boolean
  Dim objSheet As Object
  On Error GoTo FailExit
  If workbookName = vbNullString Then
    workbookName = ThisWorkbook.Name
  End If
  Set objSheet = Workbooks(workbookName).Sheets(sheetName)
  SheetExist = True
  If Not objSheet Is Nothing Then Set objSheet = Nothing
  Exit Function
FailExit:
  SheetExist = False
  If Not objSheet Is Nothing Then
    Set objSheet = Nothing
  End If
End Function

