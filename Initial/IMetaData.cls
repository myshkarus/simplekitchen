VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "IMetaData"

'возможные типы пользовательских интерфейсов
Public Enum InterfaceType
  ProductList
  ProductForm
  ProductChangeForm
  ProductGroupList
  ReferenceProductType
  ReferenceMeasureUnit
  'iForm 'форм Add
  'iList 'таблицы со списками
  'iFormChange 'форма Change
End Enum

Public Enum StandardColumn
  FirstLevel = 3
  SecondLevel = 57
  ThirdLevel = 110
End Enum

Public Enum Options
  sheetName = 1
  sheetControls = 2
End Enum

Public Function Create(ByVal sheetID As table, Optional ByVal value As String) As Worksheet
Attribute Create.VB_UserMemId = 0
End Function

Public Function Update(ByVal sheetID As table, ByVal opt As Options, ByVal value As Variant) As Variant
End Function

Public Property Get parameter(ByVal sheetID As table, ByVal opt As Options) As Variant
End Property

Private Sub Class_Initialize()
  Err.Raise 5, ClassName, Err.Source & vbCrLf & "- Interface class must not be instantiated"
End Sub
