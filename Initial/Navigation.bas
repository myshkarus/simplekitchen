Attribute VB_Name = "Navigation"
Option Explicit

Public Const TableList As String = "TableList" 'лист со списком всех функциональных листов файла

Public Sub ProductList()
  Dim pView As ProductView
  Set pView = New ProductView
  pView.Interface(InterfaceType.ProductList).Activate
  Set pView = Nothing
End Sub

Public Sub ProductForm()
  Dim pView As ProductView
  Set pView = New ProductView
  pView.Interface(InterfaceType.ProductForm).Activate
  Set pView = Nothing
End Sub

Public Sub ProductChange()
  Dim pView As ProductView
  Set pView = New ProductView
  pView.Interface(InterfaceType.ProductChangeForm).Activate
  Set pView = Nothing
End Sub

Public Sub MeasureRef()
  Dim rView As ReferencesView
  Set rView = New ReferencesView
  rView.Interface(InterfaceType.ReferenceMeasureUnit).Activate
End Sub

Public Sub ProductTypeRef()
  Dim rView As ReferencesView
  Set rView = New ReferencesView
  rView.Interface(InterfaceType.ReferenceProductType).Activate
End Sub
