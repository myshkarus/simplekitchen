VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Set pCollection = Nothing
Set rCollection = Nothing
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
  If sh.Name = "продукты" Then
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayFormulas = False
    ActiveWindow.DisplayRuler = False
    Application.DisplayFormulaBar = True
  End If
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
  Dim sP As Shape
  Dim n As Name
  If sh.Name = "справочники" Then
'    For Each sP In sh.Shapes
'      sP.Delete
'    Next sP
'    ActiveWindow.DisplayGridlines = True
'    ActiveWindow.DisplayHeadings = True
'    ActiveWindow.DisplayFormulas = True
'    Application.DisplayFormulaBar = True
'    sh.Cells.Clear
'    sh.Cells.EntireColumn.ColumnWidth = 10
'  For Each n In ThisWorkbook.Names
'    'Debug.Print n.Value
'    If n.name <> "ID" Then n.Delete
'  Next n
   End If

  
End Sub


