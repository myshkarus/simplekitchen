VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Depricated_ReferencesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CHKBOXINDENT As Long = 3

Private Type TSheetData
  sheet As Worksheet
  ProductTypeFirstCell As Range
  MeasureUnitFirstCell As Range
  'ProductFormFirstCell As Range
  'ProductFormRange As Range
  'ListFormFirstCell As Range
  'ListFormRange As Range
  Interface As InterfaceType
End Type

Private Type TView
  iView As InterfaceType
  Builded As Boolean
End Type

Private iProtected As Boolean

Private this As TSheetData
Private curInterface As TView
Private iReferences As IMetaData

Public Sub Activate()
  Dim ctrl As IControl
  Application.ScreenUpdating = False
  Select Case this.Interface
  Case InterfaceType.ReferenceProductType
    this.sheet.Range(this.MeasureUnitFirstCell, this.MeasureUnitFirstCell.Offset(0, 30)).EntireColumn.Hidden = True
    this.sheet.Range(this.ProductTypeFirstCell, this.MeasureUnitFirstCell.Offset(0, -1)).EntireColumn.Hidden = False
    For Each ctrl In rCollection
      With ctrl
        If .mark = "C" Then
          If .name = "ref_measure" Then .Style = csOff
          If .name = "ref_producttype" Then .Style = csOn
          .Movable = False
        End If
      End With
    Next ctrl
    this.sheet.Range("ref_вид").Select

  Case InterfaceType.ReferenceMeasureUnit
    this.sheet.Range(this.MeasureUnitFirstCell, this.MeasureUnitFirstCell.Offset(0, 30)).EntireColumn.Hidden = False
    this.sheet.Range(this.ProductTypeFirstCell, this.MeasureUnitFirstCell.Offset(0, -1)).EntireColumn.Hidden = True
    For Each ctrl In rCollection
      With ctrl
        If .mark = "C" Then
          If .name = "ref_measure" Then .Style = csOn
          If .name = "ref_producttype" Then .Style = csOff
          .Movable = False
        End If

        this.sheet.Range("measure_list").Cells(1, 1).Select
      End With
    Next ctrl
    
    
  End Select
  Application.ScreenUpdating = True
End Sub

Private Sub Build(view As TView)
  Dim Button As IControl
  Dim label As IControl
  Dim checkBox As IControl
  Dim listTable As IControl
  Dim rng As Range
  Dim rightEdge As Single, leftEdge As Single, topEdge As Single, bottomEdge As Single
  Dim marked As String
  Dim btnBuilder As ButtonBuilder
  Dim lblBuilder As LabelBuilder
  Dim chkBuilder As CheckBoxBuilder
  Dim tbBuilder As TableBuilder
  Dim titleProductType() As Variant
  Dim titleMeasure() As Variant
  'Dim dataToUpdate As Variant
  'Dim recount As Long
  'Dim lblControl As CLabel

  On Error GoTo FailExit
  Application.ScreenUpdating = False

  Set btnBuilder = NewButtonBuilder(this.sheet)
  Set chkBuilder = NewCheckBoxBuilder(this.sheet)
  Set lblBuilder = NewLabelBuilder(this.sheet)
  Set tbBuilder = NewTableBuilder(this.sheet)
  titleProductType = Array("_", "код", "вид продукта")
  titleMeasure = Array("_", "код", "сокращение", "полное наименование", "примечание")
  'Me.Protected = False
  marked = "C"
  Set rng = this.ProductTypeFirstCell
  Set Button = btnBuilder.caption("СПРАВОЧНИКИ").OfPlaceAndSize(rng.Left, rng.top + 5, 250, size.Dbl).Style(ControlStyle.csAccented).Movable(False).fontSize(FontSizeEnum.TitleLarge).visible(True).Build:  rCollection.Add Button
  Set Button = btnBuilder.name("ref_producttype").mark(marked).caption("виды продуктов").OfPlaceAndSize(Button.Left, Button.Bottom + 5, 100, size.Mdl + 5).Style(csOn).Movable(False).fontSize(FontSizeEnum.Medium).OnAction("ProductTypeRef").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_measure").mark(marked).caption("ед.измерения").OfPlaceAndSize(Button.Right + CHKBOXINDENT, Button.top, 100, size.Mdl + 5).Style(csOff).Movable(False).fontSize(FontSizeEnum.Medium).OnAction("MeasureRef").visible(True).Build: rCollection.Add Button
  Set rng = this.ProductTypeFirstCell.Offset(6, 0)
  Set label = lblBuilder.caption("Найти:").OfPlaceAndSize(rng).Style(csSimplified).fontSize(FontSizeEnum.Medium).Field(2, , , "ref_вид", , , True).Build

  'label.BottomRightRange.Offset(2, 0).Resize(5, 3).Select
    
  Set listTable = tbBuilder.name("producttype_list").OfPlaceAndSize(label.BottomRightRange.Offset(2, 0).Resize(5, 3)).Header(titleProductType).Build
  
  Set Button = btnBuilder.name("ref_deleteproducttype").mark("A").OfPlaceAndSize(listTable.Right - size.Std, listTable.top - size.Std - CHKBOXINDENT * 2).Style(ControlStyle.csAlternative).Movable(True).SetIcon("IconDelete").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_changeproducttype").mark("A").OfPlaceAndSize(Button.Left - size.Std - 15, Button.top, Button.Width, Button.height).Style(csOff).Movable(True).SetIcon("IconChange").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_addproducttype").mark("A").OfPlaceAndSize(Button.Left - size.Std - CHKBOXINDENT * 2, Button.top, Button.Width, Button.height).Style(csOff).Movable(True).SetIcon("IconAdd").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_sortproducttype").mark("A").OfPlaceAndSize(Button.Left - size.Std - CHKBOXINDENT * 2, Button.top, Button.height, Button.height).Style(csOff).Movable(True).SetIcon("IconSort").visible(True).Build: rCollection.Add Button

  leftEdge = rCollection.ItemByName("ref_producttype").Left
  rightEdge = rCollection.ItemByName("ref_measure").Right
  topEdge = rCollection.ItemByName("ref_producttype").Bottom + 3

  With New LineBuilder
    .DrawLine this.sheet, leftEdge, topEdge, rightEdge, topEdge
  End With

  Set rng = this.MeasureUnitFirstCell
    
  Set rng = this.MeasureUnitFirstCell.Offset(4, 0)
    
  Set listTable = tbBuilder.name("measure_list").OfPlaceAndSize(rng.Offset(2, 0).Resize(5, 5)).Header(titleMeasure).Build
  Set Button = btnBuilder.name("ref_deletemeasure").mark("A").OfPlaceAndSize(listTable.Right - size.Std, listTable.top - size.Std - CHKBOXINDENT * 2).Style(ControlStyle.csAlternative).Movable(True).SetIcon("IconDelete").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_changemeasure").mark("A").OfPlaceAndSize(Button.Left - size.Std - 15, Button.top, Button.Width, Button.height).Style(csOff).Movable(True).SetIcon("IconChange").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_addmeasure").mark("A").OfPlaceAndSize(Button.Left - size.Std - CHKBOXINDENT * 2, Button.top, Button.Width, Button.height).Style(csOff).Movable(True).SetIcon("IconAdd").visible(True).Build: rCollection.Add Button
  Set Button = btnBuilder.name("ref_sortmeasure").mark("A").OfPlaceAndSize(Button.Left - size.Std - CHKBOXINDENT * 2, Button.top, Button.height, Button.height).Style(csOff).Movable(True).SetIcon("IconSort").visible(True).Build: rCollection.Add Button
 
  recount = CountObjects
  dataToUpdate = iReferences.Update(Table.References, sheetControls, recount)

  Application.ScreenUpdating = True
  
  Set lblControl = Nothing
  Set btnBuilder = Nothing
  Set lblBuilder = Nothing
  Set chkBuilder = Nothing
  Set tbBuilder = Nothing
  Set Button = Nothing
  Set label = Nothing
  Set checkBox = Nothing
  Set listTable = Nothing
  Set rng = Nothing
  Exit Sub
FailExit:
  Err.Raise Err.Number, Err.source, Err.description
End Sub

Private Function CheckBuild() As Boolean
  Dim par As Variant
  par = iReferences.parameter(Table.References, Options.sheetControls)
  If CLng(par) = CountObjects() And CLng(par) <> 0 Then
    CheckBuild = True
  End If
End Function

Public Function Interface(ByVal iType As InterfaceType) As ReferencesView
  this.sheet.Activate
  With curInterface
    .iView = iType
    .Builded = CheckBuild
  End With
  If Not curInterface.Builded Then
    With New ControlCollection
      Set rCollection = .Self
    End With
    Call FormatSheet
    Call Build(curInterface)
  Else
    Set rCollection = Nothing                    'for testing
    If rCollection Is Nothing Then
      Set rCollection = RestoreCollection
    End If
  End If
  this.Interface = curInterface.iView
  Set Interface = Me
End Function

Public Property Get sheet() As Worksheet
  Set sheet = this.sheet
End Property

Public Function ToString() As String
  ToString = this.sheet.name
End Function

Private Sub Class_Initialize()
  Set iReferences = New MetaData
  Set this.sheet = iReferences.Create(Table.References)
  With this.sheet
    Set this.ProductTypeFirstCell = .Cells(1, StandardColumn.FirstLevel)
    Set this.MeasureUnitFirstCell = .Cells(1, StandardColumn.SecondLevel)
  End With
  'ActiveWindow.DisplayHeadings = False
End Sub

Private Sub Class_Terminate()
  Set iReferences = Nothing
  Set this.ProductTypeFirstCell = Nothing
  Set this.MeasureUnitFirstCell = Nothing
  Set this.sheet = Nothing
  ActiveWindow.DisplayHeadings = True
End Sub

Private Sub FormatSheet()
  With this.sheet
    .Cells.Interior.Color = Color.MainBackground
    .Range("A1").EntireColumn.Hidden = True
    .columns(this.ProductTypeFirstCell.Offset(0, -1).column).ColumnWidth = 2
    .columns(this.ProductTypeFirstCell.column).ColumnWidth = 3
    .columns(this.ProductTypeFirstCell.Offset(0, 1).column).ColumnWidth = 9
    .columns(this.ProductTypeFirstCell.Offset(0, 2).column).ColumnWidth = 90
    .columns(this.MeasureUnitFirstCell.column).ColumnWidth = 3
    .columns(this.MeasureUnitFirstCell.Offset(0, 1).column).ColumnWidth = 9
    .columns(this.MeasureUnitFirstCell.Offset(0, 2).column).ColumnWidth = 18
    .columns(this.MeasureUnitFirstCell.Offset(0, 3).column).ColumnWidth = 45
    .columns(this.MeasureUnitFirstCell.Offset(0, 4).column).ColumnWidth = 15
  End With
End Sub

Private Function RestoreCollection() As ControlCollection
  Dim sh As Shape, tempSh As Shape
  Dim sr As ShapeRange
  Dim control As IControl
  
  On Error GoTo FailExit
  With New ControlCollection
    For Each sh In this.sheet.Shapes
      If sh.Type = msoAutoShape Then
        If sh.AutoShapeType = msoShapeRectangle Then
          With New CButton: Set .ShapeBtn = sh: Set control = .Self: End With
        End If
      ElseIf sh.Type = msoFormControl Then
        If sh.FormControlType = xlCheckBox Then
          With New CCheckBox: Set .ShapeChkBox = sh.OLEFormat.Object: Set control = .Self: End With
        End If
      ElseIf sh.Type = msoGroup Then
        With New CButton: Set .sheet = this.sheet: Set .ShapeBtn = sh.GroupItems(1): Set control = .Self: End With
      End If
      .Add control
    Next sh
    Set RestoreCollection = .Self
    Set sh = Nothing
    Set tempSh = Nothing
    Set sr = Nothing
  End With
  Exit Function
FailExit:
  Debug.Print "restorecollection error"
End Function

Private Function CountObjects() As Long
  Dim sh As Shape
  Dim i As Long
  For Each sh In this.sheet.Shapes
    If sh.Type = msoAutoShape Then
      If sh.AutoShapeType = msoShapeRectangle Then i = i + 1
    ElseIf sh.Type = msoFormControl Then
      If sh.FormControlType = xlCheckBox Then i = i + 1
    ElseIf sh.Type = msoGroup Or sh.Type = msoLine Then
      i = i + 1
    End If
  Next sh
  CountObjects = i
End Function

'''''Public Property Let Protected(ByVal status As Boolean)
'''''  iProtected = status
'''''  With Display
'''''    If status Then
'''''      .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
'''''               AllowFormattingRows:=True, AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, AllowDeletingRows:=True, AllowSorting:= _
'''''               True, AllowFiltering:=True
'''''      .EnableSelection = xlUnlockedCells
'''''    Else: .Unprotect
'''''    End If
'''''  End With
'''''End Property
'''''
'''''Public Property Get Protected() As Boolean
'''''  Protected = iProtected
'''''End Property




