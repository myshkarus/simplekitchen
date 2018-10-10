VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Depricated_ProductView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const ABC_CYR As String = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЭЮЯ"
Private Const ABC_ENG As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const CHKBOXINDENT As Long = 3

Private Type TSheetData
  sheet As Worksheet
  ProductFormFirstCell As Range
  ProductFormRange As Range
  ListFormFirstCell As Range
  ListFormRange As Range
  Interface As InterfaceType
End Type

Private Type TView
  iView As InterfaceType
  Builded As Boolean
End Type

Private iProtected As Boolean

Private this As TSheetData
Private curInterface As TView
Private iProduct As IMetaData

Public Sub Activate()
  Dim ctrl As IControl
  Application.ScreenUpdating = False
  Select Case this.Interface
  Case InterfaceType.ProductList
    this.sheet.Range(this.ProductFormFirstCell, this.ProductFormFirstCell.Offset(0, 30)).EntireColumn.Hidden = True
    this.sheet.Range(this.ListFormFirstCell, this.ProductFormFirstCell.Offset(0, -1)).EntireColumn.Hidden = False
    For Each ctrl In pCollection
      With ctrl
        If .mark = "L" Then
          .visible = True
        ElseIf .mark = "F" Then
          .visible = False
        End If
        If .name = "product_selecttitle" Then .caption = "выбрать"
        If .name = "product_producttitle" Then .caption = "продукт"
        If .name = "product_productchange" Then .visible = True
        If .name = "product_productgroupchange" Then .visible = True
        If .name = "product_listreturn" Then .visible = False
        If .name = "product_productadd" Then .visible = True
      End With
    Next ctrl
    this.sheet.Range("A1").Select
  Case InterfaceType.ProductForm
    this.sheet.Range(this.ProductFormFirstCell, this.ProductFormFirstCell.Offset(0, 30)).EntireColumn.Hidden = False
    this.sheet.Range(this.ListFormFirstCell, this.ProductFormFirstCell.Offset(0, -1)).EntireColumn.Hidden = True
    For Each ctrl In pCollection
      With ctrl
        If .mark = "L" Then
          .visible = False
        ElseIf .mark = "F" Then
          .visible = True
        End If
        If .name = "product_product" Then .caption = "ДОБАВИТЬ ПРОДУКТ"
        
        If .name = "product_selecttitle" Then .caption = "сохранить"
        If .name = "product_producttitle" Then .caption = "к списку"
        If .name = "product_productchange" Then .visible = False
        If .name = "product_productgroupchange" Then .visible = False
        If .name = "product_listreturn" Then .visible = True
        If .name = "product_productadd" Then .visible = False
      End With
    Next ctrl
    this.sheet.Range("product_вид").Select
    
  Case InterfaceType.ProductChangeForm
    this.sheet.Range(this.ProductFormFirstCell, this.ProductFormFirstCell.Offset(0, 30)).EntireColumn.Hidden = False
    this.sheet.Range(this.ListFormFirstCell, this.ProductFormFirstCell.Offset(0, -1)).EntireColumn.Hidden = True
    For Each ctrl In pCollection
      With ctrl
        If .mark = "L" Then
          .visible = False
        ElseIf .mark = "F" Then
          .visible = True
        End If

        If .name = "product_product" Then .caption = "ИЗМЕНИТЬ ПРОДУКТ"
        If .name = "product_selecttitle" Then .caption = "сохранить"
        If .name = "product_producttitle" Then .caption = "к списку"
        If .name = "product_productchange" Then .visible = False
        If .name = "product_productgroupchange" Then .visible = False
        If .name = "product_listreturn" Then .visible = True
        If .name = "product_productadd" Then .visible = False
      End With
    Next ctrl
    this.sheet.Range("product_вид").Select
    
    
  End Select
  Application.ScreenUpdating = True
End Sub

Private Sub Build(view As TView)
  Dim Button As IControl
  Dim label As IControl
  Dim checkBox As IControl
  Dim rng As Range
  Dim rightEdge As Single, leftEdge As Single, topEdge As Single, bottomEdge As Single
  Dim marked As String
  Dim btnBuilder As ButtonBuilder
  Dim lblBuilder As LabelBuilder
  Dim chkBuilder As CheckBoxBuilder
  Dim tblBuilder As TableBuilder
  Dim dataToUpdate As Variant
  Dim recount As Long
  Dim lblControl As CLabel
  Dim tbl As IControl
  Dim titleListProduct As Variant
  titleListProduct = Array("_", "код", "вид", "продукт", "ед.изм.", "остаток")

  On Error GoTo FailExit
  Application.ScreenUpdating = False

  Set btnBuilder = NewButtonBuilder(this.sheet)
  Set chkBuilder = NewCheckBoxBuilder(this.sheet)
  Set lblBuilder = NewLabelBuilder(this.sheet)
  Set tblBuilder = NewTableBuilder(this.sheet)

  'Me.Protected = False
  marked = "L"
  Set rng = this.ListFormFirstCell
  Set Button = btnBuilder _
  .caption("ПОДОБРАТЬ ПРОДУКТЫ") _
  .OfPlaceAndSize(rng.Left, rng.top + 2, 300, cs.Double).Style(ControlStyle.csAccented).Movable(True).fontSize(fs.Title).visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.caption("для блюда").OfPlaceAndSize(rng.Left, Button.Bottom, 90, cs.Middle).Style(ControlStyle.csAccented).Movable(True).fontSize(fs.SubTitle).visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_dish").caption("Cуп овощной rwertert j;werjtkew;ktew ;k's;k'w;ket'ew;lkt  lk';lk';k';lk  ';k';lk';k';lk").OfPlaceAndSize(Button.Right, Button.top - 3, 450, cs.Middle * 2).Style(ControlStyle.csUnderlined).Movable(True).fontSize(fs.DataTitle).visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.caption("Критерий поиска:").OfPlaceAndSize(rng.Left, Button.Bottom, 100, cs.Standard).Style(ControlStyle.csSimplified).Movable(True).fontSize(fs.Button).visible(True).Build: pCollection.Add Button
  Set checkBox = chkBuilder.name("product_searchbytype").mark(marked).OfPlaceAndSize(Button.Right, Button.top, cs.Standard, cs.Standard).Movable(True).visible(True).enabled(True).Build: pCollection.Add checkBox
  Set Button = btnBuilder.caption("по виду").OfPlaceAndSize(checkBox.Right, Button.top, 50, cs.Standard).Style(ControlStyle.csSimplified).Movable(True).fontSize(fs.Button).visible(True).Build: pCollection.Add Button
  Set checkBox = chkBuilder.name("product_searchbyname").mark(marked).OfPlaceAndSize(Button.Right, Button.top, cs.Standard, cs.Standard).Movable(True).visible(True).enabled(True).Build: pCollection.Add checkBox
  Set Button = btnBuilder.caption("по наименованию").OfPlaceAndSize(checkBox.Right, checkBox.top, 100, cs.Standard).Style(ControlStyle.csSimplified).Movable(True).fontSize(fs.Button).visible(True).Build: pCollection.Add Button
  Call CreateAlphabet(Button.Bottom, True, True)
  Set Button = btnBuilder.name("product_select").mark(marked).OfPlaceAndSize(pCollection.ItemByName("product_dish").Right + 15, rng.top + cs.Middle, cs.Large + 2, cs.Large + 2).Style(csSelected).Movable(False).SetIcon("IconMain").visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_save").mark("F").OfPlaceAndSize(Button.Left, Button.top, Button.Width, Button.height).Style(csSelected).Movable(False).SetIcon("IconSave").visible(False).Build: pCollection.Add Button
  Set Button = btnBuilder _
  .name("product_productadd") _
  .caption("Добавить") _
  .OfPlaceAndSize(Button.Right + 25, Button.top, cs.Wide, cs.Standard) _
  .Style(csOff) _
  .Movable(False) _
  .fontSize(fs.Button) _
  .OnAction("ProductForm") _
  .visible(True) _
  .Build: pCollection.Add Button
  
  Set Button = btnBuilder.name("product_listreturn").caption("Вернуться").OfPlaceAndSize(Button.Left, Button.top, cs.Wide, cs.Standard).Style(csOff).Movable(False).fontSize(fs.Button).OnAction("ProductList").visible(False).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_productgroupadd").caption("Добавить").OfPlaceAndSize(Button.Right + 15, Button.top, cs.Wide, cs.Standard).Style(csOff).fontSize(fs.Button).OnAction("ProductGroupAdd").visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_productaddreturn").caption("Вернуться").OfPlaceAndSize(Button.Left, Button.top, Button.Width, Button.height).Style(csOff).fontSize(fs.Button).OnAction("ProductForm").visible(False).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_productchange").caption("Изменить").OfPlaceAndSize(pCollection.ItemByName("product_productadd").Left, pCollection.ItemByName("product_selectIcon").Bottom - Button.height, Button.Width, Button.height).Style(csOff).fontSize(fs.Button).OnAction("ProductChange").visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_productgroupchange").caption("Изменить").OfPlaceAndSize(pCollection.ItemByName("product_productgroupadd").Left, pCollection.ItemByName("product_productchange").top, Button.Width, Button.height).Style(csOff).fontSize(fs.Button).OnAction("ProductGroupChange").visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_selecttitle").caption("выбрать").OfPlaceAndSize(pCollection.ItemByName("product_selectIcon").Left + pCollection.ItemByName("product_selectIcon").Width / 2 - 70 / 2, pCollection.ItemByName("product_selectIcon").top - cs.Middle, 70, cs.Standard).Alignment(msoAlignCenter).Style(csSimplified).fontSize(fs.Menu).visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_producttitle").caption("продукт").OfPlaceAndSize(pCollection.ItemByName("product_productadd").Left + pCollection.ItemByName("product_productadd").Width / 2 - 56 / 2, pCollection.ItemByName("product_selecttitle").top, 56, cs.Standard).Alignment(msoAlignCenter).Style(csSimplified).fontSize(fs.Menu).visible(True).Build: pCollection.Add Button
  Set Button = btnBuilder.name("product_grouptitle").caption("виды").OfPlaceAndSize(pCollection.ItemByName("product_productgroupadd").Left + pCollection.ItemByName("product_productgroupadd").Width / 2 - 56 / 2, pCollection.ItemByName("product_selecttitle").top, 56, cs.Standard).Alignment(msoAlignCenter).Style(csSimplified).fontSize(fs.Menu).visible(True).Build: pCollection.Add Button
  
  leftEdge = pCollection.ItemByName("product_selectIcon").Left - 10
  rightEdge = pCollection.ItemByName("product_productgroupadd").Right + 3
  topEdge = pCollection.ItemByName("product_selectIcon").top - 6
  bottomEdge = pCollection.ItemByName("product_selectIcon").Bottom + 6
  With New LineBuilder
    .DrawLine this.sheet, leftEdge, bottomEdge, rightEdge, bottomEdge
    .DrawLine this.sheet, leftEdge, topEdge, rightEdge, topEdge
  End With

  Set tbl = tblBuilder.name("product_listproduct").OfPlaceAndSize(pCollection.ItemByName("All_off").BottomRightRange.Offset(2, -1).Resize(3, 6)).Header(titleListProduct).Build
   
  Set rng = this.ProductFormFirstCell.Offset(0, 1)
  
  '====================Дальше работать!!!
  Set Button = btnBuilder.name("product_product").caption("ДОБАВИТЬ ПРОДУКТ").OfPlaceAndSize(rng.Left, rng.top + 5, 300, cs.Double).Style(ControlStyle.csAccented).Movable(True).fontSize(fs.Title).visible(True).Build: pCollection.Add Button
  Set label = lblBuilder.caption("группа (вид):").OfPlaceAndSize(Button.BottomLeftRange.Offset(2, 0)).Style(csSimplified).fontSize(fs.DataLabel).Field(1, 5, , "product_вид", , , True).Build
  Set lblControl = label ': Debug.Print lblControl.CRange.address
  Set label = lblBuilder.caption("код продукта:").OfPlaceAndSize(RangeAtPosition(this.sheet, label.Right, label.top).Offset(0, 1)).Style(csSimplified).fontSize(fs.DataLabel).Alignment(xlRight).Field(1, , , "product_код", , , True).Build
  Set label = lblBuilder.caption("название:").OfPlaceAndSize(lblControl.CRange.Offset(2, 0)).Style(csSimplified).fontSize(fs.DataLabel).Field(1, 8, , "product_продукт", , , True).Build
  Set label = lblBuilder.caption("ед.изм:").OfPlaceAndSize(label.BottomLeftRange.Offset(1, 0)).Style(csSimplified).fontSize(fs.DataLabel).Field(1, , , "product_измерение", , , True).Build
  Set label = lblBuilder.caption("описание:").OfPlaceAndSize(label.BottomLeftRange.Offset(1, 0)).Style(csSimplified).fontSize(fs.DataLabel).Field(1, 8, 2, "product_описание", , , True).Build
  Set label = lblBuilder.caption("характеристики:").OfPlaceAndSize(label.BottomLeftRange.Offset(2, 0)).Style(csUnderlined).fontSize(fs.DataLabel).Build
  Set label = lblBuilder.caption("калорийность (ккал):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).fontSize(fs.DataLabel).Field(2, , , "product_калорийность", , , True).Build
  Set label = lblBuilder.caption("белки (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).fontSize(fs.DataLabel).Field(2, , , "product_белки", , , True).Build
  Set label = lblBuilder.caption("жиры (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).fontSize(fs.DataLabel).Field(2, , , "product_жиры", , , True).Build
  Set label = lblBuilder.caption("углеводы (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).fontSize(fs.DataLabel).Field(2, , , "product_углеводы", , , True).Build

  recount = CountObjects
  
  dataToUpdate = iProduct.Update(product, sheetControls, recount)

  Application.ScreenUpdating = True
  this.sheet.Range("A2").Select
  
  Set lblControl = Nothing
  Set btnBuilder = Nothing
  Set lblBuilder = Nothing
  Set chkBuilder = Nothing
  Set Button = Nothing
  Set label = Nothing
  Set checkBox = Nothing
  Set rng = Nothing
  Exit Sub
FailExit:
  Err.Raise Err.Number, Err.source, Err.description
End Sub

Private Sub CreateAlphabet(ByVal startPosition As Single, Optional ByVal eng As Boolean, Optional ByVal cyr As Boolean)
  Dim litera As String
  Dim firstLine As Single, secondLine As Single
  Dim i As Integer
  Dim abcbutton As IControl
  Dim resetbutton As IControl
  Dim allbutton As IControl
  Dim lPosition As Single
  Dim marked As String
  Dim btnBuilder As ButtonBuilder

  Set btnBuilder = NewButtonBuilder(this.sheet)

  On Error GoTo FailExit
  marked = "L"
  firstLine = startPosition + 4
  Set allbutton = btnBuilder.name("All_off").caption("ВСЕ").OfPlaceAndSize(this.ListFormFirstCell.Left + 3, firstLine, cs.Double, cs.Small).fontSize(fs.ButtonSmall).OnAction("ABC_Click").Style(csOff).Movable(True).visible(True).Build: pCollection.Add allbutton
  Set allbutton = btnBuilder.name("All_on").caption("ВСЕ").OfPlaceAndSize(allbutton.Left, firstLine, cs.Double, cs.Small).fontSize(fs.ButtonSmall).OnAction("ABC_Click").Style(csOn).Movable(True).visible(False).Build: pCollection.Add allbutton
  Set resetbutton = btnBuilder.name("Reset").caption("ОЧИСТИТЬ").OfPlaceAndSize(allbutton.Right + CHKBOXINDENT, firstLine, 70, cs.Small).fontSize(fs.ButtonSmall).OnAction("ResetAlphabet").Style(ControlStyle.csCanceled).Movable(True).visible(True).Build:: pCollection.Add resetbutton
  lPosition = resetbutton.Right + CHKBOXINDENT
  For i = 1 To Len(ABC_CYR)
    litera = Mid$(ABC_CYR, i, 1)
    Set abcbutton = btnBuilder.name(litera & "_off").caption(litera).OfPlaceAndSize(lPosition, firstLine, cs.Small, cs.Small).fontSize(fs.ButtonSmall).OnAction("'ABC_Click " & Chr(34) & litera & Chr(34) & "'").Style(csOff).Movable(True).visible(True).Build: pCollection.Add abcbutton
    Set abcbutton = btnBuilder.name(litera & "_on").caption(litera).OfPlaceAndSize(lPosition, firstLine, cs.Small, cs.Small).fontSize(fs.ButtonSmall).OnAction("ABC_Click").Style(csOn).Movable(True).visible(False).Build: pCollection.Add abcbutton
    lPosition = abcbutton.Right + CHKBOXINDENT
  Next i
  lPosition = resetbutton.Right + CHKBOXINDENT
  secondLine = resetbutton.Bottom + CHKBOXINDENT
  For i = 1 To Len(ABC_ENG)
    litera = Mid$(ABC_ENG, i, 1)
    Set abcbutton = btnBuilder.name(litera & "_off").caption(litera).OfPlaceAndSize(lPosition, secondLine, cs.Small, cs.Small).fontSize(fs.ButtonSmall).OnAction("'ABC_Click " & Chr(34) & litera & Chr(34) & "'").Style(csOff).Movable(True).visible(True).Build: pCollection.Add abcbutton
    Set abcbutton = btnBuilder.name(litera & "_on").caption(litera).OfPlaceAndSize(lPosition, secondLine, cs.Small, cs.Small).fontSize(fs.ButtonSmall).OnAction("ABC_Click").Style(csOn).Movable(True).visible(False).Build: pCollection.Add abcbutton
    lPosition = abcbutton.Right + CHKBOXINDENT
  Next i

  Set abcbutton = Nothing
  Set resetbutton = Nothing
  Set allbutton = Nothing
  Set btnBuilder = Nothing
  Exit Sub
FailExit:
End Sub

Private Function CheckBuild() As Boolean
  Dim par As Variant
  par = iProduct.parameter(Table.product, Options.sheetControls)
  If CLng(par) = CountObjects() And CLng(par) <> 0 Then
    CheckBuild = True
  End If
End Function

Public Function Interface(ByVal iType As InterfaceType) As ProductView
this.sheet.Activate
  With curInterface
    .iView = iType
    .Builded = CheckBuild
  End With
  If Not curInterface.Builded Then
    With New ControlCollection
      Set pCollection = .Self
    End With
    Call FormatSheet
    Call Build(curInterface)
  Else
    'Set pCollection = Nothing                    'for testing
    If pCollection Is Nothing Then
      Set pCollection = RestoreCollection
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
  If Not IsConfigured Then SysInfo.SetSystemParameter
  
  Set iProduct = New MetaData
  Set this.sheet = iProduct.Create(product)
  With this.sheet
    Set this.ListFormFirstCell = .Cells(1, StandardColumn.FirstLevel)
    Set this.ProductFormFirstCell = .Cells(1, StandardColumn.SecondLevel)
  End With
  'ActiveWindow.DisplayHeadings = False
End Sub

Private Sub Class_Terminate()
  Set iProduct = Nothing
  Set this.ListFormFirstCell = Nothing
  Set this.ProductFormFirstCell = Nothing
  Set this.sheet = Nothing
  'ActiveWindow.DisplayHeadings = True
End Sub

Private Sub FormatSheet()
  With this.sheet
    .Cells.Interior.Color = Color.MainBackground
    .Range("A1").EntireColumn.Hidden = True
    .columns(this.ListFormFirstCell.Offset(0, -1).column).ColumnWidth = 2
    .columns(this.ListFormFirstCell.column).ColumnWidth = 5
    .columns(this.ListFormFirstCell.Offset(0, 1).column).ColumnWidth = 7
    .columns(this.ListFormFirstCell.Offset(0, 2).column).ColumnWidth = 30
    .columns(this.ListFormFirstCell.Offset(0, 3).column).ColumnWidth = 75
    .columns(this.ListFormFirstCell.Offset(0, 4).column).ColumnWidth = 12
    .columns(this.ListFormFirstCell.Offset(0, 5).column).ColumnWidth = 12
    .columns(this.ProductFormFirstCell.column).ColumnWidth = 0.1
    .columns(this.ProductFormFirstCell.Offset(0, 1).column).ColumnWidth = 14
    .columns(this.ProductFormFirstCell.Offset(0, 2).column).ColumnWidth = 10
    .Range(this.ProductFormFirstCell.Offset(0, 3).Resize(1, 3).address).ColumnWidth = 7
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

