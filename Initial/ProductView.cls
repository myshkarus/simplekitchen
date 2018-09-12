VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductView"
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
        If .Mark = "L" Then
          .Visible = True
        ElseIf .Mark = "F" Then
          .Visible = False
        End If
        If .Name = "product_selecttitle" Then .Caption = "выбрать"
        If .Name = "product_producttitle" Then .Caption = "продукт"
        If .Name = "product_productchange" Then .Visible = True
        If .Name = "product_productgroupchange" Then .Visible = True
        If .Name = "product_listreturn" Then .Visible = False
        If .Name = "product_productadd" Then .Visible = True
      End With
    Next ctrl
    this.sheet.Range("A1").Select
  Case InterfaceType.ProductForm
    this.sheet.Range(this.ProductFormFirstCell, this.ProductFormFirstCell.Offset(0, 30)).EntireColumn.Hidden = False
    this.sheet.Range(this.ListFormFirstCell, this.ProductFormFirstCell.Offset(0, -1)).EntireColumn.Hidden = True
    For Each ctrl In pCollection
      With ctrl
        If .Mark = "L" Then
          .Visible = False
        ElseIf .Mark = "F" Then
          .Visible = True
        End If
        If .Name = "product_product" Then .Caption = "ДОБАВИТЬ ПРОДУКТ"
        
        If .Name = "product_selecttitle" Then .Caption = "сохранить"
        If .Name = "product_producttitle" Then .Caption = "к списку"
        If .Name = "product_productchange" Then .Visible = False
        If .Name = "product_productgroupchange" Then .Visible = False
        If .Name = "product_listreturn" Then .Visible = True
        If .Name = "product_productadd" Then .Visible = False
      End With
    Next ctrl
    this.sheet.Range("product_вид").Select
    
  Case InterfaceType.ProductChangeForm
    this.sheet.Range(this.ProductFormFirstCell, this.ProductFormFirstCell.Offset(0, 30)).EntireColumn.Hidden = False
    this.sheet.Range(this.ListFormFirstCell, this.ProductFormFirstCell.Offset(0, -1)).EntireColumn.Hidden = True
    For Each ctrl In pCollection
      With ctrl
        If .Mark = "L" Then
          .Visible = False
        ElseIf .Mark = "F" Then
          .Visible = True
        End If

        If .Name = "product_product" Then .Caption = "ИЗМЕНИТЬ ПРОДУКТ"
        If .Name = "product_selecttitle" Then .Caption = "сохранить"
        If .Name = "product_producttitle" Then .Caption = "к списку"
        If .Name = "product_productchange" Then .Visible = False
        If .Name = "product_productgroupchange" Then .Visible = False
        If .Name = "product_listreturn" Then .Visible = True
        If .Name = "product_productadd" Then .Visible = False
      End With
    Next ctrl
    this.sheet.Range("product_вид").Select
    
    
  End Select
  Application.ScreenUpdating = True
End Sub

Private Sub Build(view As TView)
  Dim button As IControl
  Dim label As IControl
  Dim checkbox As IControl
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
  Set button = btnBuilder.Caption("ПОДОБРАТЬ ПРОДУКТЫ").OfPlaceAndSize(rng.left, rng.top + 5, 300, size.Dbl).Style(ControlStyle.csAccented).Movable(True).FontSize(FontSizeEnum.TitleLarge).Visible(True).Build:  pCollection.Add button
  Set button = btnBuilder.Caption("для блюда").OfPlaceAndSize(rng.left, button.Bottom, 90, size.Mdl).Style(ControlStyle.csAccented).Movable(True).FontSize(FontSizeEnum.TitleSmall).Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_dish").Caption("Cуп овощной ").OfPlaceAndSize(button.Right, button.top - 3, 450, size.Trp).Style(ControlStyle.csUnderlined).Movable(True).FontSize(FontSizeEnum.TitleStandard).Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Caption("Критерий поиска:").OfPlaceAndSize(rng.left, button.Bottom, 100, size.std).Style(ControlStyle.csSimplified).Movable(True).FontSize(FontSizeEnum.Medium).Visible(True).Build: pCollection.Add button
  Set checkbox = chkBuilder.Name("product_searchbytype").Mark(marked).OfPlaceAndSize(button.Right, button.top, size.std, size.std).Movable(True).Visible(True).Enabled(True).Build: pCollection.Add checkbox
  Set button = btnBuilder.Caption("по виду").OfPlaceAndSize(checkbox.Right, button.top, 50, size.std).Style(ControlStyle.csSimplified).Movable(True).FontSize(FontSizeEnum.Medium).Visible(True).Build: pCollection.Add button
  Set checkbox = chkBuilder.Name("product_searchbyname").Mark(marked).OfPlaceAndSize(button.Right, button.top, size.std, size.std).Movable(True).Visible(True).Enabled(True).Build: pCollection.Add checkbox
  Set button = btnBuilder.Caption("по наименованию").OfPlaceAndSize(checkbox.Right, checkbox.top, 100, size.std).Style(ControlStyle.csSimplified).Movable(True).FontSize(FontSizeEnum.Medium).Visible(True).Build: pCollection.Add button
  Call CreateAlphabet(button.Bottom, True, True)
  Set button = btnBuilder.Name("product_select").Mark(marked).OfPlaceAndSize(pCollection.ItemByName("product_dish").Right + 13, rng.top + FontSizeEnum.TitleLarge + 5, size.Dbl + 4, size.Dbl + 4).Style(csSelected).Movable(False).SetIcon("IconMain").Visible(True).Build: pCollection.Add button
  'Debug.Print button.Mark
  Set button = btnBuilder.Name("product_save").Mark("F").OfPlaceAndSize(button.left, button.top, button.Width, button.Height).Style(csSelected).Movable(False).SetIcon("IconSave").Visible(False).Build: pCollection.Add button
  'Debug.Print button.Mark
  Set button = btnBuilder.Name("product_productadd").Caption("Добавить").OfPlaceAndSize(button.Right + 25, button.top, size.Dbl * 2, size.Mdl).Style(csOff).Movable(False).FontSize(FontSizeEnum.Medium).OnAction("ProductForm").Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_listreturn").Caption("Вернуться").OfPlaceAndSize(button.left, button.top, size.Dbl * 2, size.Mdl).Style(csOff).Movable(False).FontSize(FontSizeEnum.Medium).OnAction("ProductList").Visible(False).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_productgroupadd").Caption("Добавить").OfPlaceAndSize(button.Right + 15, button.top, size.Dbl * 2, size.Mdl).Style(csOff).FontSize(FontSizeEnum.Medium).OnAction("ProductGroupAdd").Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_productaddreturn").Caption("Вернуться").OfPlaceAndSize(button.left, button.top, button.Width, button.Height).Style(csOff).FontSize(FontSizeEnum.Medium).OnAction("ProductForm").Visible(False).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_productchange").Caption("Изменить").OfPlaceAndSize(pCollection.ItemByName("product_productadd").left, pCollection.ItemByName("product_selectIcon").Bottom - button.Height, button.Width, button.Height).Style(csOff).FontSize(FontSizeEnum.Medium).OnAction("ProductChange").Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_productgroupchange").Caption("Изменить").OfPlaceAndSize(pCollection.ItemByName("product_productgroupadd").left, pCollection.ItemByName("product_productchange").top, button.Width, button.Height).Style(csOff).FontSize(FontSizeEnum.Medium).OnAction("ProductGroupChange").Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_selecttitle").Caption("выбрать").OfPlaceAndSize(pCollection.ItemByName("product_selectIcon").left + pCollection.ItemByName("product_selectIcon").Width / 2 - 70 / 2, pCollection.ItemByName("product_selectIcon").top - size.Mdl - 10, 70, size.Mdl).Alignment(msoAlignCenter).Style(csSimplified).FontSize(FontSizeEnum.MediumLarge).Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_producttitle").Caption("продукт").OfPlaceAndSize(pCollection.ItemByName("product_productadd").left + pCollection.ItemByName("product_productadd").Width / 2 - 56 / 2, pCollection.ItemByName("product_selecttitle").top, 56, size.Mdl).Alignment(msoAlignCenter).Style(csSimplified).FontSize(FontSizeEnum.MediumLarge).Visible(True).Build: pCollection.Add button
  Set button = btnBuilder.Name("product_grouptitle").Caption("виды").OfPlaceAndSize(pCollection.ItemByName("product_productgroupadd").left + pCollection.ItemByName("product_productgroupadd").Width / 2 - 56 / 2, pCollection.ItemByName("product_selecttitle").top, 56, size.Mdl).Alignment(msoAlignCenter).Style(csSimplified).FontSize(FontSizeEnum.MediumLarge).Visible(True).Build: pCollection.Add button
  

  leftEdge = pCollection.ItemByName("product_selectIcon").left - 10
  rightEdge = pCollection.ItemByName("product_productgroupadd").Right + 3
  topEdge = pCollection.ItemByName("product_selectIcon").top - 6
  bottomEdge = pCollection.ItemByName("product_selectIcon").Bottom + 6
  With New LineBuilder
    .DrawLine this.sheet, leftEdge, bottomEdge, rightEdge, bottomEdge
    .DrawLine this.sheet, leftEdge, topEdge, rightEdge, topEdge
  End With

  Set tbl = tblBuilder.Name("product_listproduct").OfPlaceAndSize(pCollection.ItemByName("All_off").BottomRightRange.Offset(2, -1).Resize(3, 6)).Header(titleListProduct).Build
  
  
  Set rng = this.ProductFormFirstCell.Offset(0, 1)
  Set button = btnBuilder.Name("product_product").Caption("ДОБАВИТЬ ПРОДУКТ").OfPlaceAndSize(rng.left, rng.top + 5, 300, size.Dbl).Style(ControlStyle.csAccented).Movable(True).FontSize(FontSizeEnum.TitleLarge).Visible(True).Build: pCollection.Add button
  Set label = lblBuilder.Caption("группа (вид):").OfPlaceAndSize(button.BottomLeftRange.Offset(2, 0)).Style(csSimplified).FontSize(13).Field(1, 5, , "product_вид", , , True).Build
  Set lblControl = label ': Debug.Print lblControl.CRange.address
  Set label = lblBuilder.Caption("код продукта:").OfPlaceAndSize(RangeAtPosition(this.sheet, label.Right, label.top).Offset(0, 1)).Style(csSimplified).FontSize(13).Alignment(xlRight).Field(1, , , "product_код", , , True).Build
  Set label = lblBuilder.Caption("название:").OfPlaceAndSize(lblControl.CRange.Offset(2, 0)).Style(csSimplified).FontSize(13).Field(1, 8, , "product_продукт", , , True).Build
  Set label = lblBuilder.Caption("ед.изм:").OfPlaceAndSize(label.BottomLeftRange.Offset(1, 0)).Style(csSimplified).FontSize(13).Field(1, , , "product_измерение", , , True).Build
  Set label = lblBuilder.Caption("описание:").OfPlaceAndSize(label.BottomLeftRange.Offset(1, 0)).Style(csSimplified).FontSize(13).Field(1, 8, 2, "product_описание", , , True).Build
  Set label = lblBuilder.Caption("характеристики:").OfPlaceAndSize(label.BottomLeftRange.Offset(2, 0)).Style(csUnderlined).FontSize(13).Build
  Set label = lblBuilder.Caption("калорийность (ккал):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).FontSize(13).Field(2, , , "product_калорийность", , , True).Build
  Set label = lblBuilder.Caption("белки (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).FontSize(13).Field(2, , , "product_белки", , , True).Build
  Set label = lblBuilder.Caption("жиры (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).FontSize(13).Field(2, , , "product_жиры", , , True).Build
  Set label = lblBuilder.Caption("углеводы (г/100г):").OfPlaceAndSize(label.BottomLeftRange).Style(csSimplified).FontSize(13).Field(2, , , "product_углеводы", , , True).Build

  recount = CountObjects
  dataToUpdate = iProduct.Update(product, sheetControls, recount)

  Application.ScreenUpdating = True
  this.sheet.Range("A2").Select
  
  Set lblControl = Nothing
  Set btnBuilder = Nothing
  Set lblBuilder = Nothing
  Set chkBuilder = Nothing
  Set button = Nothing
  Set label = Nothing
  Set checkbox = Nothing
  Set rng = Nothing
  Exit Sub
FailExit:
  Err.Raise Err.Number, Err.Source, Err.Description
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
  Set allbutton = btnBuilder.Name("All_off").Caption("ВСЕ").OfPlaceAndSize(this.ListFormFirstCell.left + 3, firstLine, size.Dbl, size.std).OnAction("ABC_Click").Style(csOff).Movable(True).Visible(True).Build: pCollection.Add allbutton
  Set allbutton = btnBuilder.Name("All_on").Caption("ВСЕ").OfPlaceAndSize(allbutton.left, firstLine, size.Dbl, size.std).OnAction("ABC_Click").Style(csOn).Movable(True).Visible(False).Build: pCollection.Add allbutton
  Set resetbutton = btnBuilder.Name("Reset").Caption("ОЧИСТИТЬ").OfPlaceAndSize(allbutton.Right + CHKBOXINDENT, firstLine, 70, size.std).OnAction("ResetAlphabet").Style(ControlStyle.csCanceled).Movable(True).Visible(True).Build:: pCollection.Add resetbutton
  lPosition = resetbutton.Right + CHKBOXINDENT
  For i = 1 To Len(ABC_CYR)
    litera = Mid$(ABC_CYR, i, 1)
    Set abcbutton = btnBuilder.Name(litera & "_off").Caption(litera).OfPlaceAndSize(lPosition, firstLine, size.std, size.std).OnAction("'ABC_Click " & Chr(34) & litera & Chr(34) & "'").Style(csOff).Movable(True).Visible(True).Build: pCollection.Add abcbutton
    Set abcbutton = btnBuilder.Name(litera & "_on").Caption(litera).OfPlaceAndSize(lPosition, firstLine, size.std, size.std).OnAction("ABC_Click").Style(csOn).Movable(True).Visible(False).Build: pCollection.Add abcbutton
    lPosition = abcbutton.Right + CHKBOXINDENT
  Next i
  lPosition = resetbutton.Right + CHKBOXINDENT
  secondLine = resetbutton.Bottom + CHKBOXINDENT
  For i = 1 To Len(ABC_ENG)
    litera = Mid$(ABC_ENG, i, 1)
    Set abcbutton = btnBuilder.Name(litera & "_off").Caption(litera).OfPlaceAndSize(lPosition, secondLine, size.std, size.std).OnAction("'ABC_Click " & Chr(34) & litera & Chr(34) & "'").Style(csOff).Movable(True).Visible(True).Build: pCollection.Add abcbutton
    Set abcbutton = btnBuilder.Name(litera & "_on").Caption(litera).OfPlaceAndSize(lPosition, secondLine, size.std, size.std).OnAction("ABC_Click").Style(csOn).Movable(True).Visible(False).Build: pCollection.Add abcbutton
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
  par = iProduct.parameter(table.product, Options.sheetControls)
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
  ToString = this.sheet.Name
End Function

Private Sub Class_Initialize()
  Set iProduct = New MetaData
  Set this.sheet = iProduct.Create(product)
  With this.sheet
    Set this.ListFormFirstCell = .Cells(1, StandardColumn.FirstLevel)
    Set this.ProductFormFirstCell = .Cells(1, StandardColumn.SecondLevel)
  End With
  ActiveWindow.DisplayHeadings = False
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

