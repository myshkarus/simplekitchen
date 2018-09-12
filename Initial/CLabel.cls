VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "CLabel"
Private LabelRange As Range

Private Type tControl
  TName As String
  tID As String
  tTabIndex As Long
  tStyle As ControlStyle
  tIcon As String
  tSheet As Worksheet
  tLeft As Single
  tTop As Single
  tMark As String
  
  tCaption As String
  tVisible As Boolean
  tAlignment As Variant
  tFontSize As Long
  tField As Boolean
  tFieldOffset As Integer
  tFieldColumns As Long
  tFieldRows As Long
  tFieldList As Boolean
  tFieldFlag As Boolean
  tFieldBorder As Boolean
  tFieldName As String
  tFieldRange As Range
End Type

Private this As tControl
Implements IControl

Private Property Get IControl_ID() As String
  IControl_ID = this.tID
End Property

Private Property Get IControl_ControlType() As String
  IControl_ControlType = TypeName(Self)
End Property

Private Property Get IControl_Name() As String
  IControl_Name = this.TName
End Property

Private Property Let IControl_Name(ByVal text As String)
  If text <> vbNullString Then
    this.TName = text
  End If
End Property

Private Property Let IControl_Caption(ByVal text As String)
  Caption = text
End Property

Private Property Get IControl_BottomRightRange() As Range
  Set IControl_BottomRightRange = CRange.Offset(1, 0)
End Property

Private Property Get IControl_BottomLeftRange() As Range
  Set IControl_BottomLeftRange = CRange.Offset(1, 0)
End Property

Public Property Let Mark(ByVal text As String)
  this.tMark = text
End Property

Private Property Get IControl_Mark() As String
  IControl_Mark = this.tMark
End Property

Private Property Get IControl_TabIndex() As Long
  TabIndex = this.tTabIndex
End Property

Private Property Let IControl_TabIndex(ByVal value As Long)
  this.tTabIndex = value
End Property

'для CLabel установлено условно
Private Property Get IControl_Visible() As Boolean
  IControl_Visible = this.tVisible
End Property

'для CLabel установлено условно
Private Property Let IControl_Visible(ByVal value As Boolean)
  this.tVisible = value
End Property

Private Property Get IControl_Style() As ControlStyle
  IControl_Style = this.tStyle
End Property

Private Property Let IControl_Style(ByVal value As ControlStyle)
  If Not CRange Is Nothing Then
    Call SetupLabelStyle(value)
  Else:
    this.tStyle = value
  End If
End Property

Private Property Let IControl_Movable(ByVal status As Boolean)
End Property

Private Property Get IControl_Height() As Single
  If Not CRange Is Nothing Then
    IControl_Height = CRange.Height
  End If
End Property

Private Property Get IControl_Width() As Single
  If Not CRange Is Nothing Then
    IControl_Width = CRange.Width
  End If
End Property

Private Property Get IControl_Top() As Single
  If Not CRange Is Nothing Then
    IControl_Top = CRange.top
  End If
End Property

Private Property Get IControl_Left() As Single
  If Not CRange Is Nothing Then
    IControl_Left = CRange.left
  End If
End Property

Private Property Get IControl_Bottom() As Single
  If Not CRange Is Nothing Then
    If this.tField Then
      IControl_Bottom = this.tFieldRange.top + this.tFieldRange.Height
    Else:
    IControl_Bottom = CRange.top + CRange.Height
  End If
End If
End Property

'доработать с Field
Private Property Get IControl_Right() As Single
  If Not CRange Is Nothing Then
    If this.tField Then
      IControl_Right = this.tFieldRange.left + this.tFieldRange.Width
    Else:
    IControl_Right = CRange.left + CRange.Width
  End If
End If
End Property

Public Property Get FontSize() As Long
  If Not CRange Is Nothing Then
    FontSize = CRange.Font.size
  Else: FontSize = this.tFontSize
  End If
End Property

Public Property Let FontSize(ByVal value As Long)
  If Not CRange Is Nothing Then
    If value = 0 Then
      CRange.Font.size = FontSizeEnum.Standard
    Else: CRange.Font.size = value
    End If
  Else: this.tFontSize = value
  End If
End Property

Public Property Get Alignment() As Variant
  'xlCenter, xlLeft, clRight
  If Not CRange Is Nothing Then
    Alignment = CRange.HorizontalAlignment
  Else: Alignment = this.tAlignment
  End If
End Property

Public Property Let Alignment(ByVal val As Variant)
  If Not CRange Is Nothing Then
    If IsEmpty(val) Then
      CRange.HorizontalAlignment = xlLeft
    Else
      CRange.HorizontalAlignment = val
      If val = xlRight Then CRange.IndentLevel = 1
    End If
  Else: this.tAlignment = val
  End If
End Property

Public Property Get Caption() As String
  If Not CRange Is Nothing Then
    Caption = CRange.text
  Else: Caption = this.tCaption
  End If
End Property

Public Property Let Caption(ByVal text As String)
  If Not CRange Is Nothing Then
    CRange.value = text
  Else: this.tCaption = text
  End If
End Property

Private Sub IControl_Create(ByVal wsh As Worksheet)
  Dim rngField As Range
  On Error Resume Next
  If Not wsh Is Nothing Then Set this.tSheet = wsh
  If CRange Is Nothing Then Exit Sub
  
  With CRange
    this.tID = .ID
    If this.tFieldOffset <> 0 Then
      Set rngField = ControlField(this.tFieldOffset, this.tFieldColumns, this.tFieldRows, this.tFieldName, this.tFieldList, this.tFieldFlag, this.tFieldBorder)
    End If
    If Not rngField Is Nothing Then
      If this.TName <> vbNullString And this.tFieldName <> vbNullString Then
        ThisWorkbook.Names.Add Name:=this.TName, RefersTo:=CRange
      End If
    Else:
      ThisWorkbook.Names.Add Name:=this.TName, RefersTo:=CRange
    End If
    Call SetupLabelStyle(this.tStyle)
    Caption = this.tCaption
    FontSize = this.tFontSize
    'Debug.Print .address
    Alignment = this.tAlignment
  End With
End Sub

Public Property Get CRange() As Range
  Set CRange = LabelRange
End Property

Public Property Set CRange(ByVal rng As Range)
  If Not rng Is Nothing Then
    Set LabelRange = rng
  End If
End Property

Public Property Let FieldOffset(ByVal startOffset As Integer)
  this.tFieldOffset = startOffset
End Property

Public Property Let FieldColumns(ByVal columns As Long)
  If columns = 0 Then
    this.tFieldColumns = 1
  Else
    this.tFieldColumns = columns
  End If
End Property

Public Property Let FieldRows(ByVal rows As Long)
  If rows = 0 Then
    this.tFieldRows = 1
  Else
    this.tFieldRows = rows
  End If
End Property

Public Property Let FieldList(ByVal list As Boolean)
  this.tFieldList = list
End Property

Public Property Let FieldFlag(ByVal flag As Boolean)
  this.tFieldFlag = flag
End Property

Public Property Let FieldBorder(ByVal border As Boolean)
  this.tFieldBorder = border
End Property

Public Property Let FieldName(ByVal Name As String)
  this.tFieldName = Name
End Property

Private Function ControlField(ByVal startOffset As Integer, ByVal columns As Long, Optional ByVal rows As Long, _
                              Optional ByVal rangeName As String, Optional ByVal list As Boolean, Optional ByVal flag As Boolean, Optional ByVal border As Boolean) As Range
  Dim rng As Range
  If Not CRange Is Nothing Then
    Set rng = CRange.Offset(0, startOffset)
    If columns > 1 Then
      Set rng = rng.Resize(1, columns)
      rng.WrapText = False
      rng.HorizontalAlignment = xlLeft
      If rows > 1 Then
        Set rng = rng.Resize(rows, columns)
      End If
    End If
    If rng.columns.count > 1 Then
      rng.MergeCells = True
      If rng.rows.count > 1 Then
        rng.WrapText = True
        rng.VerticalAlignment = xlTop
      End If
    End If
    With rng
      .Interior.Color = Color.OffText
      .Font.Color = Color.OffBackground
      .Font.Bold = True
      .Font.Name = "Calibri"
      .Font.size = FontSizeEnum.Medium
      .Locked = False
      .FormulaHidden = False
      If border Then
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = Color.OnBackground
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
    End With
    this.tField = True
    Set this.tFieldRange = rng
    If rangeName <> vbNullString Then
      ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=this.tFieldRange
    Else
      If this.TName <> vbNullString Then
        ThisWorkbook.Names.Add Name:=this.TName, RefersTo:=this.tFieldRange
      End If
    End If
  End If
  Set ControlField = rng
  Set rng = Nothing
End Function

Private Sub SetupLabelStyle(ByVal Style As ControlStyle)
  If Not CRange Is Nothing Then
  With CRange
    .Locked = True
    Select Case Style
    Case csDefault
      .Font.Color = Color.OffBackground: .Font.Name = "Calibri"
      .Font.Bold = True
    Case csSimplified
      .Font.Color = Color.MainText: .Font.Name = "Calibri Light"
    Case csAccented
      .Font.Color = Color.AccentText: .Font.Name = "Calibri Light"
    Case csUnderlined
      .Font.Color = Color.MainText: .Font.Name = "Calibri": .Font.Underline = True
    End Select
  End With
  End If
End Sub

Private Function ToString() As String
  ToString = ClassName & " [id=" & ID & "] " & " to be build"
End Function

Private Sub Class_Initialize()
  this.tField = False
  'Set ButtonShape = ThisWorkbook.Worksheets(TableList).Shapes.AddShape(msoShapeRectangle, 0, 0, size.Std, size.Std)
End Sub

Public Property Get Self() As IControl
  Set Self = Me
End Property

Private Sub Class_Terminate()
  Set LabelRange = Nothing
  Set this.tSheet = Nothing
End Sub

Public Property Get Icon() As String
  Icon = this.tIcon
End Property

Public Property Let Icon(ByVal Name As String)
  this.tIcon = Name
End Property

'Private Function ControlIcon(ByVal iconName As String) As ShapeRange
'  Dim tempName As String
'  If this.tSheet Is Nothing Then
'    MsgBox "CButton.Sheet must be instantiated first!", vbOKOnly Or vbCritical, "Class " & ClassName & " error"
'    Exit Function
'  End If
'  'On Error Resume Next
'  On Error GoTo 0
'  tempName = ButtonShape.name
'  If Len(iconName) <> 0 Then
'    ThisWorkbook.Worksheets(TableList).Shapes(iconName).Copy
'    this.tSheet.Paste
'    With this.tSheet.Shapes(iconName)
'      .Width = ButtonShape.Width * 0.8
'      .Height = .Width
'      .Left = ButtonShape.Left + ButtonShape.Width / 2 - .Width / 2
'      .Top = ButtonShape.Top + ButtonShape.Height / 2 - .Height / 2
'    End With
'    this.tSheet.Shapes.Range(Array(ButtonShape.name, iconName)).Select
'    Selection.ShapeRange.Group.Select
'    'тут что-нибудь придумать (чтобы присвоить shaperange id)
'    IControl_Name = tempName & "Icon"
'    Selection.name = tempName
'    Selection.Placement = xlFreeFloating
'    Set ControlIcon = Selection
'  End If
'End Function

