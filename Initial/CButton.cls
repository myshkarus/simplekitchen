VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "CButton"
Private ButtonShape As Shape
Private ButtonIconShape As ShapeRange

Private Type tControl
  tID As String
  tTabIndex As Long                              'last Characters of AlternativeText
  tStyle As ControlStyle                         'do not restore!
  tIcon As String                                'do not restore!
  tSheet As Worksheet
  tLeft As Single
  tTop As Single
  tMark As String                                'first Character of AlternativeText
  TName As String
  tVisible As Boolean
  tMovable As Boolean
  tFontSize As Long
  tAlignment As Variant
  tOnAction As String
  tCaption As String
  tHeight As Single
  tWidth As Single
End Type

Private this As tControl
Implements IControl

Public Property Set ShapeBtn(ByRef sh As Shape)
  If Not sh Is Nothing Then
    If Not sh.parent Is Nothing Then
      If TypeName(sh.parent) = "GroupShapes" Then
        Set ButtonIconShape = ShapeIconBtn(sh.ParentGroup.Name)
      End If
    End If
    If ButtonShape Is Nothing Then Set ButtonShape = sh
    With ButtonShape
      this.tTabIndex = Right$(.AlternativeText, IIf(Len(.AlternativeText) > 1, Len(.AlternativeText) - 1, Len(.AlternativeText))) ' ?? len=1
      this.tLeft = .left
      this.tTop = .top
      this.tMark = left$(.AlternativeText, IIf(Len(.AlternativeText) > 1, 1, 0)) ' ?? len=0
    End With
  End If
End Property

Private Function ShapeIconBtn(ByVal text As String) As ShapeRange
  If text <> vbNullString Then
    If Not this.tSheet Is Nothing Then Set ShapeIconBtn = this.tSheet.Shapes.Range(Array(text))
  End If
End Function

Public Property Set sheet(ByVal wsh As Worksheet)
  If Not wsh Is Nothing Then
    Set this.tSheet = wsh
  End If
End Property

Private Property Get IControl_ID() As String
  IControl_ID = this.tID
End Property

Private Property Get IControl_ControlType() As String
  IControl_ControlType = TypeName(Self)
End Property

Private Property Get IControl_Name() As String
  If Not ButtonShape Is Nothing Then IControl_Name = ButtonShape.Name
End Property

Private Property Let IControl_Name(ByVal text As String)
  this.TName = text
  If this.TName <> vbNullString And Not ButtonShape Is Nothing Then
    ButtonShape.Name = this.TName
  End If
End Property

Private Property Let IControl_Caption(ByVal text As String)
  Caption = text
End Property

Private Property Get IControl_Mark() As String
  If Not ButtonShape Is Nothing Then
    With ButtonShape
      IControl_Mark = left$(.AlternativeText, IIf(Len(.AlternativeText) > 1, 1, 0))
    End With
  End If
End Property

Private Property Get IControl_TabIndex() As Long
  If Not ButtonShape Is Nothing Then
    With ButtonShape
      IControl_TabIndex = Right$(.AlternativeText, IIf(Len(.AlternativeText) > 1, Len(.AlternativeText) - 1, Len(.AlternativeText)))
    End With
  End If
End Property

Private Property Let IControl_TabIndex(ByVal value As Long)
  this.tTabIndex = value
End Property

Private Property Get IControl_Visible() As Boolean
  If Not ButtonIconShape Is Nothing Then
    IControl_Visible = ButtonIconShape.Visible
  ElseIf Not ButtonShape Is Nothing Then
    IControl_Visible = ButtonShape.Visible
  Else
    IControl_Visible = this.tVisible
  End If
End Property

Private Property Let IControl_Visible(ByVal value As Boolean)
  If Not ButtonIconShape Is Nothing Then
    ButtonIconShape.Visible = value
  Else
    If ButtonShape Is Nothing Then
      this.tVisible = value
    Else: ButtonShape.Visible = value
    End If
  End If
End Property

Private Property Let IControl_Movable(ByVal status As Boolean)
  this.tMovable = status
  If Not ButtonShape Is Nothing Then
    With ButtonShape
      If status Then
        .Placement = xlMoveAndSize
      Else: .Placement = xlFreeFloating
      End If
    End With
  End If
End Property

Private Property Let IControl_Style(ByVal value As ControlStyle)
  this.tStyle = value
  If Not ButtonShape Is Nothing Then
    SetupButtonStyle (value)
  End If
End Property

Private Property Get IControl_Height() As Single
  If Not ButtonShape Is Nothing Then IControl_Height = ButtonShape.Height
End Property

Private Property Get IControl_Width() As Single
  If Not ButtonShape Is Nothing Then IControl_Width = ButtonShape.Width
End Property

Private Property Get IControl_Top() As Single
  IControl_Top = this.tTop
End Property

Private Property Get IControl_Left() As Single
  IControl_Left = this.tLeft
End Property

Private Property Get IControl_Bottom() As Single
  If Not ButtonShape Is Nothing Then IControl_Bottom = this.tTop + ButtonShape.Height
End Property

Private Property Get IControl_Right() As Single
  If Not ButtonShape Is Nothing Then IControl_Right = this.tLeft + ButtonShape.Width
End Property

Private Property Get IControl_BottomRightRange() As Range
  If Not ButtonShape Is Nothing Then Set IControl_BottomRightRange = ButtonShape.BottomRightCell
End Property

Private Property Get IControl_BottomLeftRange() As Range
  If Not ButtonShape Is Nothing Then Set IControl_BottomLeftRange = ButtonShape.TopLeftCell.Offset(1, 0)
End Property

Public Property Get FontSize() As Long
  If Not ButtonShape Is Nothing Then FontSize = ButtonShape.TextFrame2.TextRange.Font.size
End Property

Public Property Let FontSize(ByVal value As Long)
  this.tFontSize = value
End Property

Public Property Get Alignment() As MsoParagraphAlignment
  If Not ButtonShape Is Nothing Then Alignment = ButtonShape.TextFrame2.TextRange.ParagraphFormat.Alignment
End Property

Public Property Let Alignment(ByVal val As MsoParagraphAlignment)
  this.tAlignment = val
End Property

Public Property Let Mark(ByVal text As String)
  this.tMark = text
End Property

Public Property Get OnAction() As String
  If Not ButtonShape Is Nothing Then OnAction = CStr(ButtonShape.OnAction)
End Property

Public Property Let OnAction(ByVal val As String)
  this.tOnAction = val
End Property

Public Property Get Caption() As String
  If Not ButtonShape Is Nothing Then Caption = ButtonShape.TextFrame2.TextRange.Characters.text
End Property

Public Property Let Caption(ByVal text As String)
  this.tCaption = text
  If Not ButtonShape Is Nothing Then
    ButtonShape.TextFrame2.TextRange.Characters.text = text
  End If
End Property

Public Property Let Height(ByVal value As Single)
  this.tHeight = value
End Property

Public Property Let Width(ByVal value As Single)
  this.tWidth = value
End Property

Public Property Let top(ByVal value As Single)
  this.tTop = value
End Property

Public Property Let CLeft(ByVal value As Single)
  this.tLeft = value
End Property

Private Function ToString() As String
  ToString = ClassName & " [id=" & ID & "] " & " to be build"
End Function

Public Property Get Self() As IControl
  Set Self = Me
End Property

Private Sub Class_Terminate()
  Set ButtonShape = Nothing
  Set this.tSheet = Nothing
End Sub

Private Sub SetupButtonStyle(ByVal Style As ControlStyle)
  With ButtonShape
    .Fill.Transparency = 0
    .Placement = xlMoveAndSize
    .TextFrame2.TextRange.Font.Bold = msoTrue
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    .Locked = msoFalse
    .LockAspectRatio = msoTrue
    Select Case Style
    Case csOff
      .Fill.ForeColor.RGB = Color.OffBackground
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.OffText
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .Shadow.Visible = msoFalse
    Case csOn
      .Fill.ForeColor.RGB = Color.OnBackground
      .line.ForeColor.RGB = Color.MainText
      .Shadow.Type = msoShadow41
      .Shadow.Blur = 3
      .Shadow.OffsetX = 0
      .Shadow.OffsetY = 4
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.OnText
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Case csCanceled
      .Fill.ForeColor.RGB = Color.OffText
      .line.ForeColor.RGB = Color.CancelLine
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.CancelText
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Case csSelected
      .Fill.ForeColor.RGB = Color.SelectBackground
      .line.ForeColor.RGB = Color.SelectLine
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Case csAccented
      .Fill.Visible = msoFalse
      .line.Visible = msoFalse
      .TextFrame2.TextRange.Font.Bold = msoFalse
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.AccentText
      .TextFrame2.TextRange.Font.Name = "Calibri Light"
      IIf Alignment <> msoAlignCenter, .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft, .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .TextFrame2.MarginLeft = 3
      .Locked = msoTrue
    Case csSimplified
      .Fill.Visible = msoFalse
      .line.Visible = msoFalse
      .TextFrame2.TextRange.Font.Bold = msoFalse
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.MainText
      .TextFrame2.TextRange.Font.Name = "Calibri Light"
      IIf Alignment <> msoAlignCenter, .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft, .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .TextFrame2.MarginLeft = 3
      .Locked = msoTrue
    Case csUnderlined
      .Fill.Visible = msoFalse
      .line.Visible = msoFalse
      .TextFrame2.TextRange.Font.Bold = msoFalse
      .TextFrame2.TextRange.Font.UnderlineStyle = msoUnderlineSingleLine
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.MainText
      .TextFrame2.TextRange.Font.Name = "Calibri"
      .TextFrame2.TextRange.ParagraphFormat.WordWrap = msoTrue
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
      .TextFrame2.VerticalAnchor = msoAnchorTop
      .TextFrame2.MarginLeft = 3
      .TextFrame2.MarginTop = 0
      .Locked = msoTrue
    Case csAlternative
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Color.OffText
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .Fill.Visible = msoFalse
      .line.Visible = msoTrue
      .line.BackColor.RGB = Color.OffBackground
      .Shadow.Visible = msoFalse
    End Select
  End With
End Sub

Public Property Get Icon() As String
  Icon = this.tIcon
End Property

Public Property Let Icon(ByVal Name As String)
  this.tIcon = Name
End Property

Private Function ControlIcon(ByVal iconName As String) As ShapeRange
  Dim tempName As String
  Dim newIconName As String
  On Error GoTo 0
  If this.tSheet Is Nothing Then
    MsgBox "CButton.Sheet must be instantiated first!", vbOKOnly Or vbCritical, "Class " & ClassName & " error"
    Exit Function
  End If
  tempName = ButtonShape.Name: IControl_Name = tempName & "Icon"
  If Len(iconName) <> 0 Then
    ThisWorkbook.Worksheets(TableList).Shapes(iconName).Copy
    this.tSheet.Paste
    newIconName = iconName & Rnd(1000)
    this.tSheet.Shapes(iconName).Name = newIconName
    With this.tSheet.Shapes(newIconName)
      .Width = ButtonShape.Width * 0.8
      .Height = .Width
      .left = ButtonShape.left + ButtonShape.Width / 2 - .Width / 2
      .top = ButtonShape.top + ButtonShape.Height / 2 - .Height / 2
      .Placement = xlFreeFloating
    End With
    this.tSheet.Shapes.Range(Array(ButtonShape.Name, newIconName)).Select
    Selection.ShapeRange.Group.Select
    With Selection
      .Name = tempName
      If this.tMovable Then
        .Placement = xlMoveAndSize
      Else: .Placement = xlFreeFloating
      End If
    End With
    Set ControlIcon = this.tSheet.Shapes.Range(Array(tempName))
  End If
End Function

Private Sub IControl_Create(ByVal wsh As Worksheet)
  On Error Resume Next
  If Not wsh Is Nothing Then Set this.tSheet = wsh
  Set ButtonShape = ThisWorkbook.Worksheets(TableList).Shapes.AddShape(msoShapeRectangle, 0, 0, size.std, size.std)
  With ButtonShape
    If Not IsEmpty(this.tAlignment) Then .TextFrame2.TextRange.ParagraphFormat.Alignment = this.tAlignment
    If this.tHeight = 0 Then
      .Height = size.std
    Else: .Height = this.tHeight
    End If
    If this.tWidth = 0 Then
      .Width = size.std
    Else: .Width = this.tWidth
    End If
    Call SetupButtonStyle(this.tStyle)
    If this.tMovable Then
      .Placement = xlMoveAndSize
    Else: .Placement = xlFreeFloating
    End If
    .TextFrame2.TextRange.Characters.text = this.tCaption
    If this.tFontSize = 0 Then
      .TextFrame2.TextRange.Font.size = FontSizeEnum.Standard
    Else: .TextFrame2.TextRange.Font.size = this.tFontSize
    End If
    If this.tOnAction <> vbNullString Then .OnAction = this.tOnAction
    .AlternativeText = left$(this.tMark, 1) & CStr(this.tTabIndex)
    If this.TName <> vbNullString Then .Name = this.TName
  End With
  ButtonShape.Cut
  this.tSheet.Paste
  Set ButtonShape = this.tSheet.Shapes(Selection.Name)
  With ButtonShape
    .left = this.tLeft
    .top = this.tTop
    If Icon <> vbNullString Then
      Set ButtonIconShape = ControlIcon(Icon)
    End If
    IControl_Visible = this.tVisible
  End With
End Sub

Private Sub Class_Initialize()
  this.tID = Hash(TypeName(Me) & CLng(Rnd * (2 ^ 31)))
End Sub

