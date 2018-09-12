VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'locked
Private Const ClassName As String = "CCheckBox"
Private CheckBoxShape As checkbox

Private Type tControl
  tID As String
  tTabIndex As Long                              'last Characters of AlternativeText
  tStyle As ControlStyle                         'do not restore!
  tSheet As Worksheet
  tLeft As Single
  tTop As Single
  tMark As String                                'first Character of AlternativeText
  TName As String
  tVisible As Boolean
  tMovable As Boolean
  tCaption As String
  tHeight As Single
  tWidth As Single
  tEnabled As Boolean
  tValue As Boolean
  tLinkedCell As String
End Type

Private this As tControl
Implements IControl

Public Property Set ShapeChkBox(ByVal ck As checkbox)
  If Not ck Is Nothing Then
    If CheckBoxShape Is Nothing Then Set CheckBoxShape = ck
    With CheckBoxShape
      'this.tTabIndex = Right$(.ShapeRange.AlternativeText, IIf(Len(.ShapeRange.AlternativeText) > 1, Len(.ShapeRange.AlternativeText) - 1, Len(.ShapeRange.AlternativeText))) ' ?? len=1
      this.tLeft = .left
      this.tTop = .top
      'this.tMark = Left$(.ShapeRange.AlternativeText, IIf(Len(.ShapeRange.AlternativeText) > 1, 1, 0)) ' ?? len=0
    End With
  End If
End Property

Private Property Get IControl_ID() As String
  IControl_ID = this.tID
End Property

Private Property Get IControl_ControlType() As String
  IControl_ControlType = TypeName(Self)
End Property

Private Property Get IControl_Name() As String
  If Not CheckBoxShape Is Nothing Then IControl_Name = CheckBoxShape.Name
End Property

Private Property Let IControl_Name(ByVal text As String)
  this.TName = text
End Property

Private Property Let IControl_Caption(ByVal text As String)
  Caption = text
End Property

Private Property Get IControl_BottomRightRange() As Range
  If Not CheckBoxShape Is Nothing Then Set IControl_BottomRightRange = CheckBoxShape.BottomRightCell
End Property

Private Property Get IControl_BottomLeftRange() As Range
  If Not CheckBoxShape Is Nothing Then Set IControl_BottomLeftRange = CheckBoxShape.TopLeftCell.Offset(1, 0)
End Property

Private Property Get IControl_Mark() As String
  If Not CheckBoxShape Is Nothing Then
    With CheckBoxShape
      IControl_Mark = left$(.ShapeRange.AlternativeText, IIf(Len(.ShapeRange.AlternativeText) > 1, 1, 0))
    End With
  End If
End Property

Private Property Get IControl_TabIndex() As Long
  If Not CheckBoxShape Is Nothing Then
    With CheckBoxShape
      IControl_TabIndex = Right$(.ShapeRange.AlternativeText, IIf(Len(.ShapeRange.AlternativeText) > 1, Len(.ShapeRange.AlternativeText) - 1, Len(.ShapeRange.AlternativeText)))
    End With
  End If
End Property

Private Property Let IControl_TabIndex(ByVal value As Long)
  this.tTabIndex = value
End Property

Private Property Get IControl_Visible() As Boolean
  If Not CheckBoxShape Is Nothing Then IControl_Visible = CheckBoxShape.Visible
End Property

Private Property Let IControl_Visible(ByVal value As Boolean)
  If CheckBoxShape Is Nothing Then
    this.tVisible = value
  Else: CheckBoxShape.Visible = value
  End If
End Property

Private Property Let IControl_Style(ByVal value As ControlStyle)
  this.tStyle = value
End Property

Private Property Let IControl_Movable(ByVal status As Boolean)
  this.tMovable = status
End Property

Private Property Get IControl_Height() As Single
  If Not CheckBoxShape Is Nothing Then IControl_Height = CheckBoxShape.Height
End Property

Private Property Get IControl_Width() As Single
  If Not CheckBoxShape Is Nothing Then IControl_Width = CheckBoxShape.Width
End Property

Private Property Get IControl_Top() As Single
  IControl_Top = this.tTop
End Property

Private Property Get IControl_Left() As Single
  IControl_Left = this.tLeft
End Property

Private Property Get IControl_Bottom() As Single
  If Not CheckBoxShape Is Nothing Then IControl_Bottom = this.tTop + CheckBoxShape.Height
End Property

Private Property Get IControl_Right() As Single
  If Not CheckBoxShape Is Nothing Then IControl_Right = this.tLeft + CheckBoxShape.Width
End Property

Public Property Let Mark(ByVal text As String)
  this.tMark = text
End Property

Public Property Get Caption() As String
  If Not CheckBoxShape Is Nothing Then Caption = CheckBoxShape.Caption
End Property

Public Property Let Caption(ByVal text As String)
  this.tCaption = text
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

Public Property Get value() As Boolean
  If Not CheckBoxShape Is Nothing Then value = CheckBoxShape.value
End Property

Public Property Let value(ByVal val As Boolean)
  this.tValue = val
End Property

Public Property Get Enabled() As Boolean
  If Not CheckBoxShape Is Nothing Then Enabled = CheckBoxShape.Enabled
End Property

Public Property Let Enabled(ByVal val As Boolean)
  this.tEnabled = val
End Property

Public Property Get LinkedCell() As String
  If Not CheckBoxShape Is Nothing Then LinkedCell = CheckBoxShape.LinkedCell
End Property

Public Property Let LinkedCell(ByVal address As String)
  this.tLinkedCell = address
End Property

Private Function ToString() As String
  ToString = ClassName & " [id=" & ID & "] " & " to be build"
End Function

Public Property Get Self() As IControl
  Set Self = Me
End Property

Private Sub Class_Terminate()
  Set CheckBoxShape = Nothing
  Set this.tSheet = Nothing
End Sub

Private Sub SetupCheckBoxStyle(ByVal Style As ControlStyle)
  With CheckBoxShape
    .Display3DShading = False
    .Placement = xlMoveAndSize
    .Locked = False
    .LockedText = True
  End With
End Sub

Private Sub IControl_Create(ByVal wsh As Worksheet)
  On Error Resume Next
  If Not wsh Is Nothing Then Set this.tSheet = wsh
  Set CheckBoxShape = ThisWorkbook.Worksheets(TableList).CheckBoxes.Add(0, 0, size.std, size.std)
  With CheckBoxShape
    If this.tHeight = 0 Then
      .Height = size.std
    Else: .Height = this.tHeight
    End If
    If this.tWidth = 0 Then
      .Width = size.std
    Else: .Width = this.tWidth
    End If
    Call SetupCheckBoxStyle(this.tStyle)
    If this.tMovable Then
      .Placement = xlMoveAndSize
    Else: .Placement = xlFreeFloating
    End If
    .value = this.tValue
    .Enabled = this.tEnabled
    If this.tLinkedCell <> vbNullString Then .LinkedCell = this.tLinkedCell
    .ShapeRange.AlternativeText = left$(this.tMark, 1) & CStr(this.tTabIndex)
    .Caption = this.tCaption
    .Visible = this.tVisible
    If this.TName <> vbNullString Then .Name = this.TName
  End With
  CheckBoxShape.Cut
  this.tSheet.Paste
  Set CheckBoxShape = this.tSheet.CheckBoxes(Selection.Name)
  With CheckBoxShape
    .left = this.tLeft
    .top = this.tTop
  End With
End Sub

Private Sub Class_Initialize()
  this.tID = Hash(TypeName(Me) & CLng(Rnd * (2 ^ 31)))
End Sub
