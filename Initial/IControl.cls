VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "IControl"

Public Enum Color
  MainText = 7237230
  MainBackground = 12763848  '15790320 '15132391
  TitleText = 49264
  OnBackground = 13158600
  OffBackground = 6705232
  SelectBackground = 9359785
  OnText = 16448250
  OffText = 16777215
  CancelText = 3367935
  SelectLine = 4697456
  CancelLine = 3243501
  TableTitleBackground = 9730168
  AccentText = 12611584
End Enum

Public Enum FontSizeEnum
  Standard = 10
  Small = 9
  Medium = 11
  MediumLarge = 13
  TitleSmall = 16
  TitleStandard = 18
  TitleLarge = 25
End Enum

Public Enum ControlStyle
  csNone
  csOff
  csOn
  csSelected
  csCanceled
  csSimplified
  csAccented
  csUnderlined
  'ctBool
  csGlowed
  csDefault
  csAlternative
End Enum

Public Enum size
  Sml = 12
  std = 14
  Mdl = 16
  Dbl = 34
  Trp = 45
End Enum

Public Enum bType
  typeOn = 1
  typeOff = 2
End Enum

Public Property Get ID() As String
End Property

Public Property Get ControlType() As String
End Property

Public Property Get Name() As String
End Property

Public Property Let Name(ByVal text As String)
End Property

Public Property Get Mark() As String
End Property

Public Property Get TabIndex() As Long
End Property

Public Property Let TabIndex(ByVal value As Long)
End Property

Public Property Get Visible() As Boolean
End Property

Public Property Let Visible(ByVal value As Boolean)
End Property

Public Property Let Style(ByVal value As ControlStyle)
End Property

Public Property Let Movable(ByVal state As Boolean)
End Property

Public Sub Create(ByVal wsh As Worksheet)
End Sub

Public Property Get Height() As Single
End Property

Public Property Get Width() As Single
End Property

Public Property Get top() As Single
End Property

Public Property Get left() As Single
End Property

Public Property Get Bottom() As Single
End Property

Public Property Get Right() As Single
End Property

Public Property Get BottomRightRange() As Range
End Property

Public Property Get BottomLeftRange() As Range
End Property

Public Property Let Caption(ByVal value As String)
End Property

Private Sub Class_Initialize()
  Err.Raise 5, ClassName, Err.Source & vbCrLf & "- Interface class must not be instantiated"
End Sub
