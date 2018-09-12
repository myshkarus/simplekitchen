VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "ButtonBuilder"

Private Type TAttribute
  TName As String
  tVisible As Boolean
  tCaption As String
  tStyle As ControlStyle
  tLeft As Single
  tTop As Single
  tWidth As Single
  tHeight As Single
  tTabIndex As Long
  tIcon As String
  tOnAction As String
  tMark As String
  tAlignment As MsoParagraphAlignment
  tFontSize As Long
  tMovable As Boolean
End Type

Private this As TAttribute
Private thisReset As TAttribute
Private mSheet As Worksheet

Public Function Self() As ButtonBuilder
  Set Self = Me
End Function

Private Function myButton() As IControl
  Dim btn As IControl
  With New CButton
    .Caption = this.tCaption
    .CLeft = this.tLeft
    .top = this.tTop
    .Height = this.tHeight
    .Width = this.tWidth
    .Alignment = this.tAlignment
    .Mark = this.tMark
    .FontSize = this.tFontSize
    .Icon = this.tIcon
    .OnAction = this.tOnAction
    Set btn = .Self
  End With
  With btn
    .Name = this.TName
    .Style = this.tStyle
    .Movable = this.tMovable
    .TabIndex = this.tTabIndex
    .Visible = this.tVisible
    .Create mSheet
  End With
  Set myButton = btn
  this = thisReset
  Set btn = Nothing
End Function

Public Function Build() As IControl
  Set Build = myButton
End Function

Public Function OfPlaceAndSize(Optional ByVal left As Single, Optional ByVal top As Single, Optional ByVal Width As Single, Optional ByVal Height As Single) As ButtonBuilder
  With this
    If left > 0 Then .tLeft = left
    If top > 0 Then .tTop = top
    If Width > 0 Then .tWidth = Width
    If Height > 0 Then .tHeight = Height
  End With
  Set OfPlaceAndSize = Self
End Function

Public Function Name(ByVal text As String) As ButtonBuilder
  If text = vbNullString Then Exit Function
  this.TName = text
  Set Name = Self
End Function

Public Function SetIcon(ByVal Icon As String) As ButtonBuilder
  If Icon = vbNullString Then Exit Function
  this.tIcon = Icon
  Set SetIcon = Self
End Function

Public Function Visible(ByVal value As Boolean) As ButtonBuilder
  this.tVisible = value
  Set Visible = Self
End Function

Public Function Style(ByVal value As ControlStyle) As ButtonBuilder
  this.tStyle = value
  Set Style = Self
End Function

Public Function Movable(ByVal status As Boolean) As ButtonBuilder
  this.tMovable = status
  Set Movable = Self
End Function

Public Function Caption(ByVal text As String) As ButtonBuilder
  If text = vbNullString Then Exit Function
  this.tCaption = text
  Set Caption = Self
End Function

Public Function FontSize(ByVal value As Variant) As ButtonBuilder
  this.tFontSize = value
  Set FontSize = Self
End Function

Public Function OnAction(ByVal macros As String) As ButtonBuilder
  this.tOnAction = macros
  Set OnAction = Self
End Function

Public Function Alignment(ByVal val As MsoParagraphAlignment) As ButtonBuilder
  this.tAlignment = val
  Set Alignment = Self
End Function

Public Function Mark(ByVal val As String) As ButtonBuilder
  this.tMark = val
  Set Mark = Self
End Function

Public Function ToString() As String
  ToString = ClassName & " " & "to be build"
End Function

Public Property Get sheet() As Worksheet
  Set sheet = mSheet
End Property

Friend Property Set sheet(ByVal wsh As Worksheet)
  If Not wsh Is Nothing Then Set mSheet = wsh
End Property

Private Sub Class_Terminate()
  Set mSheet = Nothing
End Sub
