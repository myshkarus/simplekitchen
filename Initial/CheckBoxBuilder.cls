VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "CheckBoxBuilder"

Private Type TAttribute
  TName As String
  tVisible As Boolean
  tCaption As String
  tStyle As ControlStyle
  tLeft As Single
  tTop As Single
  tWidth As Single
  tHeight As Single
  'tID As String
  tTabIndex As Long
  tValue As Boolean
  tLinkedCell As String
  tEnabled As Boolean
  tMark As String
  tMovable As Boolean
End Type

Private this As TAttribute
Private thisReset As TAttribute
Private mSheet As Worksheet

Public Function Self() As CheckBoxBuilder
  Set Self = Me
End Function

Private Function myCheckBox() As IControl
  Dim chk As IControl
  With New CCheckBox
    .Caption = this.tCaption
    .CLeft = this.tLeft
    .top = this.tTop
    .Height = this.tHeight
    .Width = this.tWidth
    .Enabled = this.tEnabled
    .LinkedCell = this.tLinkedCell
    .Mark = this.tMark
    Set chk = .Self
  End With
  With chk
    .Name = this.TName
    .Style = this.tStyle
    .Movable = this.tMovable
    .TabIndex = this.tTabIndex
    .Visible = this.tVisible
    .Create mSheet
  End With
  Set myCheckBox = chk
  this = thisReset
  Set chk = Nothing
End Function

Public Function Build() As IControl
  Set Build = myCheckBox
End Function

Public Function Mark(ByVal val As String) As CheckBoxBuilder
  this.tMark = val
  Set Mark = Self
End Function

Public Function OfPlaceAndSize(Optional ByVal left As Single, Optional ByVal top As Single, Optional ByVal Width As Single, Optional ByVal Height As Single) As CheckBoxBuilder
  With this
    If left > 0 Then .tLeft = left
    If top > 0 Then .tTop = top
    If Width > 0 Then .tWidth = Width
    If Height > 0 Then .tHeight = Height
  End With
  Set OfPlaceAndSize = Self
End Function

Public Function Name(ByVal text As String) As CheckBoxBuilder
  If text = vbNullString Then Exit Function
  this.TName = text
  Set Name = Self
End Function

Public Function value(ByVal val As Boolean) As CheckBoxBuilder
  this.tValue = val
  Set value = Self
End Function

Public Function Enabled(ByVal val As Boolean) As CheckBoxBuilder
  this.tEnabled = val
  Set Enabled = Self
End Function

Public Function Visible(ByVal value As Boolean) As CheckBoxBuilder
  this.tVisible = value
  Set Visible = Self
End Function

Public Function Style(ByVal value As ControlStyle) As CheckBoxBuilder
  this.tStyle = value
  Set Style = Self
End Function

Public Function Movable(ByVal status As Boolean) As CheckBoxBuilder
  this.tMovable = status
  Set Movable = Self
End Function

Public Function Caption(ByVal text As String) As CheckBoxBuilder
  If text = vbNullString Then Exit Function
  this.tCaption = text
  Set Caption = Self
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
